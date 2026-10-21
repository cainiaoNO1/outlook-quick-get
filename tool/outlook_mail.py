#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook 取件工具 (Python 版)  v1.0
============================================================
用 refresh_token 换 Access Token -> IMAP XOAUTH2 登录 outlook.office365.com
-> 拉取最新邮件并自动提取验证码。

依赖: Python 3.8+ 以及 requests  (pip install requests)

用法:
    python outlook_mail.py                  # 收件箱最新 5 封
    python outlook_mail.py -n 10            # 最新 10 封
    python outlook_mail.py --code           # 只输出验证码
    python outlook_mail.py -m "Junk"        # 看垃圾邮件
    python outlook_mail.py -w 15 --exit-on-code   # 守验证码

【重要】微软 IMAP 要求令牌的 audience 是 outlook.office.com。
不指定 scope 刷新出来的令牌通常属于 Graph 等资源, 登录会报
"AUTHENTICATE failed."。本脚本默认带上 IMAP scope, 失败会自动回落。
"""
from __future__ import annotations

import argparse
import email
import imaplib
import json
import os
import re
import sys
import time
from email.header import decode_header
from email.utils import parsedate_to_datetime
from typing import Any, Dict, List, Optional, Tuple

try:
    import requests
except ImportError:  # pragma: no cover
    sys.exit("缺少 requests 库, 请先执行: pip install requests")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ACCOUNT_FILE = os.path.join(BASE_DIR, "account.txt")
TOKEN_CACHE = os.path.join(BASE_DIR, ".token_cache.json")

IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993
TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
IMAP_SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"

KEYWORDS = (
    "验证码|校验码|动态码|确认码|安全码|激活码|登录码|一次性密码|密码|"
    "verification code|verification|security code|login code|code|passcode|password|otp|pin"
)


def log(*a: Any) -> None:
    print(*a, file=sys.stderr, flush=True)


# ------------------------------------------------------------------ #
# 账号文件                                                            #
# ------------------------------------------------------------------ #
class Account:
    def __init__(self, email_addr: str, password: str, client_id: str,
                 refresh_token: str, path: str, raw: str) -> None:
        self.email = email_addr
        self.password = password
        self.client_id = client_id
        self.refresh_token = refresh_token
        self.path = path
        self.raw = raw


def load_account(path: Optional[str]) -> Account:
    target = os.path.abspath(path) if path else ACCOUNT_FILE
    if not os.path.exists(target):
        sys.exit(f"找不到账号文件: {target}\n"
                 f'请写入一行: 邮箱----密码----client_id----refresh_token')
    with open(target, "r", encoding="utf-8") as f:
        raw = f.read()
    squashed = re.sub(r"\s+", "", raw)
    parts = squashed.split("----")
    if len(parts) != 4:
        sys.exit(f"账号文件格式不对: 需要 4 段(用 ---- 分隔), 实际 {len(parts)} 段\n文件: {target}")
    email_addr, password, client_id, refresh_token = parts
    if "@" not in email_addr:
        sys.exit(f"邮箱地址看起来不对: {email_addr}")
    if not refresh_token:
        sys.exit("refresh_token 为空")
    return Account(email_addr, password, client_id, refresh_token, target, raw)


def persist_refresh_token(acct: Account, new_token: str) -> None:
    """微软会轮换 refresh_token, 不写回的话下次可能就用不了。"""
    try:
        if "#" in acct.raw:
            log("[提示] account.txt 含注释行, 为免破坏内容不自动更新 refresh_token")
            return
        with open(acct.path, "w", encoding="utf-8") as f:
            f.write(f"{acct.email}----{acct.password}----{acct.client_id}----{new_token}\n")
        acct.refresh_token = new_token
        log("[提示] 微软已轮换 refresh_token, 已自动写回 account.txt")
    except OSError as e:
        log(f"[提示] 写回 account.txt 失败(不影响本次使用): {e}")


def read_json(path: str) -> Dict[str, Any]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def write_json(path: str, data: Dict[str, Any]) -> None:
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


# ------------------------------------------------------------------ #
# OAuth2                                                              #
# ------------------------------------------------------------------ #
def describe_error(status: int, body: Dict[str, Any], text: str) -> str:
    code = str(body.get("error") or "")
    desc = str(body.get("error_description") or "").split("\n")[0] or text[:400]
    hint = ""
    if "invalid_grant" in code + desc:
        hint = ("\n>>> refresh_token 已失效/被吊销: 多半是被轮换后没保存、改过密码、或长期未使用。"
                "需要重新获取邮箱密钥。")
    elif "invalid_client" in code + desc:
        hint = "\n>>> client_id 与这份 refresh_token 不匹配。"
    elif re.search(r"AADSTS90023|No applicable permissions|unauthorized_client|consent", code + desc):
        hint = "\n>>> 该账号/应用没有邮箱权限(缺少 IMAP 授权)。"
    return f"换取 Access Token 失败 (HTTP {status}) {code}: {desc}{hint}"


def get_access_token(acct: Account, no_cache: bool, scope: Optional[str], no_save: bool) -> str:
    cache = read_json(TOKEN_CACHE)
    if (not no_cache and cache.get("email") == acct.email and cache.get("access_token")
            and float(cache.get("expires_at", 0)) > time.time() * 1000 + 120000):
        left = int((float(cache["expires_at"]) - time.time() * 1000) / 60000)
        log(f"[1/3] 复用缓存的 Access Token (约剩 {left} 分钟)")
        return cache["access_token"]

    attempts: List[Tuple[Optional[str], str]] = (
        [(scope, "自定义 scope")] if scope else [(IMAP_SCOPE, "IMAP scope"), (None, "原授权 scope")]
    )
    last_err = "换取 Access Token 失败"
    for idx, (sc, label) in enumerate(attempts):
        log(f"[1/3] 正在用 refresh_token 换取 Access Token ({label}) ...")
        payload = {
            "client_id": acct.client_id,
            "refresh_token": acct.refresh_token,
            "grant_type": "refresh_token",
        }
        if sc:
            payload["scope"] = sc
        try:
            resp = requests.post(TOKEN_URL, data=payload, timeout=30)
        except Exception as e:
            sys.exit(f"请求 token 端点失败: {e}")

        try:
            body = resp.json()
        except ValueError:
            body = {}

        if resp.status_code == 200 and body.get("access_token"):
            if body.get("refresh_token") and body["refresh_token"] != acct.refresh_token and not no_save:
                persist_refresh_token(acct, body["refresh_token"])
            ttl = int(body.get("expires_in") or 3600) * 1000
            write_json(TOKEN_CACHE, {
                "email": acct.email,
                "scope": sc or "",
                "access_token": body["access_token"],
                "expires_at": time.time() * 1000 + ttl - 60000,
                "saved_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
            })
            log(f"[1/3] 拿到 Access Token (有效期 {ttl // 60000} 分钟)")
            return body["access_token"]

        last_err = describe_error(resp.status_code, body, resp.text)
        if re.search(r"invalid_grant|invalid_client", last_err):
            sys.exit(last_err)
        if idx < len(attempts) - 1:
            log("[1/3] 该 scope 不适用, 换一种方式重试 ...")

    sys.exit(last_err)


# ------------------------------------------------------------------ #
# 解析 / 提取                                                         #
# ------------------------------------------------------------------ #
def decode_str(value: Optional[str]) -> str:
    """解码 RFC2047 标题/发件人。"""
    if not value:
        return ""
    chunks: List[str] = []
    for text, charset in decode_header(value):
        if isinstance(text, bytes):
            chunks.append(text.decode(charset or "utf-8", errors="ignore"))
        else:
            chunks.append(text)
    return re.sub(r"\s+", " ", "".join(chunks)).strip()


def render_body(msg: email.message.Message) -> List[str]:
    texts: List[str] = []
    plain: List[str] = []
    html: List[str] = []
    for part in msg.walk():
        ctype = part.get_content_type()
        if ctype not in ("text/plain", "text/html"):
            continue
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        charset = part.get_content_charset() or "utf-8"
        try:
            text = payload.decode(charset, errors="ignore")
        except LookupError:
            text = payload.decode("utf-8", errors="ignore")
        if ctype == "text/plain":
            plain.append(text)
        else:
            html.append(html_to_text(text))
    if plain:
        return plain
    return html


def html_to_text(html: str) -> str:
    s = re.sub(r"<(script|style)[\s\S]*?</\1>", " ", html, flags=re.I)
    s = re.sub(r"<br\s*/?>", "\n", s, flags=re.I)
    s = re.sub(r"</(p|div|tr|h[1-6]|li|table)>", "\n", s, flags=re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = (s.replace("&nbsp;", " ").replace("&lt;", "<").replace("&gt;", ">")
         .replace("&quot;", '"').replace("&amp;", "&"))
    s = re.sub(r"[ \t\u00a0]+", " ", s)
    return "\n".join(line.strip() for line in s.splitlines() if line.strip())


def extract_codes(subject: str, text: str) -> List[str]:
    hay = f"{subject}\n{text}"
    scores: Dict[str, int] = {}

    def add(code: str, score: int) -> None:
        if not code:
            return
        if re.fullmatch(r"(19|20)\d{2}", code):  # 年份基本不可能是验证码
            return
        if code not in scores or scores[code] < score:
            scores[code] = score

    # 1) 关键词附近的纯数字验证码
    for m in re.finditer(rf"(?:{KEYWORDS})[^0-9A-Za-z]{{0,30}}([0-9]{{4,8}})", hay, re.I):
        add(m.group(1), 100)
    # 2) 关键词附近的字母数字混合验证码(必须含数字)
    for m in re.finditer(rf"(?:{KEYWORDS})[^0-9A-Za-z]{{0,30}}([A-Z0-9]{{5,8}})", hay, re.I):
        code = m.group(1)
        if re.search(r"\d", code):
            add(code, 90)
    # 3) 独立成词的 6 位数字
    for m in re.finditer(r"(?<![\w.,/:-])(\d{6})(?![\w.,/-])", hay):
        add(m.group(1), 60)
    # 4) 独立成词的 5/7/8 位数字
    for m in re.finditer(r"(?<![\w.,/:-])(\d{5,8})(?![\w.,/-])", hay):
        add(m.group(1), 35)

    ranked = sorted(((s, c) for c, s in scores.items() if s >= 50),
                    key=lambda x: (-x[0], len(x[1])))
    return [c for _, c in ranked[:5]]


def extract_links(text: str) -> List[str]:
    out: List[str] = []
    for m in re.finditer(r"https?://[^\s\"'<>)\]}]+", text or "", re.I):
        url = m.group(0).rstrip(".,;")
        if url not in out:
            out.append(url)
        if len(out) >= 3:
            break
    return out


def format_mail(raw_bytes: bytes, uid: str, folder: str = "", snip: bool = True) -> Dict[str, Any]:
    msg = email.message_from_bytes(raw_bytes)
    subject = decode_str(msg.get("Subject"))
    sender = decode_str(msg.get("From"))
    date_raw = msg.get("Date") or ""
    parsed = None
    try:
        parsed = parsedate_to_datetime(date_raw) if date_raw else None
    except Exception:
        parsed = None
    date = parsed.strftime("%Y-%m-%d %H:%M:%S") if parsed else date_raw

    body = "\n".join(render_body(msg)).strip()
    codes = extract_codes(subject, body)
    links = extract_links(body)
    if snip and len(body) > 500:
        body = body[:500] + " …"
    return {"folder": folder, "uid": uid, "from": sender, "subject": subject, "date": date,
            "body": body, "codes": codes, "links": links,
            "ts": parsed.timestamp() if parsed else 0.0}


# ------------------------------------------------------------------ #
# IMAP                                                                #
# ------------------------------------------------------------------ #
def imap_login(acct: Account, token: str) -> imaplib.IMAP4_SSL:
    log(f"[2/3] 连接 {IMAP_HOST}:{IMAP_PORT} 并做 XOAUTH2 认证 ...")
    try:
        conn = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    except Exception as e:
        sys.exit(f"连接 IMAP 服务器失败: {e}")
    auth_string = f"user={acct.email}\x01auth=Bearer {token}\x01\x01"
    try:
        conn.authenticate("XOAUTH2", lambda _x: auth_string.encode("utf-8"))
    except imaplib.IMAP4.error as e:
        sys.exit(f"XOAUTH2 认证失败: {e}\n"
                 f">>> 通常是令牌没有 IMAP 权限(scope 不对) 或 refresh_token 已失效。")
    return conn


def uid_list(conn: imaplib.IMAP4_SSL, criteria: str) -> List[int]:
    typ, data = conn.uid("SEARCH", None, criteria)
    if typ != "OK":
        raise RuntimeError(f"SEARCH 失败: {data}")
    ids: List[int] = []
    for chunk in data:
        if chunk:
            ids.extend(int(x) for x in chunk.split() if x.isdigit())
    return ids


def fetch_one(conn: imaplib.IMAP4_SSL, uid: int) -> Optional[bytes]:
    typ, data = conn.uid("FETCH", str(uid), "(BODY.PEEK[])")
    if typ != "OK":
        return None
    for item in data:
        if isinstance(item, tuple) and len(item) > 1:
            return item[1]
    return None


def print_mail(m: Dict[str, Any], index: Any) -> None:
    bar = "─" * 58
    print(f"【{index}】[{m.get('folder') or '-'}] UID {m['uid']}   {m['date']}")
    print(f"发件人 : {m['from'] or '(未知)'}")
    print(f"主  题 : {m['subject'] or '(无主题)'}")
    print(f"验证码 : {('  /  '.join(m['codes'])) if m['codes'] else '(未识别)'}")
    if m["links"]:
        print(f"链  接 : {m['links'][0]}")
    print("正  文 :")
    print(m["body"] or "(空)")
    print(bar)


DEFAULT_SCAN = ["INBOX", "Junk"]


def resolve_folders(args: argparse.Namespace) -> List[str]:
    """验证码邮件常被判成垃圾邮件, 默认两个文件夹一起扫。"""
    if args.mailbox:
        return [args.mailbox]
    if args.scan:
        return [s.strip() for s in args.scan.split(",") if s.strip()]
    return list(DEFAULT_SCAN)


def run_once(conn: imaplib.IMAP4_SSL, acct: Account, args: argparse.Namespace) -> None:
    folders = resolve_folders(args)
    stats: List[Tuple[str, int]] = []
    mails: List[Dict[str, Any]] = []
    total = 0

    for name in folders:
        try:
            typ, data = conn.select(f'"{name}"')
            if typ != "OK":
                raise RuntimeError(str(data))
        except Exception as e:
            if len(folders) > 1:
                log(f'[警告] 打不开文件夹 "{name}", 跳过: {e}')
                continue
            sys.exit(f'打开文件夹 "{name}" 失败: {e}  (可用 --folders 查名称)')

        ids = uid_list(conn, "UNSEEN" if args.unseen else "ALL")
        stats.append((name, len(ids)))
        total += len(ids)
        # FETCH 依赖当前 SELECT 的文件夹, 必须在这个循环里取完
        for uid in ids[-args.limit:]:
            raw = fetch_one(conn, uid)
            if raw:
                mails.append(format_mail(raw, str(uid), name, snip=not args.raw))

    if not mails:
        log(f"[3/3] {' + '.join(folders)} 里没有邮件")
        if args.json:
            print(json.dumps({"account": acct.email, "mailbox": "+".join(folders),
                              "folders": [{"name": n, "count": c} for n, c in stats],
                              "total": 0, "mails": []}, ensure_ascii=False, indent=2))
        return

    mails.sort(key=lambda m: m["ts"], reverse=True)
    top = mails[:args.limit]
    log("[3/3] " + " + ".join(f"{n} {c} 封" for n, c in stats) + f"  → 读取最新 {len(top)} 封")

    if args.json:
        for m in top:
            m.pop("ts", None)
        print(json.dumps({"account": acct.email, "mailbox": "+".join(folders),
                          "folders": [{"name": n, "count": c} for n, c in stats],
                          "total": total, "mails": top}, ensure_ascii=False, indent=2))
        return
    if args.code:
        out = [m["codes"][0] if m["codes"] else f"LINK:{m['links'][0]}"
               for m in top if m["codes"] or m["links"]]
        if out:
            print("\n".join(out))
        return

    summary = " + ".join(f"{n} {c} 封" for n, c in stats if c) or " + ".join(folders)
    print("\n" + "=" * 58)
    print(f"账号   : {acct.email}")
    print(f"文件夹 : {summary}   (共 {total} 封)")
    print("=" * 58 + "\n")
    for i, m in enumerate(top, 1):
        print_mail(m, i)


def run_watch(conn: imaplib.IMAP4_SSL, acct: Account, args: argparse.Namespace) -> None:
    folders = resolve_folders(args)
    last: Dict[str, int] = {}
    for name in folders:
        try:
            conn.select(f'"{name}"')
            ids = uid_list(conn, "ALL")
            last[name] = max(ids) if ids else 0
        except Exception as e:
            log(f'[警告] 打不开文件夹 "{name}": {e}')
            last[name] = 0

    base = ", ".join(f"{n}={u}" for n, u in last.items())
    log(f"[3/3] 开始轮询 {' + '.join(folders)} (当前 {base}, 间隔 {args.watch}s, Ctrl+C 退出)")

    while True:
        time.sleep(args.watch)
        got_code = False
        for name in folders:
            since = last.get(name, 0)
            try:
                conn.select(f'"{name}"')
                fresh = sorted(u for u in uid_list(conn, f"UID {since + 1}:*") if u > since)
            except Exception as e:
                log(f'[错误] 轮询 "{name}" 失败: {e}')
                continue
            if not fresh:
                continue
            last[name] = max(since, max(fresh))

            for uid in fresh:
                raw = fetch_one(conn, uid)
                if not raw:
                    continue
                m = format_mail(raw, str(uid), name, snip=not args.raw)
                print(f"\n[{time.strftime('%H:%M:%S')}] 新邮件!  ({name})")
                print_mail(m, "新")
                if m["codes"]:
                    got_code = True
                    log(">>> 验证码: " + " / ".join(m["codes"]))
        if got_code and args.exit_on_code:
            log("已取到验证码, 按 --exit-on-code 退出。")
            return


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="outlook_mail.py",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Outlook 取件工具: 拉取邮件并提取验证码",
    )
    p.add_argument("-n", "--limit", type=int, default=5, help="查看最新 N 封 (默认 5)")
    p.add_argument("-m", "--mailbox", default=None,
                   help="只查指定文件夹 (默认同时扫 INBOX + Junk)")
    p.add_argument("--scan", default=None,
                   help='自定义要扫的文件夹, 逗号分隔 (默认 "INBOX,Junk"')
    p.add_argument("-a", "--account", default=None, help="账号文件路径 (默认同目录 account.txt)")
    p.add_argument("-u", "--unseen", action="store_true", help="只看未读")
    p.add_argument("--code", action="store_true", help="只输出验证码")
    p.add_argument("--json", action="store_true", help="输出 JSON")
    p.add_argument("--raw", action="store_true", help="正文不截断")
    p.add_argument("-w", "--watch", type=int, default=0, nargs="?",
                   const=15, help="轮询新邮件间隔秒数 (默认 15)")
    p.add_argument("--exit-on-code", action="store_true", help="轮询时取到验证码即退出")
    p.add_argument("--folders", action="store_true", help="列出所有文件夹")
    p.add_argument("--no-cache", action="store_true", help="忽略 Access Token 缓存")
    p.add_argument("--no-save", action="store_true", help="不写回轮换后的 refresh_token")
    p.add_argument("--scope", default=None, help="自定义 scope (一般不需要)")
    return p


def main() -> None:
    args = build_parser().parse_args()
    acct = load_account(args.account)
    token = get_access_token(acct, args.no_cache, args.scope, args.no_save)
    conn = imap_login(acct, token)
    try:
        if args.folders:
            typ, data = conn.list()
            for line in data or []:
                if isinstance(line, bytes):
                    print("  " + line.decode("utf-8", errors="ignore"))
            return
        if args.watch and args.watch > 0:
            run_watch(conn, acct, args)
        else:
            run_once(conn, acct, args)
    finally:
        try:
            conn.logout()
        except Exception:
            pass


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n已中断", file=sys.stderr)
        sys.exit(130)
