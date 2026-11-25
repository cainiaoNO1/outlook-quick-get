#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook 快速取件 - 后端服务
============================================================
把桌面版「Outlook 取件工具」的取件能力（refresh_token -> OAuth2 ->
IMAP XOAUTH2 / Microsoft Graph -> 邮件 + 验证码识别）封装成 HTTP 接口，
供同源网页调用。

接口:
    POST /api/security-session   下发 AES-GCM/HMAC 会话密钥
    POST /api/fetch-imap         IMAP 取件（收件箱 + 垃圾邮件）
    POST /api/fetch-graph        Microsoft Graph 取件
    POST /api/refresh-token      刷新 refresh_token
    GET  /api/health             健康检查
    其余 GET                     静态文件

仅依赖 Python 标准库 + cryptography（AES-GCM），无需 pip 安装其它包。
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import imaplib
import json
import mimetypes
import os
import re
import secrets
import shutil
import ssl
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from email.header import decode_header
from email.utils import getaddresses, parsedate_to_datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
except ImportError:  # pragma: no cover
    sys.stderr.write("缺少 cryptography 库: apt install python3-cryptography\n")
    raise

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WWW_DIR = os.path.join(BASE_DIR, "www")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993

IMAP_SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"
GRAPH_SCOPE = "https://graph.microsoft.com/Mail.Read offline_access"

DEFAULT_FOLDERS = ["INBOX", "Junk"]
MAX_BODY_CHARS = 200_000

# ------------------------------------------------------------------ #
# 配置                                                                #
# ------------------------------------------------------------------ #

DEFAULT_CONFIG = {
    "host": "127.0.0.1",
    "port": 8090,
    "default_client_id": "",
    "imap_fetch_cap": 20,
    "request_timeout": 60,
    "use_tool": True,
    "tool_dir": "",
    "tool_timeout": 90,
}


def load_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg.update(json.load(f) or {})
        except Exception as exc:  # pragma: no cover
            log(f"[warn] 读取 config.json 失败: {exc}")
    for key in ("DSH_MAIL_PORT", "PORT"):
        if os.environ.get(key):
            try:
                cfg["port"] = int(os.environ[key])
            except ValueError:
                pass
    if os.environ.get("DSH_MAIL_HOST"):
        cfg["host"] = os.environ["DSH_MAIL_HOST"]
    return cfg


CONFIG = load_config()


def log(*args) -> None:
    sys.stderr.write(" ".join(str(a) for a in args) + "\n")
    sys.stderr.flush()


# ------------------------------------------------------------------ #
# 安全会话（AES-GCM + HMAC 信封）                                      #
# ------------------------------------------------------------------ #

SESSION_TTL_SECONDS = 30 * 60
CLOCK_SKEW_MS = 5 * 60 * 1000
_sessions: dict[str, dict] = {}
_sessions_lock = threading.Lock()


def _prune_sessions(now: float) -> None:
    dead = [sid for sid, s in _sessions.items() if s["expires_at"] <= now]
    for sid in dead:
        _sessions.pop(sid, None)


def create_session() -> dict:
    now = time.time()
    with _sessions_lock:
        _prune_sessions(now)
        session_id = secrets.token_urlsafe(18)
        key = secrets.token_bytes(32)
        session = {
            "session_id": session_id,
            "token": secrets.token_urlsafe(24),
            "key": key,
            "expires_at": now + SESSION_TTL_SECONDS,
            "nonces": set(),
        }
        _sessions[session_id] = session
    return {
        "success": True,
        "sessionId": session["session_id"],
        "sessionToken": session["token"],
        "sessionKey": b64url(key),
        "expiresAt": datetime.fromtimestamp(session["expires_at"], timezone.utc).isoformat(),
        "ttl": SESSION_TTL_SECONDS,
    }


def b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode().rstrip("=")


def b64url_decode(value: str) -> bytes:
    pad = "=" * (-len(value) % 4)
    return base64.urlsafe_b64decode(value + pad)


class EnvelopeError(Exception):
    def __init__(self, message: str, code: str = "SECURITY_ENVELOPE_INVALID"):
        super().__init__(message)
        self.code = code


def open_envelope(body: dict) -> dict:
    """校验并解密前端信封；未加密的请求体原样返回（兼容 curl 调试）。"""
    if not isinstance(body, dict) or not body.get("secure"):
        return body if isinstance(body, dict) else {}

    required = ("sessionId", "sessionToken", "nonce", "timestamp", "iv", "ciphertext", "signature")
    missing = [k for k in required if not body.get(k)]
    if missing:
        raise EnvelopeError(f"安全信封缺少字段: {', '.join(missing)}")

    now_ms = int(time.time() * 1000)
    try:
        ts = int(body["timestamp"])
    except (TypeError, ValueError):
        raise EnvelopeError("请求时间戳无效")
    if abs(now_ms - ts) > CLOCK_SKEW_MS:
        raise EnvelopeError("请求已过期，请刷新页面重试")

    with _sessions_lock:
        _prune_sessions(time.time())
        session = _sessions.get(body["sessionId"])
        if not session:
            raise EnvelopeError("安全会话不存在或已过期")
        if not hmac.compare_digest(session["token"], str(body["sessionToken"])):
            raise EnvelopeError("安全会话令牌不匹配")
        if body["nonce"] in session["nonces"]:
            raise EnvelopeError("重复的请求随机数")
        key = session["key"]

    signed = ".".join(str(body[k]) for k in ("sessionId", "nonce", "timestamp", "iv", "ciphertext"))
    expected = hmac.new(key, signed.encode(), hashlib.sha256).digest()
    try:
        provided = b64url_decode(str(body["signature"]))
    except Exception:
        raise EnvelopeError("请求签名格式无效")
    if not hmac.compare_digest(expected, provided):
        raise EnvelopeError("请求签名校验失败")

    try:
        iv = b64url_decode(str(body["iv"]))
        ciphertext = b64url_decode(str(body["ciphertext"]))
        plaintext = AESGCM(key).decrypt(iv, ciphertext, None)
    except Exception:
        raise EnvelopeError("加密请求体无法解密")

    with _sessions_lock:
        session = _sessions.get(body["sessionId"])
        if session:
            session["nonces"].add(body["nonce"])
            if len(session["nonces"]) > 5000:
                session["nonces"].clear()

    try:
        payload = json.loads(plaintext.decode("utf-8"))
    except Exception:
        raise EnvelopeError("加密请求体不是合法 JSON")
    if not isinstance(payload, dict):
        raise EnvelopeError("加密请求体格式错误")
    return payload


# ------------------------------------------------------------------ #
# OAuth2                                                              #
# ------------------------------------------------------------------ #

class TokenError(Exception):
    def __init__(self, message: str, code: str = "", reason: str = "", action: str = "", detail: str = ""):
        super().__init__(message)
        self.code = code
        self.reason = reason
        self.action = action
        self.detail = detail


_token_cache: dict[str, dict] = {}
_token_lock = threading.Lock()


def http_post_form(url: str, form: dict, timeout: int = 30) -> tuple[int, str]:
    data = urllib.parse.urlencode(form).encode()
    req = urllib.request.Request(
        url,
        data=data,
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
            "User-Agent": "MailGrabber/1.0",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, resp.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read().decode("utf-8", "replace")


def explain_token_error(status: int, payload: dict, text: str) -> TokenError:
    code = str(payload.get("error") or "")
    desc = str(payload.get("error_description") or "").split("\n")[0] or text[:400]
    blob = f"{code} {desc}"
    hint, action = "", ""
    if "invalid_grant" in blob:
        hint = "refresh_token 已失效或被吊销"
        action = "重新获取该邮箱的 refresh_token 后再导入"
    elif "invalid_client" in blob:
        hint = "client_id 与 refresh_token 不匹配"
        action = "检查导入的 clientid 是否属于同一应用"
    elif re.search(r"AADSTS90023|No applicable permissions|unauthorized_client|consent|AADSTS65001", blob):
        hint = "该应用没有所需的邮箱权限"
        action = "重新授权该应用，或改用其它协议取件"
    return TokenError(
        message=f"换取 Access Token 失败 (HTTP {status}) {code}: {desc}".strip(),
        code=code or "TOKEN_EXCHANGE_FAILED",
        reason=hint,
        action=action,
        detail=desc,
    )


def get_access_token(email: str, client_id: str, refresh_token: str, scope: str, cache_key: str) -> str:
    now = time.time()
    with _token_lock:
        cached = _token_cache.get(cache_key)
        if cached and cached["expires_at"] > now + 120:
            return cached["token"]

    status, text = http_post_form(TOKEN_URL, {
        "client_id": client_id,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
        "scope": scope,
    })
    payload = {}
    try:
        payload = json.loads(text)
    except Exception:
        payload = {}
    if status != 200 or not payload.get("access_token"):
        raise explain_token_error(status, payload, text)

    ttl = int(payload.get("expires_in") or 3600)
    with _token_lock:
        _token_cache[cache_key] = {
            "token": payload["access_token"],
            "expires_at": now + ttl - 60,
            "rotated_refresh_token": payload.get("refresh_token"),
        }
    return payload["access_token"]


def refresh_refresh_token(email: str, client_id: str, refresh_token: str) -> dict:
    status, text = http_post_form(TOKEN_URL, {
        "client_id": client_id,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
        "scope": IMAP_SCOPE,
    })
    payload = {}
    try:
        payload = json.loads(text)
    except Exception:
        payload = {}
    if status != 200 or not payload.get("access_token"):
        raise explain_token_error(status, payload, text)
    new_rt = payload.get("refresh_token") or refresh_token
    return {
        "success": True,
        "email": email,
        "refreshToken": new_rt,
        "newRefreshToken": new_rt,
        "changed": new_rt != refresh_token,
        "expiresIn": int(payload.get("expires_in") or 3600),
    }


# ------------------------------------------------------------------ #
# 极简 IMAP 客户端                                                     #
# ------------------------------------------------------------------ #

class ImapError(Exception):
    pass


class ImapClient:
    def __init__(self, host: str, port: int, timeout: int = 25, prewrapped=None):
        self.host = host
        self.port = port
        self.timeout = timeout
        self.buf = b""
        self.sock = prewrapped
        self.tag_seq = 0

    def connect(self) -> "ImapClient":
        if self.sock is None:
            ctx = ssl.create_default_context()
            self.sock = ctx.wrap_socket(
                __import__("socket").create_connection((self.host, self.port), self.timeout),
                server_hostname=self.host,
            )
        self.sock.settimeout(self.timeout)
        line = self._read_line()
        if not re.match(r"^\*\s+(OK|PREAUTH)", line, re.I):
            raise ImapError(f"IMAP 服务器问候异常: {line[:120]}")
        return self

    # ---------------- 底层读取 ----------------

    def _fill(self) -> None:
        chunk = self.sock.recv(65536)
        if not chunk:
            raise ImapError("IMAP 连接被服务器关闭")
        self.buf += chunk

    def _read_line(self) -> str:
        while b"\r\n" not in self.buf:
            self._fill()
        idx = self.buf.index(b"\r\n")
        line = self.buf[:idx].decode("latin-1")
        self.buf = self.buf[idx + 2:]
        return line

    def _read_exact(self, n: int) -> bytes:
        while len(self.buf) < n:
            self._fill()
        data = self.buf[:n]
        self.buf = self.buf[n:]
        return data

    def _read_response(self) -> dict:
        """读取一个完整响应（含 literal）。"""
        lines: list[str] = []
        literals: list[bytes] = []
        while True:
            line = self._read_line()
            lines.append(line)
            m = re.search(r"\{(\d+)\}$", line)
            if m:
                literals.append(self._read_exact(int(m.group(1))))
                continue
            break
        first = lines[0] if lines else ""
        if first.startswith("+"):
            kind = "continuation"
        elif first.startswith("*"):
            kind = "untagged"
        else:
            kind = "tagged"
        return {"lines": lines, "literals": literals, "kind": kind, "first": first}

    def command(self, cmd: str) -> tuple[str, list[dict]]:
        self.tag_seq += 1
        tag = f"T{self.tag_seq}"
        self.sock.sendall(f"{tag} {cmd}\r\n".encode())
        untagged: list[dict] = []
        while True:
            resp = self._read_response()
            if resp["kind"] == "continuation":
                raise ImapError("IMAP 协议异常: 命令阶段收到续行")
            if resp["kind"] == "tagged":
                m = re.match(rf"^{tag}\s+(OK|NO|BAD)\b\s*(.*)$", resp["first"], re.I)
                if not m:
                    continue
                if m.group(1).upper() != "OK":
                    raise ImapError(f"IMAP {m.group(1)}: {m.group(2)}")
                return m.group(2), untagged
            untagged.append(resp)

    # ---------------- 协议动作 ----------------

    def login(self, email: str, token: str) -> None:
        self.tag_seq += 1
        tag = f"T{self.tag_seq}"
        self.sock.sendall(f"{tag} AUTHENTICATE XOAUTH2\r\n".encode())
        resp = self._read_response()
        if resp["kind"] == "tagged":
            if re.match(rf"^{tag}\s+OK", resp["first"], re.I):
                return
            raise ImapError(f"XOAUTH2 认证失败: {resp['first'][:200]}")
        if resp["kind"] != "continuation":
            raise ImapError("XOAUTH2 认证失败: 服务器未要求续行")
        sasl = base64.b64encode(f"user={email}\x01auth=Bearer {token}\x01\x01".encode()).decode()
        self.sock.sendall(sasl.encode() + b"\r\n")
        while True:
            resp = self._read_response()
            if resp["kind"] == "continuation":
                detail = ""
                try:
                    raw = re.sub(r"^\+\s?", "", resp["first"])
                    detail = base64.b64decode(raw + "===").decode("utf-8", "replace")
                except Exception:
                    detail = ""
                self.sock.sendall(b"\r\n")
                raise ImapError(f"XOAUTH2 认证失败{': ' + detail if detail else ''}")
            if resp["kind"] == "tagged":
                m = re.match(rf"^{tag}\s+(OK|NO|BAD)\b\s*(.*)$", resp["first"], re.I)
                if not m:
                    continue
                if m.group(1).upper() == "OK":
                    return
                raise ImapError(f"XOAUTH2 认证失败: {m.group(2)[:200]}")

    def select(self, mailbox: str) -> int:
        _, untagged = self.command(f'SELECT "{mailbox}"')
        exists = 0
        for r in untagged:
            m = re.match(r"^\*\s+(\d+)\s+EXISTS", r["first"])
            if m:
                exists = int(m.group(1))
        return exists

    def plain_login(self, user: str, password: str) -> None:
        """LOGIN 明文认证（第三方邮箱，例如 Mail.com / GMX）。"""
        safe_user = user.replace("\\", "\\\\").replace('"', '\\"')
        safe_pass = password.replace("\\", "\\\\").replace('"', '\\"')
        self.command(f'LOGIN "{safe_user}" "{safe_pass}"')

    def uid_search(self, criteria: str = "ALL") -> list[int]:
        _, untagged = self.command(f"UID SEARCH {criteria}")
        ids: list[int] = []
        for r in untagged:
            if not re.match(r"^\*\s+SEARCH", r["first"], re.I):
                continue
            for part in r["first"].split()[2:]:
                if part.isdigit():
                    ids.append(int(part))
        return ids

    def uid_fetch(self, uid: int) -> bytes | None:
        _, untagged = self.command(f"UID FETCH {uid} (BODY.PEEK[])")
        for r in untagged:
            if r["literals"]:
                return r["literals"][0]
        return None

    def list_folders(self) -> list[str]:
        _, untagged = self.command('LIST "" "*"')
        out = []
        for r in untagged:
            m = re.match(r'^\*\s+LIST\s+\([^)]*\)\s+(?:"(?:[^"\\]|\\.)*"|NIL)\s+(.+)$', r["first"], re.I)
            if m:
                out.append(m.group(1).strip().strip('"'))
        return out

    def logout(self) -> None:
        try:
            self.command("LOGOUT")
        except Exception:
            pass
        try:
            if self.sock:
                self.sock.close()
        except Exception:
            pass


# ------------------------------------------------------------------ #
# MIME 解析                                                            #
# ------------------------------------------------------------------ #

CHARSET_ALIAS = {
    "utf8": "utf-8", "utf-8": "utf-8", "gb2312": "gbk", "gbk": "gbk", "gb18030": "gb18030",
    "cp936": "gbk", "big5": "big5", "iso-8859-1": "latin-1", "latin1": "latin-1",
    "us-ascii": "utf-8", "ascii": "utf-8", "windows-1252": "cp1252", "iso-2022-jp": "iso2022_jp",
    "shift_jis": "shift_jis", "koi8-r": "koi8_r",
}


def decode_bytes(data: bytes, charset: str | None) -> str:
    cs = str(charset or "utf-8").lower().split("*")[0].strip().strip("\"'")
    cs = CHARSET_ALIAS.get(cs, cs)
    for candidate in (cs, "utf-8", "gbk", "latin-1"):
        try:
            return data.decode(candidate)
        except (LookupError, UnicodeDecodeError):
            continue
    return data.decode("utf-8", "replace")


def decode_rfc2047(raw: str | None) -> str:
    if not raw:
        return ""
    chunks: list[str] = []
    for text, charset in decode_header(raw):
        if isinstance(text, bytes):
            chunks.append(decode_bytes(text, charset))
        else:
            chunks.append(text)
    return re.sub(r"\s+", " ", "".join(chunks)).strip()


def html_to_text(html: str) -> str:
    s = re.sub(r"<(script|style)[\s\S]*?</\1>", " ", html, flags=re.I)
    s = re.sub(r"<br\s*/?>", "\n", s, flags=re.I)
    s = re.sub(r"</(p|div|tr|h[1-6]|li|table)>", "\n", s, flags=re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = (s.replace("&nbsp;", " ").replace("&lt;", "<").replace("&gt;", ">")
         .replace("&quot;", '"').replace("&#39;", "'").replace("&amp;", "&"))
    s = re.sub(r"[ \t\u00a0]+", " ", s)
    return "\n".join(line.strip() for line in s.splitlines() if line.strip())


def split_params(value: str) -> list[str]:
    out, cur, quoted = [], "", False
    for ch in value:
        if ch == '"':
            quoted = not quoted
            cur += ch
            continue
        if ch == ";" and not quoted:
            out.append(cur)
            cur = ""
            continue
        cur += ch
    out.append(cur)
    return out


def parse_content_type(value: str) -> tuple[str, dict]:
    parts = split_params(value or "")
    main = (parts[0] or "").strip().lower()
    params: dict[str, str] = {}
    for p in parts[1:]:
        if "=" not in p:
            continue
        k, v = p.split("=", 1)
        params[k.strip().lower()] = v.strip().strip("\"'")
    return main, params


def parse_headers(head_text: str) -> dict[str, str]:
    headers: dict[str, str] = {}
    key = None
    for line in head_text.split("\r\n"):
        if line[:1] in (" ", "\t") and key:
            headers[key] += " " + line.strip()
            continue
        if ":" not in line:
            continue
        key = line.split(":", 1)[0].strip().lower()
        headers[key] = line.split(":", 1)[1].strip()
    return headers


def split_entity(data: bytes) -> tuple[dict[str, str], bytes]:
    for sep in (b"\r\n\r\n", b"\n\n"):
        idx = data.find(sep)
        if idx >= 0:
            return parse_headers(data[:idx].decode("latin-1")), data[idx + len(sep):]
    return parse_headers(data.decode("latin-1")), b""


def decode_transfer(body: bytes, encoding: str | None) -> bytes:
    enc = str(encoding or "").lower().strip()
    if enc == "base64":
        clean = re.sub(rb"[^A-Za-z0-9+/=]", b"", body)
        try:
            return base64.b64decode(clean + b"===")
        except Exception:
            return body
    if enc == "quoted-printable":
        try:
            import quopri
            return quopri.decodestring(body)
        except Exception:
            return body
    return body


def split_multipart(body: bytes, boundary: str) -> list[bytes]:
    mark = b"--" + boundary.encode("latin-1", "replace")
    chunks = body.split(mark)
    parts: list[bytes] = []
    for chunk in chunks[1:]:
        if chunk.startswith(b"--"):
            break
        if chunk.startswith(b"\r\n"):
            chunk = chunk[2:]
        elif chunk.startswith(b"\n"):
            chunk = chunk[1:]
        if chunk.endswith(b"\r\n"):
            chunk = chunk[:-2]
        elif chunk.endswith(b"\n"):
            chunk = chunk[:-1]
        parts.append(chunk)
    return parts


def collect_parts(data: bytes, out: list[dict], depth: int = 0) -> None:
    if depth > 8:
        return
    headers, body = split_entity(data)
    ctype, params = parse_content_type(headers.get("content-type", "text/plain"))
    if ctype.startswith("multipart/") and params.get("boundary"):
        for seg in split_multipart(body, params["boundary"]):
            collect_parts(seg, out, depth + 1)
        return
    if ctype in ("text/plain", "text/html", ""):
        raw = decode_transfer(body, headers.get("content-transfer-encoding"))
        text = decode_bytes(raw, params.get("charset"))
        out.append({
            "type": ctype or "text/plain",
            "text": html_to_text(text) if ctype == "text/html" else text,
            "html": text if ctype == "text/html" else "",
        })


# ------------------------------------------------------------------ #
# 验证码 / 链接识别                                                    #
# ------------------------------------------------------------------ #

CODE_KEYWORDS = (
    "验证码|校验码|动态码|确认码|安全码|激活码|登录码|一次性密码|密码|"
    "verification code|verification|security code|login code|code|passcode|password|otp|pin"
)


def extract_codes(subject: str, text: str) -> list[str]:
    hay = f"{subject or ''}\n{text or ''}"
    scores: dict[str, int] = {}

    def add(code: str, score: int) -> None:
        if not code:
            return
        if re.fullmatch(r"(19|20)\d{2}", code):
            return
        if code not in scores or scores[code] < score:
            scores[code] = score

    for m in re.finditer(rf"(?:{CODE_KEYWORDS})[^0-9A-Za-z]{{0,30}}([0-9]{{4,8}})", hay, re.I):
        add(m.group(1), 100)
    for m in re.finditer(rf"(?:{CODE_KEYWORDS})[^0-9A-Za-z]{{0,30}}([A-Z0-9]{{5,8}})", hay, re.I):
        code = m.group(1)
        if re.search(r"\d", code):
            add(code, 90)
    for m in re.finditer(r"(?<![\w.,/:-])(\d{6})(?![\w.,/-])", hay):
        add(m.group(1), 60)
    for m in re.finditer(r"(?<![\w.,/:-])(\d{5,8})(?![\w.,/-])", hay):
        add(m.group(1), 35)

    ranked = sorted(((s, c) for c, s in scores.items() if s >= 50), key=lambda x: (-x[0], len(x[1])))
    return [c for _, c in ranked[:5]]


def extract_links(text: str) -> list[str]:
    out: list[str] = []
    for m in re.finditer(r"https?://[^\s\"'<>)\]}]+", text or "", re.I):
        url = m.group(0).rstrip(".,;")
        if url not in out:
            out.append(url)
        if len(out) >= 3:
            break
    return out


# ------------------------------------------------------------------ #
# 邮件组装                                                             #
# ------------------------------------------------------------------ #

def iso_from_header(raw_date: str | None) -> str:
    if not raw_date:
        return datetime.now(timezone.utc).isoformat()
    try:
        dt = parsedate_to_datetime(raw_date)
        if dt is None:
            raise ValueError
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.isoformat()
    except Exception:
        return datetime.now(timezone.utc).isoformat()


def build_email(*, message_id: str, subject: str, sender: str, date_iso: str,
                body_text: str, body_html: str, folder: str, protocol: str) -> dict:
    text = (body_text or "").strip()
    if not text and body_html:
        text = html_to_text(body_html)
    codes = extract_codes(subject, text)
    preview = re.sub(r"\s+", " ", text).strip()[:300]
    from_name, from_addr = "", sender or ""
    parsed = getaddresses([sender]) if sender else []
    if parsed:
        name, addr = parsed[0]
        from_name = decode_rfc2047(name) if name else ""
        from_addr = addr or sender
    return {
        "messageId": message_id,
        "subject": subject or "(无主题)",
        "from": from_addr,
        "fromName": from_name,
        "date": date_iso,
        "bodyPreview": preview,
        "bodyText": text[:MAX_BODY_CHARS],
        "bodyHtml": (body_html or "")[:MAX_BODY_CHARS],
        "folder": folder,
        "protocol": protocol,
        "codes": codes,
        "links": extract_links(text),
    }


def parse_imap_message(raw: bytes, folder: str, uid: int) -> dict:
    headers, _ = split_entity(raw)
    parts: list[dict] = []
    collect_parts(raw, parts)
    plain = "\n".join(p["text"] for p in parts if p["type"] == "text/plain")
    html = "\n".join(p["html"] for p in parts if p["type"] == "text/html" and p["html"])
    subject = decode_rfc2047(headers.get("subject"))
    sender = decode_rfc2047(headers.get("from"))
    message_id = headers.get("message-id", "").strip().strip("<>")
    if not message_id:
        message_id = f"imap-{folder}-{uid}"
    return build_email(
        message_id=message_id,
        subject=subject,
        sender=sender,
        date_iso=iso_from_header(headers.get("date")),
        body_text=plain,
        body_html=html,
        folder=folder,
        protocol="imap",
    )


def matches_filters(email: dict, keyword: str, sender_filter: str, filters: list | None = None) -> bool:
    if sender_filter:
        hay = f"{email.get('from', '')} {email.get('fromName', '')}".lower()
        if sender_filter.lower() not in hay:
            return False
    if keyword:
        hay = f"{email.get('subject', '')}\n{email.get('bodyText', '')}".lower()
        if keyword.lower() not in hay:
            return False
    # 预设过滤条件：全部命中才算通过（AND）
    for rule in filters or []:
        if not isinstance(rule, dict):
            continue
        text = str(rule.get("text") or "").strip().lower()
        if not text:
            continue
        scope = str(rule.get("scope") or "all").lower()
        subject = str(email.get("subject") or "").lower()
        body = str(email.get("bodyText") or "").lower()
        hay = {"subject": subject, "body": body}.get(scope, f"{subject}\n{body}")
        if text not in hay:
            return False
    return True


# ------------------------------------------------------------------ #
# 预设过滤条件（持久化到 filters.json）                                 #
# ------------------------------------------------------------------ #

FILTERS_FILE = os.path.join(BASE_DIR, "filters.json")
_filters_lock = threading.Lock()


def load_filters() -> list[dict]:
    try:
        with open(FILTERS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        items = data.get("filters") if isinstance(data, dict) else data
        return [f for f in (items or []) if isinstance(f, dict) and str(f.get("text") or "").strip()]
    except (OSError, json.JSONDecodeError):
        return []


def save_filters(items: list[dict]) -> list[dict]:
    cleaned = []
    seen = set()
    for item in items or []:
        if not isinstance(item, dict):
            continue
        text = str(item.get("text") or "").strip()
        if not text:
            continue
        scope = str(item.get("scope") or "all").lower()
        if scope not in ("all", "subject", "body"):
            scope = "all"
        key = f"{scope}|{text.lower()}"
        if key in seen:
            continue
        seen.add(key)
        cleaned.append({
            "id": str(item.get("id") or secrets.token_hex(6)),
            "text": text,
            "scope": scope,
            "createdAt": item.get("createdAt") or datetime.now(timezone.utc).isoformat(),
        })
    tmp = FILTERS_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"filters": cleaned}, f, ensure_ascii=False, indent=2)
    os.replace(tmp, FILTERS_FILE)
    return cleaned


# ------------------------------------------------------------------ #
# 账号库（明文持久化到 accounts.json）                                  #
# ------------------------------------------------------------------ #

ACCOUNTS_FILE = os.path.join(BASE_DIR, "accounts.json")
_accounts_lock = threading.Lock()
ACCOUNT_FIELDS = ("email", "password", "clientId", "refreshToken")


def load_accounts() -> dict:
    try:
        with open(ACCOUNTS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        accounts = data.get("accounts") if isinstance(data, dict) else None
        if not isinstance(accounts, dict):
            return {}
        return {k: v for k, v in accounts.items() if isinstance(v, dict)}
    except (OSError, json.JSONDecodeError):
        return {}


def save_accounts(accounts: dict) -> None:
    tmp = ACCOUNTS_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"accounts": accounts}, f, ensure_ascii=False, indent=2)
    os.replace(tmp, ACCOUNTS_FILE)


def account_key(email_addr) -> str:
    return str(email_addr or "").strip().lower()


def resolve_credentials(payload: dict) -> dict:
    """用账号库补齐取件请求缺失的凭据；请求里显式给出的值优先。"""
    resolved = dict(payload)
    key = account_key(resolved.get("email"))
    if "@" not in key:
        return resolved
    with _accounts_lock:
        account = load_accounts().get(key)
    if not isinstance(account, dict):
        return resolved
    for field in ("password", "clientId", "refreshToken"):
        if not str(resolved.get(field) or "").strip():
            saved = account.get(field)
            if isinstance(saved, str) and saved:
                resolved[field] = saved
    return resolved


def update_account_refresh_token(email_addr: str, refresh_token: str) -> bool:
    """刷新成功后回写新令牌，仅当该邮箱已在账号库中。"""
    key = account_key(email_addr)
    if "@" not in key or not refresh_token:
        return False
    with _accounts_lock:
        accounts = load_accounts()
        account = accounts.get(key)
        if not isinstance(account, dict):
            return False
        account["refreshToken"] = refresh_token
        account["updatedAt"] = datetime.now(timezone.utc).isoformat()
        save_accounts(accounts)
    return True


def api_accounts_save(payload: dict) -> dict:
    items = payload.get("accounts")
    if not isinstance(items, list):
        items = []
    saved = 0
    errors: list[str] = []
    with _accounts_lock:
        accounts = load_accounts()
        for index, item in enumerate(items):
            label = f"第 {index + 1} 条记录"
            if not isinstance(item, dict):
                errors.append(f"{label}格式错误")
                continue
            email_addr = str(item.get("email") or "").strip()
            if not email_addr or "@" not in email_addr:
                errors.append(f"{label}邮箱地址无效")
                continue
            missing = [f for f in ACCOUNT_FIELDS if f not in item]
            if missing:
                errors.append(f"{email_addr}: 缺少字段 {', '.join(missing)}")
                continue
            accounts[account_key(email_addr)] = {
                "email": email_addr,
                "password": str(item.get("password") or ""),
                "clientId": str(item.get("clientId") or "").strip(),
                "refreshToken": str(item.get("refreshToken") or "").strip(),
                "updatedAt": datetime.now(timezone.utc).isoformat(),
            }
            saved += 1
        if saved:
            save_accounts(accounts)
    return {"success": True, "saved": saved, "errors": errors}


def api_accounts_get(payload: dict) -> dict:
    emails = payload.get("emails")
    if not isinstance(emails, list):
        emails = []
    wanted: list[str] = []
    for value in emails:
        key = account_key(value)
        if key and key not in wanted:
            wanted.append(key)
    result: dict[str, dict] = {}
    with _accounts_lock:
        accounts = load_accounts()
        for key in wanted:
            account = accounts.get(key)
            if isinstance(account, dict):
                item = {field: str(account.get(field) or "") for field in ACCOUNT_FIELDS}
                item["updatedAt"] = str(account.get("updatedAt") or "")
                result[key] = item
    return {"success": True, "accounts": result}


def api_accounts_delete(payload: dict) -> dict:
    emails = payload.get("emails")
    if not isinstance(emails, list):
        emails = []
    keys = {key for key in (account_key(v) for v in emails) if key}
    with _accounts_lock:
        accounts = load_accounts()
        deleted = 0
        for key in keys:
            if key in accounts:
                accounts.pop(key)
                deleted += 1
        if deleted:
            save_accounts(accounts)
    return {"success": True, "deleted": deleted}


# ------------------------------------------------------------------ #
# 桥接：直接调用桌面版 outlook_mail.js（本项目的后端接口）              #
# ------------------------------------------------------------------ #

TOOL_DIR = CONFIG.get("tool_dir") or os.path.join(BASE_DIR, "tool")
TOOL_SCRIPT = os.path.join(TOOL_DIR, "outlook_mail.js")
TOOL_ACCOUNT = os.path.join(TOOL_DIR, "account.txt")
# 工具用单个 account.txt 作输入，并发取件会互相覆盖，必须串行化
_tool_lock = threading.Lock()


def tool_available() -> bool:
    return bool(CONFIG.get("use_tool", True)) and os.path.isfile(TOOL_SCRIPT) and shutil.which("node") is not None


def write_tool_account(email_addr: str, password: str, client_id: str, refresh_token: str) -> None:
    os.makedirs(TOOL_DIR, exist_ok=True)
    line = f"{email_addr}----{password}----{client_id}----{refresh_token}\n"
    tmp = TOOL_ACCOUNT + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(line)
    os.replace(tmp, TOOL_ACCOUNT)
    try:
        os.chmod(TOOL_ACCOUNT, 0o600)
    except OSError:
        pass


def read_tool_account() -> str | None:
    try:
        with open(TOOL_ACCOUNT, "r", encoding="utf-8") as f:
            return f.read().strip() or None
    except OSError:
        return None


def parse_tool_date(value: str) -> str:
    """工具输出的日期是本地时间 'YYYY-MM-DD HH:MM:SS'。"""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            dt = datetime.strptime(value, fmt).replace(tzinfo=timezone.utc)
            return dt.isoformat()
        except ValueError:
            continue
    return datetime.now(timezone.utc).isoformat()


def run_tool(limit: int, folders: list[str], unseen: bool = False) -> list[dict]:
    """执行 node outlook_mail.js --json，返回工具解析后的邮件数组。"""
    cmd = [
        "node", TOOL_SCRIPT,
        "--account", TOOL_ACCOUNT,
        "-n", str(max(1, limit)),
        "--json",
        "--no-save",
        "--scan", ",".join(folders),
    ]
    if unseen:
        cmd.append("--unseen")
    proc = subprocess.run(
        cmd,
        cwd=TOOL_DIR,
        capture_output=True,
        text=True,
        timeout=int(CONFIG.get("tool_timeout", 90)),
    )
    out = (proc.stdout or "").strip()
    if not out:
        err = (proc.stderr or "").strip().splitlines()
        raise ImapError("；".join(err[-3:]) if err else "取件工具没有返回数据")
    try:
        data = json.loads(out)
    except json.JSONDecodeError:
        raise ImapError(f"取件工具输出无法解析: {out[:200]}")
    return data.get("mails") or []


def tool_mail_to_email(mail: dict, email_addr: str) -> dict:
    """把工具的邮件结构转成前端渲染结构。"""
    sender = str(mail.get("from") or "")
    date_iso = parse_tool_date(str(mail.get("date") or ""))
    body = str(mail.get("body") or "")
    return {
        "messageId": f"{email_addr}-{mail.get('folder')}-{mail.get('uid')}",
        "subject": str(mail.get("subject") or "(无主题)"),
        "from": sender,
        "fromName": "",
        "date": date_iso,
        "bodyPreview": re.sub(r"\s+", " ", body).strip()[:300],
        "bodyText": body[:MAX_BODY_CHARS],
        "bodyHtml": "",
        "folder": str(mail.get("folder") or ""),
        "protocol": "imap",
        # 验证码直接采用工具自身的识别结果
        "codes": [str(c) for c in (mail.get("codes") or [])],
        "links": [str(u) for u in (mail.get("links") or [])],
    }


# ------------------------------------------------------------------ #
# 取件：IMAP                                                          #
# ------------------------------------------------------------------ #

def fetch_imap(payload: dict) -> dict:
    payload = resolve_credentials(payload)
    email_addr = str(payload.get("email") or "").strip()
    client_id = str(payload.get("clientId") or CONFIG.get("default_client_id") or "").strip()
    refresh_token = str(payload.get("refreshToken") or "").strip()
    password = str(payload.get("password") or "")
    keyword = str(payload.get("keyword") or "").strip()
    sender_filter = str(payload.get("sender") or "").strip()
    filters = payload.get("filters") or []
    try:
        limit = int(payload.get("limit") or 1)
    except (TypeError, ValueError):
        limit = 10
    limit = max(1, min(limit, int(CONFIG.get("imap_fetch_cap", 20))))

    if not email_addr or "@" not in email_addr:
        raise TokenError("邮箱地址无效", code="INVALID_EMAIL", reason="导入数据里没有可用的邮箱地址")
    if not client_id:
        raise TokenError("缺少 client_id", code="MISSING_CLIENT_ID", reason="导入格式应为 邮箱----密码----clientid----刷新令牌")
    if not refresh_token:
        raise TokenError("缺少 refresh_token", code="MISSING_REFRESH_TOKEN", reason="导入格式应为 邮箱----密码----clientid----刷新令牌")

    scan = int(CONFIG.get("imap_fetch_cap", 20))
    want = limit if not (keyword or sender_filter) else min(scan * 3, 60)

    # 首选：直接调用桌面版 outlook_mail.js（本项目指定的后端接口）
    if tool_available():
        try:
            with _tool_lock:
                write_tool_account(email_addr, password, client_id, refresh_token)
                raw_mails = run_tool(want, DEFAULT_FOLDERS)
            mails = [tool_mail_to_email(m, email_addr) for m in raw_mails]
            filtered = [m for m in mails if matches_filters(m, keyword, sender_filter, filters)]
            filtered.sort(key=lambda m: m["date"], reverse=True)
            result = filtered[:limit]
            return {
                "success": True,
                "protocol": "imap",
                "engine": "outlook_mail.js",
                "count": len(result),
                "totalScanned": len(mails),
                "folders": DEFAULT_FOLDERS,
                "warnings": [],
                "emails": result,
            }
        except subprocess.TimeoutExpired:
            raise ImapError("取件工具执行超时（90 秒），请稍后重试或降低取件数量")
        except (ImapError, OSError) as exc:
            # 工具不可用/失败时回落到内置 IMAP 客户端，避免整条链路不可用
            log(f"[warn] outlook_mail.js 调用失败，回落内置 IMAP: {exc}")
        except Exception as exc:
            log(f"[warn] outlook_mail.js 异常，回落内置 IMAP: {exc!r}")

    token = get_access_token(email_addr, client_id, refresh_token, IMAP_SCOPE, f"imap:{email_addr}:{client_id}")

    client = ImapClient(IMAP_HOST, IMAP_PORT, timeout=30)
    mails = []
    folder_errors: list[str] = []
    try:
        client.connect()
        client.login(email_addr, token)
        for folder in DEFAULT_FOLDERS:
            try:
                client.select(folder)
                uids = client.uid_search("ALL")
            except ImapError as exc:
                folder_errors.append(f"{folder}: {exc}")
                continue
            for uid in uids[-want:]:
                try:
                    raw = client.uid_fetch(uid)
                except ImapError as exc:
                    folder_errors.append(f"{folder}/{uid}: {exc}")
                    continue
                if not raw:
                    continue
                try:
                    mails.append(parse_imap_message(raw, folder, uid))
                except Exception as exc:
                    folder_errors.append(f"{folder}/{uid} 解析失败: {exc}")
    finally:
        client.logout()

    if not mails and folder_errors:
        raise ImapError("；".join(folder_errors[:3]))

    filtered = [m for m in mails if matches_filters(m, keyword, sender_filter, filters)]
    filtered.sort(key=lambda m: m["date"], reverse=True)
    result = filtered[:limit]
    return {
        "success": True,
        "protocol": "imap",
        "engine": "builtin",
        "count": len(result),
        "totalScanned": len(mails),
        "folders": DEFAULT_FOLDERS,
        "warnings": folder_errors[:5],
        "emails": result,
    }


# ------------------------------------------------------------------ #
# 取件：Graph                                                         #
# ------------------------------------------------------------------ #

GRAPH_SELECT = "id,internetMessageId,subject,from,receivedDateTime,body,bodyPreview,webLink,parentFolderId"


def graph_request(path: str, token: str, timeout: int = 30) -> dict:
    req = urllib.request.Request(
        GRAPH_ROOT + path,
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8", "replace"))
    except urllib.error.HTTPError as exc:
        text = exc.read().decode("utf-8", "replace")
        detail = text[:400]
        try:
            err = json.loads(text).get("error", {})
            code = err.get("code", "")
            detail = err.get("message", detail)
        except Exception:
            code = ""
        raise TokenError(
            message=f"Graph API 错误: {code or exc.code} {detail}".strip(),
            code=code or f"GRAPH_HTTP_{exc.code}",
            reason="Graph 无法访问该邮箱" if code in ("ErrorInvalidUser", "MailboxNotEnabledForRESTAPI") else "Graph 请求被拒绝",
            action="该账号可能未授权 Graph 权限，可关闭 Graph 只保留 IMAP 取件",
            detail=text[:1500],
        )


def graph_fetch_folder(token: str, folder: str, top: int, keyword: str) -> list[dict]:
    params = {
        "$top": str(min(top, 50)),
        "$select": GRAPH_SELECT,
    }
    if keyword:
        safe = keyword.replace('"', " ")
        params["$search"] = f'"{safe}"'
    else:
        params["$orderby"] = "receivedDateTime desc"
    path = f"/me/mailFolders/{folder}/messages?" + urllib.parse.urlencode(params)
    data = graph_request(path, token)
    return data.get("value") or []


def fetch_graph(payload: dict) -> dict:
    payload = resolve_credentials(payload)
    email_addr = str(payload.get("email") or "").strip()
    client_id = str(payload.get("clientId") or CONFIG.get("default_client_id") or "").strip()
    refresh_token = str(payload.get("refreshToken") or "").strip()
    keyword = str(payload.get("keyword") or "").strip()
    sender_filter = str(payload.get("sender") or "").strip()
    filters = payload.get("filters") or []
    try:
        limit = int(payload.get("limit") or 1)
    except (TypeError, ValueError):
        limit = 10
    limit = max(1, min(limit, 50))

    if not email_addr or "@" not in email_addr:
        raise TokenError("邮箱地址无效", code="INVALID_EMAIL", reason="导入数据里没有可用的邮箱地址")
    if not client_id:
        raise TokenError("缺少 client_id", code="MISSING_CLIENT_ID", reason="导入格式应为 邮箱----密码----clientid----刷新令牌")
    if not refresh_token:
        raise TokenError("缺少 refresh_token", code="MISSING_REFRESH_TOKEN", reason="导入格式应为 邮箱----密码----clientid----刷新令牌")

    token = get_access_token(email_addr, client_id, refresh_token, GRAPH_SCOPE, f"graph:{email_addr}:{client_id}")

    mails: list[dict] = []
    errors: list[str] = []
    for folder in ("inbox", "junkemail"):
        try:
            items = graph_fetch_folder(token, folder, max(limit * 2, 20), keyword)
        except TokenError as exc:
            errors.append(f"{folder}: {exc}")
            continue
        for item in items:
            body = item.get("body") or {}
            content = body.get("content") or ""
            content_type = (body.get("contentType") or "text").lower()
            sender_obj = ((item.get("from") or {}).get("emailAddress") or {})
            mails.append(build_email(
                # 用 RFC822 的 internetMessageId，IMAP 与 Graph 才能按同一封邮件去重
                message_id=(item.get("internetMessageId") or item.get("id") or "").strip(),
                subject=item.get("subject") or "",
                sender=f"{sender_obj.get('name', '')} <{sender_obj.get('address', '')}>".strip(),
                date_iso=item.get("receivedDateTime") or datetime.now(timezone.utc).isoformat(),
                body_text=content if content_type == "text" else "",
                body_html=content if content_type == "html" else "",
                folder=folder,
                protocol="graph",
            ))

    if not mails and errors:
        raise TokenError(
            message="；".join(errors[:2]),
            code="GRAPH_UNAVAILABLE",
            reason="该账号或应用没有 Graph 权限",
            action="关闭 Graph 开关，仅使用 IMAP 取件",
            detail="；".join(errors[:2]),
        )

    filtered = [m for m in mails if matches_filters(m, keyword, sender_filter, filters)]
    filtered.sort(key=lambda m: m["date"], reverse=True)
    result = filtered[:limit]
    return {
        "success": True,
        "protocol": "graph",
        "count": len(result),
        "totalScanned": len(mails),
        "warnings": errors[:5],
        "emails": result,
    }


# ------------------------------------------------------------------ #
# 取件：Mail.com（账号 + 密码，IMAP 明文登录）                          #
# ------------------------------------------------------------------ #

MAIL_HOSTS = {
    "mail.com": ("imap.mail.com", 993),
    "gmx.com": ("imap.gmx.com", 993),
    "gmx.net": ("imap.gmx.net", 993),
    "gmx.de": ("imap.gmx.net", 993),
}

MAIL_FOLDERS = ["INBOX", "Spam"]


def resolve_mail_host(email_addr: str) -> tuple[str, int]:
    domain = email_addr.rsplit("@", 1)[-1].lower()
    if domain in MAIL_HOSTS:
        return MAIL_HOSTS[domain]
    # 通用回落：多数邮箱服务用 imap.<域名>
    return f"imap.{domain}", 993


def fetch_mail(payload: dict) -> dict:
    payload = resolve_credentials(payload)
    email_addr = str(payload.get("email") or "").strip()
    password = str(payload.get("password") or "")
    keyword = str(payload.get("keyword") or "").strip()
    sender_filter = str(payload.get("sender") or "").strip()
    filters = payload.get("filters") or []
    try:
        limit = int(payload.get("limit") or 1)
    except (TypeError, ValueError):
        limit = 10
    limit = max(1, min(limit, 20))

    if not email_addr or "@" not in email_addr:
        raise TokenError("邮箱地址无效", code="INVALID_EMAIL", reason="导入格式应为 邮箱:密码")
    if not password:
        raise TokenError("缺少密码", code="MISSING_PASSWORD", reason="导入格式应为 邮箱:密码")

    host, port = resolve_mail_host(email_addr)
    scan = 30
    client = ImapClient(host, port, timeout=30)
    mails: list[dict] = []
    warnings: list[str] = []
    try:
        client.connect()
        try:
            client.plain_login(email_addr, password)
        except ImapError as exc:
            raise TokenError(
                message=f"IMAP 登录失败: {exc}",
                code="MAIL_AUTH_FAILED",
                reason="账号或密码不正确，或该邮箱未开启 IMAP",
                action="到邮箱网页设置里确认已开启 IMAP 访问，并使用「应用专用密码」",
                detail=str(exc),
            )
        for folder in MAIL_FOLDERS:
            try:
                client.select(folder)
                uids = client.uid_search("ALL")
            except ImapError as exc:
                warnings.append(f"{folder}: {exc}")
                continue
            for uid in uids[-scan:]:
                try:
                    raw = client.uid_fetch(uid)
                except ImapError as exc:
                    warnings.append(f"{folder}/{uid}: {exc}")
                    continue
                if not raw:
                    continue
                try:
                    mail = parse_imap_message(raw, folder, uid)
                    mail["protocol"] = "mail"
                    mails.append(mail)
                except Exception as exc:
                    warnings.append(f"{folder}/{uid} 解析失败: {exc}")
    finally:
        client.logout()

    if not mails and warnings:
        raise ImapError("；".join(warnings[:3]))

    filtered = [m for m in mails if matches_filters(m, keyword, sender_filter, filters)]
    filtered.sort(key=lambda m: m["date"], reverse=True)
    result = filtered[:limit]
    return {
        "success": True,
        "protocol": "mail",
        "count": len(result),
        "totalScanned": len(mails),
        "folders": MAIL_FOLDERS,
        "host": host,
        "sessionExpiresAt": (datetime.now(timezone.utc) + timedelta(hours=1)).isoformat(),
        "warnings": warnings[:5],
        "emails": result,
    }


# ------------------------------------------------------------------ #
# HTTP 服务                                                            #
# ------------------------------------------------------------------ #

class Handler(BaseHTTPRequestHandler):
    server_version = "MailGrabber/1.0"
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):  # noqa: A003
        log(f"{self.address_string()} - {fmt % args}")

    # ---------------- 工具 ----------------

    def _send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path: str) -> None:
        try:
            with open(path, "rb") as f:
                data = f.read()
        except OSError:
            self.send_error(404, "Not Found")
            return
        ctype, _ = mimetypes.guess_type(path)
        if path.endswith(".js"):
            ctype = "application/javascript; charset=utf-8"
        elif path.endswith(".css"):
            ctype = "text/css; charset=utf-8"
        elif path.endswith(".html"):
            ctype = "text/html; charset=utf-8"
        self.send_response(200)
        self.send_header("Content-Type", ctype or "application/octet-stream")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-cache")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(data)

    def _read_body(self) -> dict:
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        if length > 4 * 1024 * 1024:
            raise ValueError("请求体过大")
        raw = self.rfile.read(length)
        try:
            data = json.loads(raw.decode("utf-8"))
        except Exception:
            raise ValueError("请求体不是合法 JSON")
        return data if isinstance(data, dict) else {}

    # ---------------- 路由 ----------------

    def do_GET(self):  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path == "/api/health":
            self._send_json({"success": True, "ok": True, "time": datetime.now(timezone.utc).isoformat()})
            return
        if path == "/api/filters":
            with _filters_lock:
                self._send_json({"success": True, "filters": load_filters()})
            return
        if path.startswith("/api/"):
            self._send_json({"success": False, "error": "未知接口"}, 404)
            return
        if path == "/":
            path = "/index.html"
        safe = os.path.normpath(os.path.join(WWW_DIR, path.lstrip("/")))
        if not safe.startswith(WWW_DIR) or not os.path.isfile(safe):
            self.send_error(404, "Not Found")
            return
        self._send_file(safe)

    def do_POST(self):  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        try:
            body = self._read_body()
        except ValueError as exc:
            self._send_json({"success": False, "error": str(exc)}, 400)
            return

        if path == "/api/security-session":
            self._send_json(create_session())
            return

        if path == "/api/filters":
            with _filters_lock:
                self._send_json({"success": True, "filters": save_filters(body.get("filters") or [])})
            return

        if path not in (
            "/api/fetch-imap", "/api/fetch-graph", "/api/fetch-mail", "/api/refresh-token",
            "/api/accounts/save", "/api/accounts/get", "/api/accounts/delete",
        ):
            self._send_json({"success": False, "error": "未知接口"}, 404)
            return

        try:
            payload = open_envelope(body)
        except EnvelopeError as exc:
            self._send_json({"success": False, "code": exc.code, "error": str(exc)}, 403)
            return

        try:
            if path == "/api/accounts/save":
                result = api_accounts_save(payload)
            elif path == "/api/accounts/get":
                result = api_accounts_get(payload)
            elif path == "/api/accounts/delete":
                result = api_accounts_delete(payload)
            elif path == "/api/fetch-imap":
                result = fetch_imap(payload)
            elif path == "/api/fetch-graph":
                result = fetch_graph(payload)
            elif path == "/api/fetch-mail":
                result = fetch_mail(payload)
            else:
                payload = resolve_credentials(payload)
                email_addr = str(payload.get("email") or "").strip()
                client_id = str(payload.get("clientId") or CONFIG.get("default_client_id") or "").strip()
                refresh_token = str(payload.get("refreshToken") or "").strip()
                if not email_addr or "@" not in email_addr:
                    raise TokenError("邮箱地址无效", code="INVALID_EMAIL", reason="导入数据里没有可用的邮箱地址")
                if not client_id:
                    raise TokenError("缺少 client_id", code="MISSING_CLIENT_ID", reason="导入格式应为 邮箱----密码----clientid----刷新令牌")
                if not refresh_token:
                    raise TokenError("缺少 refresh_token", code="MISSING_REFRESH_TOKEN", reason="导入格式应为 邮箱----密码----clientid----刷新令牌")
                result = refresh_refresh_token(email_addr, client_id, refresh_token)
                if result.get("success") and result.get("refreshToken"):
                    update_account_refresh_token(email_addr, str(result["refreshToken"]))
            self._send_json(result)
        except TokenError as exc:
            self._send_json({
                "success": False,
                "error": str(exc),
                "code": exc.code,
                "reason": exc.reason,
                "action": exc.action,
                "detail": exc.detail,
            }, 200)
        except ImapError as exc:
            self._send_json({
                "success": False,
                "error": str(exc),
                "code": "IMAP_ERROR",
                "reason": "IMAP 取件失败",
                "action": "检查该邮箱的 IMAP 授权与令牌是否有效",
                "detail": str(exc),
            }, 200)
        except Exception as exc:  # pragma: no cover
            log(f"[error] {path}: {exc!r}")
            self._send_json({
                "success": False,
                "error": f"服务内部错误: {exc}",
                "code": "INTERNAL_ERROR",
                "reason": "服务器处理异常",
                "action": "稍后重试",
                "detail": repr(exc),
            }, 200)


def main() -> None:
    host = CONFIG.get("host", "127.0.0.1")
    port = int(CONFIG.get("port", 8090))
    httpd = ThreadingHTTPServer((host, port), Handler)
    httpd.daemon_threads = True
    log(f"[start] Outlook 快速取件 已启动 http://{host}:{port}  (静态目录: {WWW_DIR})")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        log("[stop] 收到中断，正在退出")
    finally:
        httpd.server_close()


if __name__ == "__main__":
    main()
