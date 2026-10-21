#!/usr/bin/env node
/*
 * Outlook 取件工具 (零依赖 Node.js 版)  v1.0
 * ------------------------------------------------------------
 * 用微软 OAuth2 的 refresh_token 换取 Access Token,
 * 再通过 IMAP XOAUTH2 协议登录 outlook.office365.com 拉取邮件,
 * 自动提取验证码 / 验证链接。
 *
 * 只依赖 Node.js 内置模块, 不需要 npm install。
 * 账号信息从同目录的 account.txt 读取, 格式:
 *   邮箱----密码----client_id----refresh_token
 *
 * 常用:
 *   node outlook_mail.js                 看收件箱最新 5 封
 *   node outlook_mail.js -n 10           看最新 10 封
 *   node outlook_mail.js --code          只输出验证码
 *   node outlook_mail.js -m "Junk Email" 看垃圾邮件
 *   node outlook_mail.js --watch 15      每 15 秒轮询新邮件(守验证码)
 *   node outlook_mail.js --json          机器可读输出
 */
'use strict';

const tls = require('node:tls');
const fs = require('node:fs');
const path = require('node:path');
const https = require('node:https');

const VERSION = '1.0';
const BASE_DIR = __dirname;
const DEFAULT_ACCOUNT_FILE = path.join(BASE_DIR, 'account.txt');
const TOKEN_CACHE_FILE = path.join(BASE_DIR, '.token_cache.json');

const IMAP_HOST = 'outlook.office365.com';
const IMAP_PORT = 993;
const TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

/**
 * IMAP 登录必须拿到 audience=outlook.office.com 的令牌。
 * 不带 scope 刷新出来的往往是 Graph 或其他资源的令牌, 服务器会回 "AUTHENTICATE failed."。
 */
const IMAP_SCOPE = 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access';

const log = (...a) => process.stderr.write(a.join(' ') + '\n');

/* ------------------------------------------------------------------ */
/* 参数                                                                */
/* ------------------------------------------------------------------ */

function printHelp() {
  const text = `
Outlook 取件工具 v${VERSION}  (零依赖 Node.js 版)

用法: node outlook_mail.js [选项]

选项:
  -n, --limit <数>      查看最新 N 封 (默认 5)
  -m, --mailbox <名>    只查指定文件夹 (默认同时扫 INBOX + Junk)
      --scan <a,b>      自定义要扫的文件夹, 逗号分隔 (默认 "INBOX,Junk")
  -a, --account <路径>  指定账号文件 (默认同目录 account.txt)
  -u, --unseen          只看未读邮件
      --code            只输出验证码 (便于脚本取用)
      --json            输出 JSON
      --raw             正文不截断
  -w, --watch [秒]      持续轮询新邮件, 默认间隔 15 秒 (Ctrl+C 退出)
      --exit-on-code    轮询模式下取到验证码就退出
      --folders         列出所有文件夹名后退出
      --no-cache        忽略本地 Access Token 缓存, 强制重新换取
      --no-save         不把轮换后的 refresh_token 写回 account.txt
      --scope <s>       自定义 scope (一般不需要)
  -h, --help            显示本帮助

说明:
  验证码邮件经常被微软判成垃圾邮件扔进 Junk, 所以默认两个文件夹一起查,
  按时间倒序合并, 每封都会标出它所在的文件夹。

示例:
  node outlook_mail.js -n 3 --code
  node outlook_mail.js --watch 20 --exit-on-code
  node outlook_mail.js -m Junk -n 5
`;
  process.stdout.write(text.trimStart());
}

function parseArgs(argv) {
  const o = {
    limit: 5, mailbox: null, scan: null, account: null, unseen: false,
    codeOnly: false, json: false, raw: false, watch: 0, exitOnCode: false,
    folders: false, noCache: false, noSave: false, scope: null, help: false,
  };
  for (let i = 0; i < argv.length; i++) {
    let a = argv[i];
    let inline = null;
    const eq = a.indexOf('=');
    if (a.startsWith('--') && eq > 0) { inline = a.slice(eq + 1); a = a.slice(0, eq); }
    const take = (name) => {
      if (inline !== null) return inline;
      const n = argv[++i];
      if (n === undefined) throw new Error(`参数 ${name} 缺少取值`);
      return n;
    };
    switch (a) {
      case '-n': case '--limit': {
        const v = parseInt(take(a), 10);
        o.limit = Number.isFinite(v) && v > 0 ? v : 5;
        break;
      }
      case '-m': case '--mailbox': o.mailbox = take(a); break;
      case '--scan': o.scan = take(a).split(',').map((s) => s.trim()).filter(Boolean); break;
      case '-a': case '--account': o.account = take(a); break;
      case '-u': case '--unseen': o.unseen = true; break;
      case '--code': o.codeOnly = true; break;
      case '--json': o.json = true; break;
      case '--raw': o.raw = true; break;
      case '--exit-on-code': o.exitOnCode = true; break;
      case '--folders': o.folders = true; break;
      case '--no-cache': o.noCache = true; break;
      case '--no-save': o.noSave = true; break;
      case '--scope': o.scope = take(a); break;
      case '-h': case '--help': o.help = true; break;
      case '-w': case '--watch': {
        if (inline !== null) { o.watch = parseInt(inline, 10) || 15; break; }
        const nxt = argv[i + 1];
        if (nxt !== undefined && /^\d+$/.test(nxt)) { o.watch = parseInt(argv[++i], 10) || 15; }
        else o.watch = 15;
        if (o.watch < 3) o.watch = 3;
        break;
      }
      default:
        throw new Error(`未知参数: ${a}  (试试 --help)`);
    }
  }
  return o;
}

/* ------------------------------------------------------------------ */
/* 账号 / 缓存文件                                                     */
/* ------------------------------------------------------------------ */

function loadAccount(file) {
  const target = file ? path.resolve(file) : DEFAULT_ACCOUNT_FILE;
  if (!fs.existsSync(target)) {
    throw new Error(`找不到账号文件: ${target}\n请把 "邮箱----密码----client_id----refresh_token" 写入该文件。`);
  }
  const raw = fs.readFileSync(target, 'utf8');
  const squashed = raw.replace(/\s+/g, '');
  if (!squashed) throw new Error(`账号文件为空: ${target}`);
  const parts = squashed.split('----');
  if (parts.length !== 4) {
    throw new Error(`账号文件格式不对: 需要 4 段(用 ---- 分隔), 实际 ${parts.length} 段\n文件: ${target}`);
  }
  const [email, password, clientId, refreshToken] = parts;
  if (!/^[^@\s]+@[^@\s]+$/.test(email)) throw new Error(`邮箱地址看起来不对: ${email}`);
  if (!refreshToken) throw new Error('refresh_token 为空');
  return { email, password, clientId, refreshToken, file: target, raw };
}

/** refresh_token 会被微软轮换, 成功后写回账号文件, 否则下次可能失效 */
function persistRefreshToken(acct, newToken) {
  try {
    if (acct.raw.includes('#')) {
      log('[提示] account.txt 含注释行, 为避免破坏内容, 不自动更新 refresh_token');
      return false;
    }
    fs.writeFileSync(
      acct.file,
      `${acct.email}----${acct.password}----${acct.clientId}----${newToken}\n`,
      'utf8',
    );
    acct.refreshToken = newToken;
    log('[提示] 微软已轮换 refresh_token, 已自动写回 account.txt');
    return true;
  } catch (e) {
    log('[提示] 写回 account.txt 失败(不影响本次使用): ' + e.message);
    return false;
  }
}

function readJson(file) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return null; }
}
function writeJson(file, obj) {
  try { fs.writeFileSync(file, JSON.stringify(obj, null, 2), 'utf8'); } catch { /* 忽略 */ }
}

/* ------------------------------------------------------------------ */
/* OAuth2                                                              */
/* ------------------------------------------------------------------ */

function httpPostForm(url, form) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams(form).toString();
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname,
      port: 443,
      path: u.pathname + u.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body),
        'Accept': 'application/json',
      },
      timeout: 30000,
    }, (res) => {
      let data = '';
      res.setEncoding('utf8');
      res.on('data', (c) => { data += c; });
      res.on('end', () => resolve({ status: res.statusCode, text: data }));
    });
    req.on('timeout', () => req.destroy(new Error('请求 token 端点超时')));
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function exchangeRefreshToken(acct, scope) {
  const form = {
    client_id: acct.clientId,
    refresh_token: acct.refreshToken,
    grant_type: 'refresh_token',
  };
  if (scope) form.scope = scope;
  const res = await httpPostForm(TOKEN_URL, form);
  let j = null;
  try { j = JSON.parse(res.text); } catch { /* 非 JSON */ }
  return { status: res.status, json: j, text: res.text };
}

function describeTokenError(status, j, text) {
  const errCode = (j && j.error) || '';
  const desc = j && j.error_description ? String(j.error_description).split(/\r?\n/)[0] : String(text || '').slice(0, 400);
  let hint = '';
  if (/invalid_grant/i.test(errCode + desc)) {
    hint = '\n>>> refresh_token 已失效/被吊销: 常见原因是被轮换后没保存、改过密码、或长期未使用。需要重新获取邮箱密钥。';
  } else if (/invalid_client/i.test(errCode + desc)) {
    hint = '\n>>> client_id 与这份 refresh_token 不匹配。';
  } else if (/AADSTS90023|No applicable permissions|unauthorized_client|consent/i.test(errCode + desc)) {
    hint = '\n>>> 该账号/应用没有所需的邮箱权限 (缺少 IMAP 授权)。';
  }
  return `换取 Access Token 失败 (HTTP ${status}) ${errCode}: ${desc}${hint}`;
}

async function getAccessToken(acct, opts) {
  const cache = readJson(TOKEN_CACHE_FILE);
  if (!opts.noCache && cache && cache.email === acct.email && cache.access_token
      && cache.expires_at > Date.now() + 120000) {
    const left = Math.round((cache.expires_at - Date.now()) / 60000);
    log(`[1/3] 复用缓存的 Access Token (约剩 ${left} 分钟)`);
    return cache.access_token;
  }

  // 先要 IMAP 权限的令牌; 若该应用没有此权限, 再退回"沿用原授权"的方式
  const attempts = opts.scope ? [{ scope: opts.scope, label: '自定义 scope' }]
    : [{ scope: IMAP_SCOPE, label: 'IMAP scope' }, { scope: null, label: '原授权 scope' }];

  let lastErrText = null;
  for (let i = 0; i < attempts.length; i++) {
    const { scope, label } = attempts[i];
    log(`[1/3] 正在用 refresh_token 换取 Access Token (${label}) ...`);
    const r = await exchangeRefreshToken(acct, scope);

    if (r.status === 200 && r.json && r.json.access_token) {
      const j = r.json;
      if (j.refresh_token && j.refresh_token !== acct.refreshToken && !opts.noSave) {
        persistRefreshToken(acct, j.refresh_token);
      }
      const ttl = (Number(j.expires_in) || 3600) * 1000;
      writeJson(TOKEN_CACHE_FILE, {
        email: acct.email,
        scope: scope || '',
        access_token: j.access_token,
        expires_at: Date.now() + ttl - 60000,
        saved_at: new Date().toISOString(),
      });
      log(`[1/3] 拿到 Access Token (有效期 ${Math.round(ttl / 60000)} 分钟)`);
      return j.access_token;
    }

    lastErrText = describeTokenError(r.status, r.json, r.text);
    const msg = lastErrText;
    // refresh_token 本身坏了, 再换 scope 也没用, 直接报错
    if (/invalid_grant|invalid_client/i.test(msg)) throw new Error(msg);
    if (i < attempts.length - 1) log('[1/3] 该 scope 不适用, 换一种方式重试 ...');
  }
  throw new Error(lastErrText || '换取 Access Token 失败');
}

/* ------------------------------------------------------------------ */
/* 极简 IMAP 客户端 (支持 literal)                                      */
/* ------------------------------------------------------------------ */

class ImapClient {
  constructor(host, port) {
    this.host = host;
    this.port = port;
    this.buf = Buffer.alloc(0);
    this.current = { lines: [], literals: [] };
    this.needLiteral = null;
    this.responses = [];
    this.waiters = [];
    this.error = null;
    this.closed = false;
    this.tagSeq = 0;
  }

  connect() {
    return new Promise((resolve, reject) => {
      const sock = tls.connect({
        host: this.host,
        port: this.port,
        servername: this.host,
      }, () => resolve(this));
      this.sock = sock;
      sock.setNoDelay(true);
      sock.setTimeout(45000);
      sock.on('data', (d) => this._onData(d));
      sock.on('timeout', () => {
        this.error = new Error('IMAP 连接超时');
        try { sock.destroy(); } catch { /* 忽略 */ }
        this._drain();
      });
      sock.on('error', (e) => {
        this.error = this.error || e;
        this._drain();
      });
      sock.on('close', () => {
        this.closed = true;
        this._drain();
      });
      sock.once('error', reject);
    });
  }

  _onData(chunk) {
    this.buf = this.buf.length ? Buffer.concat([this.buf, chunk]) : chunk;
    for (;;) {
      if (this.needLiteral !== null) {
        if (this.buf.length < this.needLiteral) return;
        this.current.literals.push(this.buf.subarray(0, this.needLiteral));
        this.buf = this.buf.subarray(this.needLiteral);
        this.needLiteral = null;
        continue;
      }
      const idx = this.buf.indexOf('\r\n');
      if (idx < 0) return;
      const line = this.buf.subarray(0, idx).toString('latin1');
      this.buf = this.buf.subarray(idx + 2);
      this.current.lines.push(line);
      const m = /\{(\d+)\}$/.exec(line);
      if (m) { this.needLiteral = Number(m[1]); continue; }
      const resp = this.current;
      const first = resp.lines[0] || '';
      resp.text = resp.lines.join('\n');
      resp.type = first.startsWith('+') ? 'continuation'
        : (first.startsWith('*') || first.startsWith('OK') ? 'untagged' : 'tagged');
      this.current = { lines: [], literals: [] };
      this.responses.push(resp);
      this._drain();
    }
  }

  _drain() {
    while (this.waiters.length) {
      if (this.responses.length) {
        this.waiters.shift().resolve(this.responses.shift());
      } else if (this.error) {
        this.waiters.shift().reject(this.error);
      } else if (this.closed) {
        this.waiters.shift().reject(new Error('IMAP 连接已关闭'));
      } else {
        break;
      }
    }
  }

  next() {
    if (this.responses.length) return Promise.resolve(this.responses.shift());
    if (this.error) return Promise.reject(this.error);
    if (this.closed) return Promise.reject(new Error('IMAP 连接已关闭'));
    return new Promise((resolve, reject) => this.waiters.push({ resolve, reject }));
  }

  _write(s) {
    if (this.closed) throw new Error('IMAP 连接已关闭');
    this.sock.write(s);
  }

  /** 连接后的第一行是服务器问候语, 必须先消费掉 */
  async greeting() {
    const r = await this.next();
    const first = (r.lines && r.lines[0]) || '';
    if (!/^\*\s+(OK|PREAUTH)\b/i.test(first)) {
      throw new Error('IMAP 服务器问候异常: ' + (first || '(空)'));
    }
    return first;
  }

  nextTag() { return 'T' + (++this.tagSeq); }

  async command(line, tag) {
    const t = tag || this.nextTag();
    this._write(`${t} ${line}\r\n`);
    const untagged = [];
    for (;;) {
      const r = await this.next();
      if (r.type === 'tagged') {
        const m = new RegExp(`^${t}\\s+(OK|NO|BAD)\\b\\s*(.*)$`, 'i').exec(r.lines[0] || '');
        if (!m) continue;
        if (m[1].toUpperCase() !== 'OK') {
          const e = new Error(`IMAP ${m[1]}: ${m[2]}`);
          e.untagged = untagged;
          throw e;
        }
        return { text: m[2], untagged, raw: r };
      }
      if (r.type === 'continuation') {
        throw new Error('IMAP 协议异常: 命令阶段收到续行');
      }
      untagged.push(r);
    }
  }

  async login(user, accessToken) {
    const t = this.nextTag();
    this._write(`${t} AUTHENTICATE XOAUTH2\r\n`);
    let r = await this.next();
    if (r.type === 'tagged') {
      const m = new RegExp(`^${t}\\s+(OK|NO|BAD)\\b\\s*(.*)$`, 'i').exec(r.lines[0] || '');
      if (m && m[1].toUpperCase() === 'OK') return true;
      throw new Error('XOAUTH2 认证失败: ' + (r.lines[0] || ''));
    }
    if (r.type !== 'continuation') throw new Error('XOAUTH2 认证失败: 服务器未要求续行');

    const sasl = Buffer.from(`user=${user}\x01auth=Bearer ${accessToken}\x01\x01`, 'utf8').toString('base64');
    this._write(sasl + '\r\n');

    for (;;) {
      r = await this.next();
      if (r.type === 'continuation') {
        let info = '';
        try {
          const raw = (r.lines[0] || '').replace(/^\+\s?/, '');
          info = Buffer.from(raw, 'base64').toString('utf8');
        } catch { /* 忽略 */ }
        this._write('\r\n');
        try { await this.next(); } catch { /* 忽略 */ }
        const err = new Error('XOAUTH2 认证失败' + (info ? ': ' + info : ''));
        err.authFailed = true;
        throw err;
      }
      if (r.type === 'tagged') {
        const m = new RegExp(`^${t}\\s+(OK|NO|BAD)\\b\\s*(.*)$`, 'i').exec(r.lines[0] || '');
        if (!m) continue;
        if (m[1].toUpperCase() === 'OK') return true;
        throw new Error('XOAUTH2 认证失败: ' + m[2]);
      }
    }
  }

  async select(mailbox) {
    const res = await this.command(`SELECT "${mailbox.replace(/"/g, '\\"')}"`);
    const info = { exists: 0, uidNext: 0, uidValidity: 0 };
    for (const r of res.untagged) {
      const t = r.lines.join('\n');
      let m = /^\*\s+(\d+)\s+EXISTS/m.exec(t);
      if (m) info.exists = Number(m[1]);
      m = /\[UIDNEXT\s+(\d+)\]/i.exec(t);
      if (m) info.uidNext = Number(m[1]);
      m = /\[UIDVALIDITY\s+(\d+)\]/i.exec(t);
      if (m) info.uidValidity = Number(m[1]);
    }
    return info;
  }

  async listFolders() {
    const res = await this.command('LIST "" "*"');
    const out = [];
    for (const r of res.untagged) {
      for (const line of r.lines) {
        const m = /^\*\s+LIST\s+\([^)]*\)\s+(?:"(?:[^"\\]|\\.)*"|NIL)\s+(.+)$/i.exec(line);
        if (m) out.push(m[1].replace(/^"|"$/g, ''));
      }
    }
    return out;
  }

  async search(criteria) {
    const res = await this.command(`UID SEARCH ${criteria}`);
    const ids = [];
    for (const r of res.untagged) {
      const line = r.lines[0] || '';
      if (!/^\*\s+SEARCH\b/i.test(line)) continue;
      for (const p of line.split(/\s+/).slice(2)) if (/^\d+$/.test(p)) ids.push(Number(p));
    }
    return ids;
  }

  async fetchRaw(uid) {
    const res = await this.command(`UID FETCH ${uid} (BODY.PEEK[])`);
    for (const r of res.untagged) if (r.literals.length) return r.literals[0];
    return null;
  }

  /** 批量 FETCH：一条命令取多封，避免逐封往返（每封约 200ms 网络往返） */
  async fetchRawBatch(uids) {
    const out = new Map();
    if (!uids.length) return out;
    const res = await this.command(`UID FETCH ${uids.join(',')} (BODY.PEEK[])`);
    for (const r of res.untagged) {
      // UID 出现在响应块的第二行（"* 16 FETCH (BODY[] {9988}" + " UID 134)"），
      // 必须拼起整个响应块再找。
      const head = (r.lines || []).join(' ');
      const m = /\bUID\s+(\d+)/i.exec(head);
      if (m && r.literals.length) out.set(Number(m[1]), r.literals[0]);
    }
    return out;
  }

  async logout() {
    try {
      if (!this.closed) await this.command('LOGOUT');
    } catch { /* 忽略 */ }
    try { this.sock.end(); } catch { /* 忽略 */ }
  }
}

/* ------------------------------------------------------------------ */
/* MIME 解析                                                           */
/* ------------------------------------------------------------------ */

const CHARSET_ALIAS = {
  'utf8': 'utf-8', 'utf-8': 'utf-8', 'utf-8;': 'utf-8',
  'gb2312': 'gbk', 'gbk': 'gbk', 'gb18030': 'gb18030', 'cp936': 'gbk',
  'big5': 'big5', 'iso-8859-1': 'latin1', 'latin1': 'latin1',
  'us-ascii': 'utf-8', 'ascii': 'utf-8', 'windows-1252': 'windows-1252',
  'iso-2022-jp': 'iso-2022-jp', 'shift_jis': 'shift_jis', 'koi8-r': 'koi8-r',
};

function decodeBytes(buf, charset) {
  let cs = String(charset || 'utf-8').toLowerCase().split('*')[0].trim().replace(/^["']|["']$/g, '');
  cs = CHARSET_ALIAS[cs] || cs;
  try { return new TextDecoder(cs).decode(buf); } catch { /* 回落 */ }
  try { return buf.toString('utf8'); } catch { return ''; }
}

function splitParams(value) {
  const out = [];
  let cur = '';
  let q = false;
  for (const ch of value) {
    if (ch === '"') { q = !q; cur += ch; continue; }
    if (ch === ';' && !q) { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

function parseStruct(value) {
  const parts = splitParams(value || '');
  const main = (parts.shift() || '').trim().toLowerCase();
  const params = {};
  for (const p of parts) {
    const i = p.indexOf('=');
    if (i < 0) continue;
    const k = p.slice(0, i).trim().toLowerCase();
    let v = p.slice(i + 1).trim().replace(/^"|"$/g, '');
    params[k] = v;
  }
  return { main, params };
}

function parseHeaders(headText) {
  const headers = {};
  let key = null;
  for (const line of headText.split(/\r?\n/)) {
    if (/^[ \t]/.test(line) && key) { headers[key] += ' ' + line.trim(); continue; }
    const i = line.indexOf(':');
    if (i < 0) continue;
    key = line.slice(0, i).trim().toLowerCase();
    headers[key] = line.slice(i + 1).trim();
  }
  return headers;
}

function parseEntity(buf) {
  const i = buf.indexOf('\r\n\r\n');
  const j = i < 0 ? buf.indexOf('\n\n') : -1;
  if (i >= 0) return { headers: parseHeaders(buf.subarray(0, i).toString('latin1')), body: buf.subarray(i + 4) };
  if (j >= 0) return { headers: parseHeaders(buf.subarray(0, j).toString('latin1')), body: buf.subarray(j + 2) };
  return { headers: parseHeaders(buf.toString('latin1')), body: Buffer.alloc(0) };
}

function decodeTransfer(buf, enc) {
  const e = String(enc || '').toLowerCase().trim();
  if (e === 'base64') {
    try {
      const clean = buf.toString('latin1').replace(/[^A-Za-z0-9+/=]/g, '');
      return Buffer.from(clean, 'base64');
    } catch { return buf; }
  }
  if (e === 'quoted-printable') {
    const s = buf.toString('latin1');
    const out = [];
    for (let i = 0; i < s.length; i++) {
      const c = s[i];
      if (c === '=') {
        if (s[i + 1] === '\r' && s[i + 2] === '\n') { i += 2; continue; }
        if (s[i + 1] === '\n') { i += 1; continue; }
        const hex = s.substr(i + 1, 2);
        if (/^[0-9A-Fa-f]{2}$/.test(hex)) { out.push(parseInt(hex, 16)); i += 2; continue; }
      }
      out.push(c.charCodeAt(0) & 0xff);
    }
    return Buffer.from(out);
  }
  return buf;
}

function splitMultipart(body, boundary) {
  const s = body.toString('latin1');
  const mark = '--' + boundary;
  const chunks = s.split(mark);
  const parts = [];
  for (let i = 1; i < chunks.length; i++) {
    let c = chunks[i];
    if (c.startsWith('--')) break;
    if (c.startsWith('\r\n')) c = c.slice(2);
    else if (c.startsWith('\n')) c = c.slice(1);
    if (c.endsWith('\r\n')) c = c.slice(0, -2);
    else if (c.endsWith('\n')) c = c.slice(0, -1);
    parts.push(Buffer.from(c, 'latin1'));
  }
  return parts;
}

function htmlToText(html) {
  return html
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|h[1-6]|li|table)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#x([0-9a-f]+);/gi, (_m, h) => {
      try { return String.fromCodePoint(parseInt(h, 16)); } catch { return ''; }
    })
    .replace(/&#(\d+);/g, (_m, d) => {
      try { return String.fromCodePoint(Number(d)); } catch { return ''; }
    })
    .replace(/&amp;/gi, '&')
    .replace(/[ \t\u00a0]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .split('\n').map((l) => l.trim()).filter(Boolean).join('\n');
}

function collectTextParts(entity, out, depth) {
  if (depth > 8) return;
  const ct = parseStruct(entity.headers['content-type'] || 'text/plain');
  if (ct.main.startsWith('multipart/') && ct.params.boundary) {
    for (const seg of splitMultipart(entity.body, ct.params.boundary)) {
      collectTextParts(parseEntity(seg), out, depth + 1);
    }
    return;
  }
  if (ct.main === 'text/plain' || ct.main === 'text/html' || ct.main === '') {
    const bytes = decodeTransfer(entity.body, entity.headers['content-transfer-encoding']);
    const text = decodeBytes(bytes, ct.params.charset);
    out.push({ type: ct.main || 'text/plain', text: ct.main === 'text/html' ? htmlToText(text) : text });
  }
}

function decodeRFC2047(raw) {
  if (!raw) return '';
  return String(raw).replace(/=\?([^?]+)\?([BbQq])\?([^?]*)\?=/g, (_m, charset, enc, data) => {
    try {
      let buf;
      if (enc.toUpperCase() === 'B') {
        buf = Buffer.from(data, 'base64');
      } else {
        const bytes = [];
        const s = data.replace(/_/g, ' ');
        for (let i = 0; i < s.length; i++) {
          if (s[i] === '=' && /^[0-9A-Fa-f]{2}$/.test(s.substr(i + 1, 2))) {
            bytes.push(parseInt(s.substr(i + 1, 2), 16));
            i += 2;
          } else {
            bytes.push(s.charCodeAt(i) & 0xff);
          }
        }
        buf = Buffer.from(bytes);
      }
      return decodeBytes(buf, charset);
    } catch { return ''; }
  }).trim();
}

/* ------------------------------------------------------------------ */
/* 验证码 / 链接提取                                                   */
/* ------------------------------------------------------------------ */

const CODE_KEYWORDS = [
  '验证码', '校验码', '动态码', '确认码', '安全码', '激活码', '登录码', '一次性密码', '密码',
  'verification code', 'verification', 'security code', 'login code', 'code', 'passcode', 'password', 'otp', 'pin',
].join('|');

function extractCodes(subject, text) {
  const hay = `${subject || ''}\n${text || ''}`;
  const scores = new Map();
  const add = (code, score) => {
    if (!code) return;
    const c = String(code).trim();
    if (!c) return;
    // 排除年份: 2026 这类几乎不会是验证码
    if (/^(19|20)\d{2}$/.test(c)) return;
    if (!scores.has(c) || scores.get(c) < score) scores.set(c, score);
  };
  const push = (re, score, needDigit) => {
    for (const m of hay.matchAll(re)) {
      const c = m[1];
      if (needDigit && !/\d/.test(c)) continue; // 纯字母不算验证码
      add(c, score);
    }
  };

  // 1) 关键词邻近距离内的数字验证码 —— 最可靠
  push(new RegExp(`(?:${CODE_KEYWORDS})[^0-9A-Za-z]{0,30}([0-9]{4,8})`, 'gi'), 100);
  // 2) 关键词邻近距离内的字母数字混合验证码 (必须含数字)
  push(new RegExp(`(?:${CODE_KEYWORDS})[^0-9A-Za-z]{0,30}([A-Z0-9]{5,8})`, 'gi'), 90, true);
  // 3) 独立成词的 6 位数字 (邮箱验证码最常见形态)
  push(/(?<![\w.,/:-])(\d{6})(?![\w.,/-])/g, 60);
  // 4) 独立成词的 5/7/8 位数字
  push(/(?<![\w.,/:-])(\d{5,8})(?![\w.,/-])/g, 35);

  return [...scores.entries()]
    .filter(([, sc]) => sc >= 50)
    .sort((a, b) => b[1] - a[1] || a[0].length - b[0].length)
    .slice(0, 5)
    .map(([code]) => code);
}

function extractLinks(text) {
  const out = [];
  const seen = new Set();
  for (const m of String(text || '').matchAll(/https?:\/\/[^\s"'<>)\]}]+/gi)) {
    let u = m[0].replace(/[.,;]+$/, '');
    if (seen.has(u)) continue;
    seen.add(u);
    out.push(u);
    if (out.length >= 3) break;
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* 展示                                                                */
/* ------------------------------------------------------------------ */

function oneLine(s) {
  return String(s || '').replace(/\s+/g, ' ').trim();
}

function formatMail(entity, uid, folder) {
  const h = entity.headers;
  const subject = oneLine(decodeRFC2047(h.subject));
  const from = oneLine(decodeRFC2047(h.from));
  const dateRaw = h.date || '';
  let date = dateRaw;
  const d = new Date(dateRaw);
  if (!Number.isNaN(d.getTime())) {
    const p = (n) => String(n).padStart(2, '0');
    date = `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  }

  const parts = [];
  collectTextParts(entity, parts, 0);
  const plain = parts.filter((p) => p.type === 'text/plain').map((p) => p.text).join('\n');
  const html = parts.filter((p) => p.type === 'text/html').map((p) => p.text).join('\n');
  let body = (plain || html || '').trim();

  const codes = extractCodes(subject, body);
  const links = extractLinks(body);

  return {
    uid: Number(uid),
    folder: folder || '',
    from, subject, date, body, codes, links,
    ts: Number.isNaN(d.getTime()) ? 0 : d.getTime(),
  };
}

function printMail(mail, index, opts) {
  const line = '─'.repeat(58);
  const bodyCap = opts.raw ? body_full(mail.body) : body_snip(mail.body);
  const out = [];
  out.push(`【${index}】[${mail.folder || '-'}] UID ${mail.uid}   ${mail.date}`);
  out.push(`发件人 : ${mail.from || '(未知)'}`);
  out.push(`主  题 : ${mail.subject || '(无主题)'}`);
  if (mail.codes.length) {
    out.push(`验证码 : ${mail.codes.join('  /  ')}`);
  } else {
    out.push('验证码 : (未识别)');
  }
  if (mail.links.length) out.push(`链  接 : ${mail.links[0]}`);
  out.push('正  文 :');
  out.push(bodyCap || '(空)');
  out.push(line);
  process.stdout.write(out.join('\n') + '\n');
}

function body_full(s) { return (s || '').trim(); }
function body_snip(s) {
  const t = (s || '').trim();
  if (!t) return '';
  return t.length > 500 ? t.slice(0, 500) + ' …' : t;
}

/* ------------------------------------------------------------------ */
/* 主流程                                                              */
/* ------------------------------------------------------------------ */

async function pull(conn, folder, uids) {
  const mails = [];
  const batch = await conn.fetchRawBatch(uids);
  for (const uid of uids) {
    const raw = batch.get(uid);
    if (!raw) continue;
    mails.push(formatMail(parseEntity(raw), uid, folder));
  }
  return mails;
}

const DEFAULT_SCAN = ['INBOX', 'Junk'];

/** 验证码邮件常被判成垃圾邮件, 所以默认两个文件夹一起扫 */
function resolveFolders(opts) {
  if (opts.mailbox) return [opts.mailbox];
  if (opts.scan && opts.scan.length) return opts.scan;
  return DEFAULT_SCAN;
}

async function runOnce(conn, acct, opts) {
  const folders = resolveFolders(opts);
  const stats = [];
  const mails = [];
  let total = 0;

  for (const name of folders) {
    try {
      await conn.select(name);
    } catch (e) {
      if (folders.length > 1) { log(`[警告] 打不开文件夹 "${name}", 跳过: ${e.message}`); continue; }
      throw e;
    }
    const ids = await conn.search(opts.unseen ? 'UNSEEN' : 'ALL');
    stats.push({ name, count: ids.length });
    total += ids.length;
    // 注意: FETCH 依赖当前 SELECT 的文件夹, 必须在这个循环里取完
    const got = await pull(conn, name, ids.slice(-opts.limit));
    mails.push(...got);
  }

  if (!mails.length) {
    const where = folders.join(' + ');
    log(`[3/3] ${where} 里没有${opts.unseen ? '未读' : ''}邮件`);
    if (opts.json) {
      process.stdout.write(JSON.stringify({
        account: acct.email, mailbox: folders.join('+'), folders: stats, total: 0, mails: [],
      }, null, 2) + '\n');
    }
    return { mails: [], total };
  }

  mails.sort((a, b) => b.ts - a.ts);      // 跨文件夹按时间倒序
  const top = mails.slice(0, opts.limit);

  log(`[3/3] ${stats.map((s) => `${s.name} ${s.count} 封`).join(' + ')}  → 读取最新 ${top.length} 封`);

  if (opts.json) {
    process.stdout.write(JSON.stringify({
      account: acct.email,
      mailbox: folders.join('+'),
      folders: stats,
      total,
      mails: top.map((m) => ({
        folder: m.folder, uid: m.uid, date: m.date, from: m.from, subject: m.subject,
        codes: m.codes, links: m.links, body: body_snip(m.body),
      })),
    }, null, 2) + '\n');
    return { mails: top, total };
  }

  if (opts.codeOnly) {
    const out = [];
    for (const m of top) {
      if (m.codes.length) out.push(m.codes[0]);
      else if (m.links.length) out.push('LINK:' + m.links[0]);
    }
    process.stdout.write(out.join('\n') + (out.length ? '\n' : ''));
    return { mails: top, total };
  }

  const summary = stats.filter((s) => s.count > 0).map((s) => `${s.name} ${s.count} 封`).join(' + ');
  process.stdout.write(`\n${'='.repeat(58)}\n`);
  process.stdout.write(`账号   : ${acct.email}\n`);
  process.stdout.write(`文件夹 : ${summary || folders.join(' + ')}   (共 ${total} 封)\n`);
  process.stdout.write(`${'='.repeat(58)}\n\n`);
  top.forEach((m, i) => printMail(m, i + 1, opts));
  return { mails: top, total };
}

async function watchLoop(conn, acct, opts) {
  const folders = resolveFolders(opts);
  const lastUid = new Map();

  for (const name of folders) {
    try {
      await conn.select(name);
      const ids = await conn.search('ALL');
      lastUid.set(name, ids.length ? Math.max(...ids) : 0);
    } catch (e) {
      log(`[警告] 打不开文件夹 "${name}": ${e.message}`);
      lastUid.set(name, 0);
    }
  }
  const base = [...lastUid.entries()].map(([n, u]) => `${n}=${u}`).join(', ');
  log(`[3/3] 开始轮询 ${folders.join(' + ')} (当前 ${base}, 间隔 ${opts.watch}s, Ctrl+C 退出)`);

  let stopping = false;
  process.on('SIGINT', () => {
    if (stopping) process.exit(130);
    stopping = true;
    log('\n收到 Ctrl+C, 正在退出 ...');
  });

  for (;;) {
    if (stopping) break;
    await new Promise((r) => setTimeout(r, opts.watch * 1000));
    if (stopping) break;

    let gotCode = false;
    for (const name of folders) {
      if (stopping) break;
      const since = lastUid.get(name) || 0;
      let fresh;
      try {
        await conn.select(name);
        const ids = await conn.search(`UID ${since + 1}:*`);
        fresh = ids.filter((u) => u > since).sort((a, b) => a - b);
      } catch (e) {
        log(`[错误] 轮询 "${name}" 失败: ${e.message}`);
        continue;
      }
      if (!fresh.length) continue;
      lastUid.set(name, Math.max(since, ...fresh));

      const mails = await pull(conn, name, fresh);
      for (const m of mails) {
        const stamp = new Date().toLocaleTimeString('zh-CN', { hour12: false });
        process.stdout.write(`\n[${stamp}] 新邮件!  (${name})\n`);
        printMail(m, '新', opts);
        if (m.codes.length) { gotCode = true; log(`>>> 验证码: ${m.codes.join(' / ')}`); }
      }
    }
    if (gotCode && opts.exitOnCode) {
      log('已取到验证码, 按 --exit-on-code 退出。');
      return true;
    }
  }
  return false;
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  if (opts.help) { printHelp(); return; }

  const acct = loadAccount(opts.account);
  const token = await getAccessToken(acct, opts);

  log(`[2/3] 连接 ${IMAP_HOST}:${IMAP_PORT} 并做 XOAUTH2 认证 ...`);
  const conn = new ImapClient(IMAP_HOST, IMAP_PORT);
  await conn.connect();
  try {
    await conn.greeting();
    await conn.login(acct.email, token);

    if (opts.folders) {
      const list = await conn.listFolders();
      process.stdout.write(list.map((f) => `  ${f}`).join('\n') + '\n');
      return;
    }

    if (opts.watch > 0) {
      await watchLoop(conn, acct, opts);
    } else {
      await runOnce(conn, acct, opts);
    }
  } finally {
    await conn.logout();
  }
}

main().catch((e) => {
  log('\n[失败] ' + (e && e.message ? e.message : String(e)));
  if (process.env.DEBUG && e && e.stack) log(e.stack);
  process.exitCode = 1;
});
