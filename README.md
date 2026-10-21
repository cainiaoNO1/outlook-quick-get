# Outlook 快速取件 (outlook-quick-get)

1:1 复刻 https://mail.chatai.codes/ 邮件快速取件页面，同时支持直接在命令行/批处理中使用。

## 特性亮点

1. **双协议并行取件**：支持 IMAP XOAUTH2 + Microsoft Graph API 双协议并行取件与智能内容指纹去重。
2. **零依赖极速后端**：基于 Node.js/Python 双引擎，支持批量 `UID FETCH`（取 10 封仅需约 2 秒）。
3. **过滤条件云端保存**：支持按标题、正文自定义过滤规则，规则持久化保存于服务端 `filters.json`，随时勾选组合生效。
4. **验证码智能识别**：自动从邮件正文、标题提取 4-8 位验证码与确认链接。
5. **支持最少取 1 封**：按需可选 1 / 3 / 5 / 10 / 20 / 30 封（按账号独立计数）。

## 项目结构

```
outlook-quick-get/
├── server.py              # Web 服务端核心（静态资源分发 + 4 个 API 路由）
├── config.example.json    # 服务端配置模板
├── www/                   # 纯白蓝 macOS 玻璃风前端
│   ├── index.html
│   ├── css/style.css
│   └── js/app.js
└── tool/                  # 桌面端独立取件工具（可独立使用）
    ├── outlook_mail.js    # 零依赖 Node.js 取件脚本（支持批量 FETCH 优化）
    ├── outlook_mail.py    # Python 取件脚本
    ├── 取件.bat           # 双击快速取件
    ├── 守验证码.bat       # 15s 轮询取到验证码自动退出
    ├── account.example.txt# 账号配置文件格式
    └── 使用说明.md
```

## 快速运行 Web 版

```bash
# 启动 Web 服务（默认监听 0.0.0.0:4000）
python server.py
```
浏览器访问：`http://localhost:4000` 或 `http://服务器IP:4000`。

## 快速运行命令行版

进入 `tool/` 目录：
1. 复制 `account.example.txt` 为 `account.txt` 并填入邮箱密钥：
   ```
   邮箱----密码----client_id----refresh_token
   ```
2. 执行取件：
   ```bash
   node outlook_mail.js -n 5
   # 或在 Windows 双击「取件.bat」
   ```
