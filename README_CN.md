# Exchange Online SMTP 发信工具（GUI）

一个 Windows 友好的图形界面工具，用于通过 SMTP 给 Exchange Online / Microsoft 365 发送测试邮件。

[English](README.md)

## 下载

建议直接从 GitHub Releases 下载：

- `ExchangeOnlineSmtpSender_vX.Y.Z.exe`（单文件版，只有一个 EXE）
- `ExchangeOnlineSmtpSender_vX.Y.Z_onedir.zip`（目录版，启动更快）

## 功能

- SMTP 发送：**匿名** / **Basic** / **OAuth2 (XOAUTH2)**
- 云环境：**Global** / **21V(世纪互联)**
- 多收件人：**To / Cc / Bcc**（支持 `;` 和 `,` 分隔）
- 定时/多次发送：支持间隔 + 次数，Start/Stop
- 日志增强：
	- UI 日志显示 To/Cc/Bcc + 实际 RCPT TO 列表
	- 可选保存 SMTP 协议调试日志（带方向标记，自动合并重复噪声）
- Windows 体验：支持高分屏 DPI、定制图标

## 注意事项（Exchange Online）

- 很多租户已限制/禁用 SMTP Basic Auth，建议优先使用 OAuth2。
- SMTP OAuth2 使用 SASL XOAUTH2，需要租户/应用配置允许。

## OAuth2 流程

本工具支持 OAuth2 的 **Client Credentials（应用权限 / app-only）**，通过 **Client Secret** 获取 token。

默认情况下日志会对 token 做脱敏；只有在你明确需要排查问题时，才建议勾选 **“Save Authorization token in log”** 输出完整 token。

## 从源码运行

```powershell
pip install -r requirements.txt
python .\smtp_sender_gui.py
```

## 打包 EXE（PyInstaller）

```powershell
pyinstaller --clean --noconfirm .\ExchangeOnlineSmtpSender.spec
pyinstaller --clean --noconfirm .\ExchangeOnlineSmtpSender_onedir.spec
```

## License

MIT
