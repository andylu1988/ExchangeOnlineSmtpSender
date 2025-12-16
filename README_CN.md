# Exchange Online SMTP 发信工具（GUI）

一个 Windows 友好的图形界面工具，用于通过 SMTP 给 Exchange Online / Microsoft 365 发送测试邮件。

## 功能

- SMTP 发送：**匿名** / **Basic** / **OAuth2 (XOAUTH2)**
- 云环境：**Global** / **21V(世纪互联)**
- Tkinter 图形界面
- 支持 PyInstaller 打包独立 `.exe`

## 注意事项（Exchange Online）

- 很多租户已限制/禁用 SMTP Basic Auth，建议优先使用 OAuth2。
- SMTP OAuth2 使用 SASL XOAUTH2，需要租户/应用配置允许。

## 安装

```powershell
pip install -r requirements.txt
```

## 运行

```powershell
python smtp_sender_gui.py
```

## 打包 EXE

```powershell
pyinstaller --clean --noconfirm ExchangeOnlineSmtpSender.spec
```
