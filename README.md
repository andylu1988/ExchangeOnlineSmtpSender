# Exchange Online SMTP Sender (GUI)

A Windows-friendly GUI tool to send test emails via SMTP against Exchange Online / Microsoft 365.

[中文说明](README_CN.md)

## Download

Get the latest builds from GitHub Releases:

- `ExchangeOnlineSmtpSender_vX.Y.Z.exe` (onefile, single EXE)
- `ExchangeOnlineSmtpSender_vX.Y.Z_onedir.zip` (onedir, faster startup)

## Features

- SMTP send with **Anonymous**, **Basic Auth**, or **OAuth2 (XOAUTH2)**
- Cloud environments: **Global** and **China (21Vianet)**
- Multi-recipient support: **To / Cc / Bcc** (supports `;` and `,` separators)
- Scheduled multi-send (interval + repeat count, Start/Stop)
- Detailed logs:
	- UI app log (To/Cc/Bcc + envelope recipients)
	- Optional SMTP protocol transcript with direction prefixes and de-duplicated noise
- Windows UX: DPI-aware, custom icon

## Notes (Exchange Online)

- Basic Auth for SMTP is being restricted/disabled in many tenants. OAuth2 is recommended.
- OAuth2 support in SMTP uses SASL XOAUTH2; your tenant/app must allow it.

## OAuth2

This tool supports OAuth2 using **Client Credentials (app-only)** with **Client Secret**.

By default, tokens are **redacted** in logs. Only enable **“Save Authorization token in log”** when you explicitly need the raw token for debugging.

## Run from source

```powershell
pip install -r requirements.txt
python .\smtp_sender_gui.py
```

## Build EXE (PyInstaller)

```powershell
pyinstaller --clean --noconfirm .\ExchangeOnlineSmtpSender.spec
pyinstaller --clean --noconfirm .\ExchangeOnlineSmtpSender_onedir.spec
```

## License

MIT
