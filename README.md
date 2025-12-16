# Exchange Online SMTP Sender (GUI)

A small Windows-friendly GUI tool to send test emails via SMTP against Exchange Online / Microsoft 365.

## Features

- SMTP send with **Anonymous**, **Basic Auth**, or **OAuth2 (XOAUTH2)**
- Cloud environments: **Global** and **China (21Vianet)**
- GUI built with Tkinter
- Build standalone `.exe` with PyInstaller

## Notes (Exchange Online)

- Basic Auth for SMTP is being restricted/disabled in many tenants. OAuth2 is recommended.
- OAuth2 support in SMTP uses SASL XOAUTH2; your tenant/app must allow it.

## Install

```powershell
pip install -r requirements.txt
```

## Run

```powershell
python smtp_sender_gui.py
```

## Build EXE

```powershell
pyinstaller --clean --noconfirm ExchangeOnlineSmtpSender.spec
```
