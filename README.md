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

## Configure the Entra app (OAuth2 SMTP AUTH)

The tool implements OAuth2 for SMTP AUTH using the **client credentials** flow. The high-level steps below are based on Microsoft guidance for OAuth with IMAP/POP/SMTP.

Reference: https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth

### 1) Register an app in Microsoft Entra ID

1. Go to **Microsoft Entra admin center** → **App registrations** → **New registration**.
2. Name your app (any name is fine). For a single-tenant setup, choose **Accounts in this organizational directory only**.
3. Create the app and record:
	- **Tenant ID** (Directory ID)
	- **Client ID** (Application ID)

### 2) Create a Client Secret

1. In the app, go to **Certificates & secrets** → **New client secret**.
2. Copy the secret **Value** immediately (you won’t be able to view it again).

### 3) Add the SMTP application permission

1. In the app, go to **API permissions** → **Add a permission**.
2. Choose **APIs my organization uses** → search **Office 365 Exchange Online**.
3. Select **Application permissions**.
4. For SMTP access, add **SMTP.SendAsApp**.
5. Click **Grant admin consent** for your tenant.

### 4) Register the service principal in Exchange Online

After admin consent, register the app’s service principal in Exchange Online via PowerShell.

1. Install and connect:

```powershell
Install-Module -Name ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -Organization <tenantId>
```

2. Register in Exchange:

```powershell
New-ServicePrincipal -AppId <APPLICATION_ID> -ObjectId <OBJECT_ID>
```

Notes:
- The `<APPLICATION_ID>` is the **Client ID** from App registration.
- The `<OBJECT_ID>` is the **Object ID from the Enterprise application** instance (not the “App registrations” object id). Using the wrong object id typically causes auth failures.

### 5) Grant SendAs to the mailbox you will use as “From”

If you use client credentials and want to send as a specific mailbox, grant SendAs to the sender mailbox:

```powershell
Add-RecipientPermission -Identity "sender@contoso.com" -Trustee <SERVICE_PRINCIPAL_ID> -AccessRights SendAs
```

### 6) Fill in the tool

In the app UI:

- Auth: **OAuth2**
- Enter **Tenant ID / Client ID / Client Secret**
- Keep the default SMTP host/port (Global typically uses port **587** with STARTTLS)
- Scope for SMTP client credentials: use **`https://outlook.office365.com/.default`** (the tool sets defaults per cloud)

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
