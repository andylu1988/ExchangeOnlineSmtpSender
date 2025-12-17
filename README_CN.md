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

## 详细配置步骤（Entra 应用 + Exchange Online）

本工具的 SMTP OAuth2 基于 **Client Credentials**（应用权限 / app-only）。下面步骤参考微软针对 IMAP/POP/SMTP OAuth 的官方说明整理（已用自己的话重新组织，便于照做）：

参考链接：
https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth

### 1）在 Microsoft Entra ID 注册应用

1. 打开 **Microsoft Entra 管理中心** → **应用注册(App registrations)** → **新注册(New registration)**。
2. 填写应用名称。一般单租户场景选择 **仅此组织目录中的帐户**。
3. 创建后记录：
	- **Tenant ID**（Directory ID）
	- **Client ID**（Application ID）

### 2）创建 Client Secret

1. 进入应用 → **证书和密码(Certificates & secrets)** → **新建客户端密码(New client secret)**。
2. 立即复制保存 secret 的 **Value**（之后无法再次查看）。

### 3）添加 SMTP 的应用权限并授予管理员同意

1. 进入应用 → **API 权限(API permissions)** → **添加权限(Add a permission)**。
2. 选择 **我的组织使用的 API(APIs my organization uses)** → 搜索 **Office 365 Exchange Online**。
3. 选择 **应用程序权限(Application permissions)**。
4. SMTP 请选择并添加 **SMTP.SendAsApp**。
5. 点击 **授予管理员同意(Grant admin consent)**。

### 4）在 Exchange Online 注册 Service Principal

完成管理员同意后，需要用 Exchange Online PowerShell 把该应用的 SP 注册到 Exchange 中。

1. 安装并连接：

```powershell
Install-Module -Name ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -Organization <tenantId>
```

2. 注册到 Exchange：

```powershell
New-ServicePrincipal -AppId <APPLICATION_ID> -ObjectId <OBJECT_ID>
```

注意：
- `<APPLICATION_ID>` 是应用注册里的 **Client ID**。
- `<OBJECT_ID>` 必须使用 **企业应用(Enterprise applications)** 里该应用实例的 Object ID（不是“应用注册(App registrations)”页面那个 Object ID）。用错通常会导致认证失败。

### 5）给发件人邮箱授予 SendAs

使用 client credentials 以某个邮箱作为 From 发送时，需要对该邮箱授予 SendAs：

```powershell
Add-RecipientPermission -Identity "sender@contoso.com" -Trustee <SERVICE_PRINCIPAL_ID> -AccessRights SendAs
```

### 6）在工具里填写参数

在界面里：

- 认证方式选择 **OAuth2**
- 填写 **Tenant ID / Client ID / Client Secret**
- SMTP Host/Port 建议保持工具默认（Global 一般是 **587 + STARTTLS**）
- SMTP 的 client credentials scope 一般使用 **`https://outlook.office365.com/.default`**（工具会按 Global/21V 自动切默认值）

### 21V（世纪互联）差异点

如果你的租户/邮箱在 **Office 365 中国（世纪互联 / 21V）**，以下参数与 Global 不同：

- Authority Host：`https://login.chinacloudapi.cn`（不是 `login.microsoftonline.com`）
- SMTP Host（OAuth2/Basic）：`smtp.partner.outlook.cn`
- SMTP OAuth2 Scope（client credentials）：`https://partner.outlook.cn/.default`
- 匿名入站/EOP Host 后缀：`mail.protection.partner.outlook.cn`

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
