## v1.0.0 (2025-12-17)
- 新增：定时/多次发送（间隔 + 次数，Start/Stop）
- 增强：To/Cc/Bcc 解析与日志输出（包含实际 RCPT TO 列表）
- 增强：SMTP 协议调试日志（带方向标记，自动合并重复噪声；AUTH 始终脱敏，DATA 可选脱敏）
- 修复：认证模式切换时的默认值恢复（OAuth2 默认 587 + 正确 SMTP Host；Anonymous 默认 25 + 入站 Host）
- 优化：21V 匿名模式的 Host 推导与 MX 查询（异步 + 缓存，避免 UI 卡顿/弹窗）
- 发布：提供 onefile（单 EXE）与 onedir（启动更快）两种构建；高分屏 DPI 与应用图标优化

## v0.2.19 (2025-12-16)
- 变更：移除“一键配置/一键注册”相关功能（Graph/浏览器登录/证书/PFX/PowerShell 自动化）
- 变更：OAuth2 仅保留 Client Credentials（手动输入 Tenant ID / Client ID / Client Secret）
- 新增：获取 token 后在日志输出 `Token acquired (MSAL): ...`（默认脱敏；勾选 Save token 才输出完整 token）
- 修复：切换 Cloud 时自动使用对应的默认 SMTP Host 与 Scope（Global / 21V）

## v0.2.18 (2025-12-16)
- 新增：Global / 21V 配置分开保存（config.json -> profiles），切换 Cloud 时自动保存当前并加载目标 Cloud 的配置
- 修复：21V 一键配置浏览器登录不再固定使用 Graph PowerShell Client ID，改为自动尝试多个内置 Client ID，避免 AADSTS700016
- 说明：如需强制指定登录 Client ID，可在 config.json 中为对应 Cloud 的 profile 增加 `one_click_browser_client_id`

## v0.2.17 (2025-12-16)
- 修复：EXO 侧获取 Service Principal 改为按微软文档使用 `Get-ServicePrincipal -Identity <DisplayName>`，并在创建后自动重试等待同步
- 增强：PowerShell 日志输出 ExchangeOnlineManagement 版本及关键 cmdlet 是否存在，便于定位“模块/连接方式不支持”问题

## v0.2.16 (2025-12-16)
- 修复：避免 AzureAD `Get-AzureADServicePrincipal -SearchString` 命中旧对象导致注册到错误的 App（仅当 AppId/ObjectId 与本次创建一致才接受）
- 修复：`New-ServicePrincipal` 的 DisplayName 增加 AppId，确保唯一且便于 `Get-ServicePrincipal` 排查
- 修复：`New-ServicePrincipal` 失败不再无条件“跳过”，除非明确是“已存在”类错误，否则直接报错提示权限/连接问题

## v0.2.15 (2025-12-16)
- 修复：按微软文档要求在 EXO 注册 Service Principal 时强制使用 `New-ServicePrincipal -AppId ... -ObjectId ... -Organization <tenantId>`，提升在部分租户/环境下的兼容性

## v0.2.14 (2025-12-16)
- 增强：一键配置日志输出 token 的 tid / 登录用户，避免在错误租户里找 App
- 增强：创建完成后明确输出 AAD Application objectId + Enterprise App(AAD Service Principal) objectId
- 修复：EXO 侧 TrusteeId 自动探测（依次尝试 ExternalDirectoryObjectId/ObjectId/Identity/AAD SP ObjectId，哪个能成功 Add-RecipientPermission 就用哪个）

## v0.2.13 (2025-12-16)
- 修复：一键配置现在会先检测现有配置；如果已存在旧 App/证书，会先删除再重新配置（而不是创建后再删除）
- 新增：EXO PowerShell 自动化会解析 EXO 中的 ServicePrincipal，并输出/回填真正用于 `Add-RecipientPermission -Trustee ...` 的 ID

## v0.2.12 (2025-12-16)
- 新增：一键生成成功后自动清理旧的 Entra App / AAD Service Principal，并删除旧的本地 PFX 证书文件（避免残留导致混用）
- 优化：PowerShell 自动化中按你指定的命令形态创建 EXO Service Principal；若检测到 AzureAD 模块且已连接，会优先尝试 `Get-AzureADServicePrincipal -SearchString ...`，失败则自动回退到已知 AppId/ObjectId

## v0.2.11 (2025-12-16)
- 修复：选择 21V (China) 时，Settings 里的 OAuth2 Scope 默认值自动切换为 `https://partner.outlook.cn/...`
- 修复：连接 EXO PowerShell 的 21V 环境参数更正为 `-ExchangeEnvironmentName O365China`
- 优化：EXO 端创建 Service Principal 使用 `$AADServicePrincipalDetails` + `New-ServicePrincipal -ObjectId ...` 的命令形态（无需安装 AzureAD 模块）

## v0.2.10 (2025-12-16)
- 修复：PowerShell 自动化脚本支持 21V (China) 云环境（自动指定 `-ExchangeEnvironmentName AzureChinaCloud`）
- 优化：PowerShell 脚本逻辑，明确使用 `New-ServicePrincipal` 注册 SP，无需依赖 AzureAD 模块

## v0.2.9 (2025-12-16)
- 增强：一键配置 PowerShell 自动化脚本增加 `New-ServicePrincipal` 步骤
- 说明：此步骤用于在 Exchange Online 中注册 Service Principal，确保 `Add-RecipientPermission` 能正确识别 App

## v0.2.8 (2025-12-16)
- 新增：一键配置流程中自动调用 PowerShell 授予 SendAs 权限
- 功能：自动检测并安装 `ExchangeOnlineManagement` 模块（如果缺失）
- 优化：使用当前登录用户的 Token 连接 EXO PowerShell，无需额外配置
- 修复：一键配置时若未填写发件人邮箱，会跳过 PowerShell 授权步骤并提示

## v0.2.7 (2025-12-16)
- 优化：UI 重构为多标签页 (Tabs)，解决日志窗口过小问题
- 优化：增加 DPI 缩放支持，高分屏显示更清晰
- 优化：一键配置流程，证书模式下不再尝试授予 Delegated 权限（消除警告）
- 新增：日志页面增加 "Copy PowerShell Permission Command" 按钮，方便复制授权命令

## v0.1.0 (2025-12-16)
- 初始版本：SMTP 发送工具（GUI）
- 支持认证：匿名 / Basic / OAuth2(XOAUTH2)
- 支持云环境：Global / 21V(世纪互联)
- 支持打包：PyInstaller 生成 Windows exe

## v0.2.0 (2025-12-16)
- 修复：切换 Cloud 到 21V 时，SMTP Host 会自动更新（除非你自定义过 Host）
- 变更：移除 Device Code flow；OAuth2 仅保留 Client Credentials 和 Auth Code
- 新增：一键配置（浏览器交互登录 + Graph API 调用，创建应用 + 授权 SMTP 权限 + 回填本地配置，支持 client secret / 证书，证书会导出为 PFX）
- 新增：运行时支持使用证书（PFX）获取 token（新增依赖 cryptography）

## v0.2.1 (2025-12-16)
- 变更：一键配置不再使用 PowerShell，改为 Python 直接打开浏览器登录并调用 Graph API
- 新增依赖：requests / azure-identity（用于一键配置 Graph 调用与交互式登录）

## v0.2.2 (2025-12-16)
- 修复：部分租户一键配置报 AADSTS65002（允许配置 Browser Login Client ID，默认使用 Graph PowerShell client id）

## v0.2.3 (2025-12-16)
- 优化：一键配置界面简化，移除 Browser Login Client ID 输入框（对齐 UniversalEmailCleaner）
- 优化：一键配置逻辑重构，提供更详细的日志输出
- 修复：内置使用 Graph PowerShell Client ID 进行配置登录，彻底解决 AADSTS65002 问题

## v0.2.4 (2025-12-16)
- 修复：一键配置运行时报错 `unexpected keyword argument 'graph_endpoint'` 的问题

## v0.2.5 (2025-12-16)
- 优化：日志窗口高度增加，并支持捕获 SMTP 调试日志（smtplib debug output）
- 优化：一键配置完成后，提示用户必须手动执行 Exchange PowerShell `Add-RecipientPermission` 命令
- 优化：一键配置中 `SMTP.Send` 授权失败降级为警告（Client Credentials 模式不需要此权限）

## v0.2.6 (2025-12-16)
- 优化：UI 布局重构，使用 PanedWindow（上下分割），支持拖动调整日志窗口大小
- 优化：设置区域增加滚动条，防止在小屏幕上显示不全


