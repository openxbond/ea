# XBOND 自动化部署脚本
# 执行命令 irm https://ea.sraph.com/install.ps1 | iex
# 老win7使用以下命令：
# certutil -urlcache -split -f https://ea.sraph.com/install.ps1 $env:TEMP\install.ps1
# iex ([IO.File]::ReadAllText("$env:TEMP\install.ps1"))
# 如果出现执行权限问题，请先运行：Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

$ErrorActionPreference = "Continue"

# ==================== 系统环境检测（最先执行）====================
$osVersion = [System.Environment]::OSVersion.Version
$isWin7 = ($osVersion.Major -eq 6 -and $osVersion.Minor -eq 1)
$psVersion = $PSVersionTable.PSVersion.Major

# ==================== TLS / 网络配置 ====================

# .NET ServicePointManager：强制 TLS 1.2+（对 WebClient/HttpWebRequest 生效）
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]'Tls12,Tls13' }
catch { try { [Net.ServicePointManager]::SecurityProtocol = 3072 } catch {} }  # 3072 = Tls12

# PSObject 构造函数：PS2.0（Win7 默认）中 [PSCustomObject]@{} 不可靠，统一用 Add-Member
function New-Obj {
    param([hashtable]$Props)
    $obj = New-Object PSObject
    foreach ($key in $Props.Keys) {
        $obj | Add-Member -MemberType NoteProperty -Name $key -Value $Props[$key] -Force
    }
    return $obj
}

# Get-I18n: Returns the localized string for a value.
# Accepts a plain string or a hashtable with CN/EN keys; falls back gracefully.
function Get-I18n {
    param($Val)
    if ($Val -is [hashtable]) {
        if ($LangChoice -eq '1' -and $Val.ContainsKey('CN')) { return $Val['CN'] }
        if ($Val.ContainsKey('EN')) { return $Val['EN'] }
        $keys = @($Val.Keys)
        if ($keys.Count -gt 0) { return $Val[$keys[0]] }
        return ''
    }
    if ($Val) { return $Val.ToString() }
    return ''
}

# ===================== 【配置中心 】 =====================
$Lang = @{
    CN = @{
        Welcome         = "🚀 欢迎使用 XBOND 自动化部署工具"
        LangSelect      = "请选择语言 | Select Language: 1=中文 2=English"
        SelectTip       = "请输入选项"
        CheckNetwork    = "🔍 检测网络连通性..."
        NetworkFail     = "❌ 网络连接失败或超时，但不影响本地继续执行"
        MainMenu        = "=== 主菜单 ==="
        Menu1           = "1. 全新安装"
        Menu2           = "2. 安装/更新EA"
        Menu3           = "3. 卸载账号实例"
        Menu4           = "4. 删除EA"
        Menu5           = "5. 退出脚本"
        DeleteEAList    = "=== 缓存中的EA文件 ==="
        DeleteEASelect  = "请选择要删除的EA（多选用逗号分隔）："
        NoEACache       = "❌ 缓存中没有EA文件"
        DeleteEADone    = "✅ 已删除：{0}"
        EAList          = "=== 可用EA（多选用逗号分隔 例:1,2）==="
        EASelect        = "请选择EA："
        EANoPlat        = "❌ 所选的EA没有共同支持的平台，请重新选择"
        PlatAuto        = "ℹ️ 所选EA仅共同支持 {0} 平台，自动选中"
        PlatSelect      = "💻 所选EA共同支持平台：1={0} 2={1}，请选择："
        ProviderList    = "=== 支持 {0} 平台的服务商 ==="
        ProviderSelect  = "请选择服务商（输入数字，仅单选）："
        ProviderAuto    = "ℹ️ 仅有服务商 [{0}] 支持 {1}，自动选中"
        ProviderNoPlat  = "❌ 当前没有服务商支持 {0} 平台"
        RegisterTip     = "🎉 若您还未注册交易账号，可通过以下邀请链接进行注册"
        RegisterUrl     = "🔗 注册链接："
        AccInput        = "📝 输入账号名称（空回车结束）"
        InstallStart    = "📥 开始部署基础环境 {0}..."
        CacheHit        = "✅ 缓存命中：{0} 已存在，跳过下载"
        DownloadTrader  = "📥 正在下载交易程序：{0}, 若提示安装，请选择否"
        DownloadEA      = "📥 下载最新EA文件：{0}"
        CopyEA          = "📋 批量拷贝EA到实例"
        DeployAcc       = "✅ 部署账号：{0}"
        UpdateStart     = "🔄 开始批量更新EA..."
        UninstallMenu   = "🗑️ 卸载模式：1=卸载单个账号 2=卸载全部 3=返回"
        Uninput         = "输入要卸载的账号名称："
        UnAllConfirm    = "⚠️ 确认卸载全部实例？(Y/N)"
        Success         = "🎉 操作执行完成！"
        Error           = "❌ 执行异常：{0}"
        Cleanup         = "🧹 已清理临时文件"
        InvalidInput    = "⚠️ 输入无效，请重新输入"
        SkipExist       = "ℹ️ {0} 已存在，跳过"
        OverWrite       = "⚠️ 目录已存在，是否覆盖？(Y/N)"
        LogPath         = "📄 日志文件：{0}"
        RiskConfirm     = "⚠️ 我已阅读并完全理解以上风险披露，输入 Y 确认继续，任意键退出"
        NewEATip        = "=== 安装新EA ==="
        EACopySuccess   = "✅ 新EA已安装到 {0} 账号实例"
        NoInstalledPlat = "❌ 未检测到已安装的MT4/MT5实例，请先全新安装"
        CacheOnlyTip    = "ℹ️ 未输入账号，EA文件已下载到缓存"
        RiskDisclosure  = @'
⚠️ 高风险投资警告：保证金交易具有高度风险，您可能损失全部甚至超过初始投入资金

风险披露与免责声明

01 资本风险声明
外汇（Forex）、贵金属、指数及差价合约（CFD）等保证金交易属于高风险金融投资行为，在不利市场条件下，您的损失可能超过初始投入资金。此类投资并不适合所有投资者，在进行任何交易前，请务必充分评估自身的财务状况、风险承受能力及投资目标。

02 EA 自动化交易风险
自动化交易系统依赖于交易平台、网络连接、服务器环境及程序代码的稳定性。用户需自行承担因网络延迟、VPS服务器宕机、交易平台异常、点差扩大、滑点、流动性不足、参数设置不当或程序错误所造成的全部交易风险。历史回测成绩不代表真实交易表现。

03 非投资建议声明
本软件所提供的所有内容，仅供学习、研究及技术交流用途，不构成任何形式的投资建议、交易指引或收益承诺。

04 非持牌与非代管理财声明
本软件提供方并非任何持牌金融机构，不从事代客理财、账户托管或资金管理服务。用户始终完全自主控制其交易账户及资金。

05 用户确认
用户在使用本软件前，应已充分阅读、理解并同意本风险披露与免责声明的全部内容。
'@
    }
    EN = @{
        Welcome         = "🚀 Welcome to XBOND Auto Deploy Tool"
        LangSelect      = "Select Language: 1=中文 2=English"
        SelectTip       = "Please enter your choice"
        CheckNetwork    = "🔍 Checking network connection..."
        NetworkFail     = "❌ Network connection failed or timed out"
        MainMenu        = "=== Main Menu ==="
        Menu1           = "1. Fresh Install"
        Menu2           = "2. Install/Update EA"
        Menu3           = "3. Uninstall Instance"
        Menu4           = "4. Delete EA"
        Menu5           = "5. Exit"
        DeleteEAList    = "=== Cached EA Files ==="
        DeleteEASelect  = "Select EA to delete (Multiple: 1,2):"
        NoEACache       = "❌ No EA files in cache"
        DeleteEADone    = "✅ Deleted: {0}"
        EAList          = "=== Available EAs (Multiple: 1,2) ==="
        EASelect        = "Select EAs:"
        EANoPlat        = "❌ Selected EAs have no common platform, please reselect"
        PlatAuto        = "ℹ️ Selected EAs only support {0}, auto selected"
        PlatSelect      = "💻 Common supported platforms: 1={0} 2={1}, please select:"
        ProviderList    = "=== Providers supporting {0} ==="
        ProviderSelect  = "Select Provider (Number):"
        ProviderAuto    = "ℹ️ Only provider [{0}] supports {1}, auto selected"
        ProviderNoPlat  = "❌ No provider supports {0} currently"
        RegisterTip     = "🎉 If you haven't registered a trading account yet, you can register via the following invitation link"
        RegisterUrl     = "🔗 Register URL: "
        AccInput        = "📝 Input Account (Empty Enter to Finish)"
        InstallStart    = "📥 Deploy Base Env {0}..."
        CacheHit        = "✅ Cache Hit: {0} Exists, Skip Download"
        DownloadTrader  = "📥 Downloading transaction program: {0}, if prompted to install, please select No"
        DownloadEA      = "📥 Download Latest EA: {0}"
        CopyEA          = "📋 Batch Copy EA to Instances"
        DeployAcc       = "✅ Deploy Account: {0}"
        UpdateStart     = "🔄 Batch Update EA Start..."
        UninstallMenu   = "🗑️ Uninstall:1=Single 2=All 3=Back"
        Uninput         = "Input Account to Uninstall:"
        UnAllConfirm    = "⚠️ Confirm Uninstall All? (Y/N)"
        Success         = "🎉 Operation Completed!"
        Error           = "❌ Error: {0}"
        Cleanup         = "🧹 Temporary Files Cleaned"
        InvalidInput    = "⚠️ Invalid Input"
        SkipExist       = "ℹ️ {0} Exists, Skip"
        OverWrite       = "⚠️ Directory Exists, Overwrite? (Y/N)"
        LogPath         = "📄 Log Path: {0}"
        RiskConfirm     = "⚠️ I have read and fully understand the risk disclosure, enter Y to continue, any key to exit"
        NewEATip        = "=== Install New EA ==="
        EACopySuccess   = "✅ New EA Installed to {0} Instances"
        NoInstalledPlat = "❌ No MT4/MT5 instances found, please install first"
        CacheOnlyTip    = "ℹ️ No accounts entered, EA files downloaded to cache"
        RiskDisclosure  = @'
⚠️ High Risk Investment Warning: Margin trading involves significant risk, and you may lose all or more than your initial investment.

Risk Disclosure and Disclaimer

01 Capital Risk Statement
Margin trading in Forex, precious metals, indices, and CFDs is a high-risk financial investment. Under unfavorable market conditions, your losses may exceed your initial investment. This type of investment is not suitable for all investors. Before engaging in any trading, you must carefully assess your financial situation, risk tolerance, and investment objectives.

02 EA Automated Trading Risk
Automated trading systems rely on the stability of trading platforms, network connections, server environments, and program code. Users are solely responsible for all trading risks arising from network latency, VPS server downtime, trading platform anomalies, spread widening, slippage, insufficient liquidity, improper parameter settings, or program errors. Past backtest results do not represent real trading performance.

03 Non-Investment Advice Statement
All content provided by this software is for learning, research, and technical communication purposes only and does not constitute any form of investment advice, trading guidance, or profit guarantee.

04 Non-Licensed and Non-Asset Management Statement
The provider of this software is not a licensed financial institution and does not engage in asset management, account custody, or fund management services. Users always have full control over their trading accounts and funds.

05 User Confirmation
Users should have fully read, understood, and agreed to all contents of this risk disclosure and disclaimer before using this software.
'@
    }
}

# ===================== 【数据配置】 =====================
# Win7兼容：不能用 New-Object PSObject -Property @{} 或 [PSCustomObject]@{}
# 必须用 New-Obj 函数（内部使用 Add-Member）

# 服务商配置
$p1 = New-Obj @{
    Name        = "Exness"
    MT5         = "https://download.terminal.free/cdn/web/exness.technologies.ltd/mt5/exness5setup.exe"
    MT4         = "https://download.terminal.free/cdn/web/exness.technologies.ltd/mt4/exness4setup.exe"
    RegisterUrl = "https://one.exnessonelink.com/a/9e5s9kc62z"
}
$Providers = @($p1)

# EA配置
$ea1 = New-Obj @{
    Name        = "SupaQuant"
    DisplayName = @{ CN = "全能量化交易专家"; EN = "Quantum SupaQuant" }
    Desc        = @{ CN = "精细化仓位 · 六维止盈护航 · 对冲引擎驱动 — 将马丁策略升华为机构级量化体系"; EN = "Fine-tuned position sizing · 6-dimensional TP shield · hedge engine — elevating Martingale to institutional-grade quant" }
    MT5         = "https://ea.sraph.com/EA/SupaQuant.ex5"
    MT4         = "https://ea.sraph.com/EA/SupaQuant.ex4"
}
$ea2 = New-Obj @{
    Name        = "SupaQuantLite"
    DisplayName = @{ CN = "全能量化交易专家(简化版)"; EN = "Quantum SupaQuant Lite" }
    Desc        = @{ CN = "全能量化核心简化版— 保留核心风控引擎，轻装上阵，新手友好"; EN = "SupaQuant Core Lite — retaining the core risk engine, lightweight and ready to go, beginner-friendly" }
    MT5         = "https://ea.sraph.com/EA/SupaQuantLite.ex5"
    MT4         = "https://ea.sraph.com/EA/SupaQuantLite.ex4"
}
$ea3 = New-Obj @{
    Name        = "Matrix"
    DisplayName = @{ CN = "量子·矩阵"; EN = "Quantum Matrix" }
    Desc        = @{ CN = "自适应网格 · 五维止盈防护 · 趋势感知 — 震荡拿稳利润，趋势张弛有度"; EN = "Adaptive grid · 5-dimensional TP shield · trend awareness — steady profits in ranging, balanced in trending" }
    MT5         = "https://ea.sraph.com/EA/Matrix.ex5"
    MT4         = "https://ea.sraph.com/EA/Matrix.ex4"
    MT5_CN      = "https://ea.sraph.com/EA/矩阵.ex5"
    MT4_CN      = "https://ea.sraph.com/EA/矩阵.ex4"
}
$ea4 = New-Obj @{
    Name        = "Phoenix"
    DisplayName = @{ CN = "量子·凤凰"; EN = "Quantum Phoenix" }
    Desc        = @{ CN = "挂单突破 · 顺势加仓 — 乘势而上，让趋势利润充分奔跑"; EN = "Pending order breakout · trend scaling — ride the momentum, let trend profits run free" }
    MT5         = "https://ea.sraph.com/EA/Phoenix.ex5"
    MT4         = "https://ea.sraph.com/EA/Phoenix.ex4"
}
$ea5 = New-Obj @{
    Name        = "Chaos"
    DisplayName = @{ CN = "量子·混沌"; EN = "Quantum Chaos" }
    Desc        = @{ CN = "混沌之中，自有秩序。不与市场争输赢，而是让市场的每一步波动都为账户服务"; EN = "Order within chaos. Not to fight the market, but to let every move of the market serve your account" }
    MT5         = "https://ea.sraph.com/EA/Chaos.ex5"
    MT4         = "https://ea.sraph.com/EA/Chaos.ex4"
}
$EAs = @($ea1, $ea2, $ea3, $ea4, $ea5)

# 系统路径
$BASE_DIR = "C:\XBOND"
$CACHE_DIR = Join-Path $BASE_DIR "Cache"
$TRADER_CACHE = Join-Path $CACHE_DIR "Trader"
$EA_CACHE = Join-Path $CACHE_DIR "EA"
$TEMP_DIR = Join-Path $BASE_DIR "Temp"
$DESKTOP_DIR = Join-Path ([Environment]::GetFolderPath("Desktop")) "XBOND"
$LOG_FILE = Join-Path $BASE_DIR "deploy_log_$(Get-Date -Format 'yyyyMMdd').log"

# ===================== 【工具函数 】 =====================
function Write-Log {
    param([string]$Msg, [string]$Color = "White")
    $logMsg = "[$(Get-Date -Format 'HH:mm:ss')] $Msg"
    if (Test-Path $BASE_DIR) { Add-Content -Path $LOG_FILE -Value $logMsg -Encoding UTF8 }
    Write-Host $Msg -ForegroundColor $Color
}

function Read-Selection {
    param([array]$Options)
    # PS2.0 兼容：显式检查数组元素个数
    if ($Options -is [array]) {
        $total = $Options.Count
        if ($total -eq 0) { $total = 1 }
    }
    else {
        $total = 1
    }

    do {
        $userInput = Read-Host
        $valid = $true
        $nums = @()
        try {
            $parts = $userInput.Split(',')
            foreach ($part in $parts) {
                $part = $part.Trim()
                if ($part -ne '') {
                    $n = [int]$part
                    if ($n -ge 1 -and $n -le $total) { $nums += $n }
                }
            }
        }
        catch { $valid = $false }
    } while (-not $valid -or $nums.Count -eq 0)

    $result = @()
    foreach ($n in $nums) { $result += $Options[$n - 1] }
    return $result
}

function Invoke-Download {
    param([string]$Url, [string]$Path)

    $maxRetries = 3
    $retryCount = 0
    $delaySeconds = 2
    $fileName = Split-Path $Path -Leaf
    $ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"

    $dlDir = Split-Path $Path
    if (-not (Test-Path $dlDir)) { New-Item -Path $dlDir -ItemType Directory -Force | Out-Null }

    # 单次下载尝试，成功返回 $true，失败返回错误文本
    function Invoke-OneMethod {
        param([string]$Method)
        if (Test-Path $Path) { Remove-Item $Path -Force -ErrorAction SilentlyContinue }

        if ($Method -eq 'IWR') {
            # Win10/11 首选：Invoke-WebRequest，底层走 WinINet，开箱即用
            try {
                Invoke-WebRequest -Uri $Url -OutFile $Path -UseBasicParsing `
                    -UserAgent $ua -TimeoutSec 60 -ErrorAction Stop
                return $true
            }
            catch { return $_.Exception.Message }
        }

        if ($Method -eq 'WebClient') {
            # Win7 / PS2-4：WebClient（.NET SChannel，依赖注册表已修复 TLS 1.2）
            $wc = $null
            try {
                $wc = New-Object System.Net.WebClient
                $wc.Headers.Add("User-Agent", $ua)
                $wc.CachePolicy = New-Object System.Net.Cache.RequestCachePolicy(
                    [System.Net.Cache.RequestCacheLevel]::NoCacheNoStore)
                $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
                if ($proxy) {
                    $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
                    $wc.Proxy = $proxy
                }
                $wc.DownloadFile($Url, $Path)
                return $true
            }
            catch {
                return $_.Exception.Message
            }
            finally {
                if ($wc) { $wc.Dispose() }
            }
        }

        if ($Method -eq 'BITS') {
            # 备用：BITS 后台传输服务（WinHTTP，Win7+ 内置）
            try {
                Import-Module BitsTransfer -ErrorAction Stop
                Start-BitsTransfer -Source $Url -Destination $Path `
                    -TransferType Download -ErrorAction Stop
                return $true
            }
            catch {
                if (Test-Path $Path) { Remove-Item $Path -Force -ErrorAction SilentlyContinue }
                return $_.Exception.Message
            }
        }

        if ($Method -eq 'certutil') {
            # 最后备用：certutil（WinHTTP）
            $out = & certutil -urlcache -split -f $Url $Path 2>&1
            & certutil -urlcache -f $Url delete 2>&1 | Out-Null
            if ($LASTEXITCODE -eq 0) { return $true }
            if (Test-Path $Path) { Remove-Item $Path -Force -ErrorAction SilentlyContinue }
            return ($out -join ' ')
        }
    }

    # Win10/11（PS5+）优先用 IWR；Win7 从 WebClient 开始
    if ($psVersion -ge 5 -and -not $isWin7) {
        $methods = @('IWR', 'BITS', 'certutil')
    }
    else {
        Write-Log "⚠️ Windows 7 detected, prioritizing WebClient for better TLS support" "Yellow"
        $methods = @('WebClient', 'BITS', 'certutil')
    }

    while ($retryCount -lt $maxRetries) {
        $retryCount++
        Write-Log "⬇️  $fileName (尝试 $retryCount/$maxRetries)..." "Cyan"
        $lastErr = ''

        foreach ($method in $methods) {
            $result = Invoke-OneMethod $method
            if ($result -eq $true) {
                $fileSize = (Get-Item $Path -ErrorAction SilentlyContinue).Length
                if ($fileSize -gt 0) {
                    Write-Log "✅ 下载成功 [$method]: $fileName ($fileSize 字节)" "Green"
                    return $true
                }
                if (Test-Path $Path) { Remove-Item $Path -Force -ErrorAction SilentlyContinue }
                $result = "文件为空"
            }
            $lastErr = "[$method] $result"
            Write-Log "  $lastErr" "Gray"
        }

        Write-Log "❌ 第 $retryCount 次全部失败: $lastErr" "Red"
        if ($retryCount -lt $maxRetries) {
            Write-Log "  等待 $delaySeconds 秒后重试..." "Gray"
            Start-Sleep -Seconds $delaySeconds
            $delaySeconds = [Math]::Min($delaySeconds * 2, 10)
        }
    }

    Write-Log "❌ 下载失败（将使用缓存或跳过）: $fileName" "Yellow"
    if (Test-Path $Path) { Remove-Item $Path -Force -ErrorAction SilentlyContinue }
    return $false
}

function Save-EAToCache {
    param([array]$SelectedEAs, [string]$PlatName)
    New-Item -Path $EA_CACHE -ItemType Directory -Force | Out-Null
    foreach ($ea in $SelectedEAs) {
        # 根据语言选取下载URL：中文用中文命名版（如有），其他用英文命名版
        $urlPropCN = "${PlatName}_CN"
        $url = if ($LangChoice -eq '1' -and $ea.$urlPropCN) { $ea.$urlPropCN } else { $ea.$PlatName }
        if ($url) {
            # URL 提取文件名：从最后一个 / 之后获取
            $fileName = $url.Substring($url.LastIndexOf('/') + 1)
            # 如果包含查询参数，只取文件名部分
            if ($fileName.Contains('?')) {
                $fileName = $fileName.Substring(0, $fileName.IndexOf('?'))
            }
            $cachePath = Join-Path $EA_CACHE $fileName
            Write-Log ($Text.DownloadEA -f $fileName) "Cyan"
            [void](Invoke-Download -Url $url -Path $cachePath)
        }
    }
}

function Copy-EAFromCache {
    param([string]$TargetDir, [string]$PlatName, [switch]$WithTracking)
    $mql = if ($PlatName -eq "MT5") { "MQL5" } else { "MQL4" }
    $eaTargetDir = Join-Path $TargetDir "$mql\Experts"
    New-Item -Path $eaTargetDir -ItemType Directory -Force | Out-Null
    $ext = if ($PlatName -eq "MT5") { "*.ex5" } else { "*.ex4" }
    # Win7兼容：不用 -File 参数
    Get-ChildItem -Path $EA_CACHE -Filter $ext | Where-Object { -not $_.PSIsContainer } | ForEach-Object {
        $destFile = Join-Path $eaTargetDir $_.Name
        if ($WithTracking) { Tx-TrackOverwrite -Path $destFile }
        Copy-Item -Path $_.FullName -Destination $destFile -Force
    }
}

function New-Shortcut {
    param([string]$Target, [string]$Path, [string]$Account)
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($Path)
        $shortcut.TargetPath = $Target
        $shortcut.Arguments = "/portable"
        $shortcut.WorkingDirectory = Split-Path $Target
        $shortcut.Save()
    }
    catch {}
}

function Clear-Temp {
    if (Test-Path $TEMP_DIR) { Remove-Item $TEMP_DIR -Recurse -Force -ErrorAction SilentlyContinue }
}

# ===================== 【事务模块】 =====================
# Tx-Begin             : 开始事务，清空回滚日志
# Tx-TrackNew          : 登记新建的文件/目录，回滚时删除
# Tx-TrackOverwrite    : 登记即将被覆盖的文件，先备份，回滚时恢复原文件
# Tx-TrackOverwriteDir : 登记即将被覆盖的目录，先移动备份，回滚时恢复
# Tx-Commit            : 提交事务，删除备份文件，清空日志
# Tx-Rollback          : 回滚事务，按逆序撤销所有已登记操作

$script:TxLog = @()

function Tx-Begin {
    $script:TxLog = @()
}

function Tx-TrackNew {
    param([string]$Type, [string]$Path)
    $script:TxLog = $script:TxLog + @(New-Obj @{ Op = 'New'; Type = $Type; Path = $Path; Backup = '' })
}

function Tx-TrackOverwrite {
    param([string]$Path)
    if (Test-Path $Path) {
        $backup = $Path + ".txbak"
        Copy-Item $Path $backup -Force -ErrorAction SilentlyContinue
        $script:TxLog = $script:TxLog + @(New-Obj @{ Op = 'Overwrite'; Type = 'File'; Path = $Path; Backup = $backup })
    }
    else {
        $script:TxLog = $script:TxLog + @(New-Obj @{ Op = 'New'; Type = 'File'; Path = $Path; Backup = '' })
    }
}

function Tx-TrackOverwriteDir {
    param([string]$Path)
    if (Test-Path $Path) {
        $backup = $Path + ".txbak"
        Move-Item -Path $Path -Destination $backup -ErrorAction Stop
        $script:TxLog = $script:TxLog + @(New-Obj @{ Op = 'OverwriteDir'; Type = 'Dir'; Path = $Path; Backup = $backup })
    }
    else {
        $script:TxLog = $script:TxLog + @(New-Obj @{ Op = 'New'; Type = 'Dir'; Path = $Path; Backup = '' })
    }
}

function Tx-Rollback {
    if ($script:TxLog.Count -eq 0) { return }
    Write-Log "↩️ 检测到错误，开始回滚..." "Yellow"
    for ($i = $script:TxLog.Count - 1; $i -ge 0; $i--) {
        $item = $script:TxLog[$i]
        try {
            switch ($item.Op) {
                'New' {
                    if ($item.Type -eq 'Dir' -and (Test-Path $item.Path)) {
                        Remove-Item $item.Path -Recurse -Force -ErrorAction SilentlyContinue
                        Write-Log "  ↩️ 已删除目录: $($item.Path)" "Gray"
                    }
                    elseif ($item.Type -eq 'File' -and (Test-Path $item.Path)) {
                        Remove-Item $item.Path -Force -ErrorAction SilentlyContinue
                        Write-Log "  ↩️ 已删除文件: $($item.Path)" "Gray"
                    }
                }
                'Overwrite' {
                    if ($item.Backup -ne '' -and (Test-Path $item.Backup)) {
                        Copy-Item $item.Backup $item.Path -Force -ErrorAction SilentlyContinue
                        Remove-Item $item.Backup -Force -ErrorAction SilentlyContinue
                        Write-Log "  ↩️ 已还原文件: $($item.Path)" "Gray"
                    }
                    elseif (Test-Path $item.Path) {
                        Remove-Item $item.Path -Force -ErrorAction SilentlyContinue
                        Write-Log "  ↩️ 已删除文件: $($item.Path)" "Gray"
                    }
                }
                'OverwriteDir' {
                    if ($item.Backup -ne '' -and (Test-Path $item.Backup)) {
                        if (Test-Path $item.Path) { Remove-Item $item.Path -Recurse -Force -ErrorAction SilentlyContinue }
                        Move-Item -Path $item.Backup -Destination $item.Path -ErrorAction SilentlyContinue
                        Write-Log "  ↩️ 已还原目录: $($item.Path)" "Gray"
                    }
                    elseif (Test-Path $item.Path) {
                        Remove-Item $item.Path -Recurse -Force -ErrorAction SilentlyContinue
                        Write-Log "  ↩️ 已删除目录: $($item.Path)" "Gray"
                    }
                }
            }
        }
        catch { Write-Log "  回滚步骤异常: $_" "Red" }
    }
    $script:TxLog = @()
    Write-Log "↩️ 回滚完成，环境已恢复" "Yellow"
}

function Tx-Commit {
    foreach ($item in $script:TxLog) {
        try {
            if ($item.Backup -ne '' -and (Test-Path $item.Backup)) {
                if ($item.Type -eq 'Dir') {
                    Remove-Item $item.Backup -Recurse -Force -ErrorAction SilentlyContinue
                }
                else {
                    Remove-Item $item.Backup -Force -ErrorAction SilentlyContinue
                }
            }
        }
        catch {}
    }
    $script:TxLog = @()
}

# ===================== 【主流程 】 =====================

# 语言选择：Win7兼容：-notin 改为 -ne -and -ne
do { $LangChoice = Read-Host "`n$($Lang.CN.LangSelect)" } while ($LangChoice -ne '1' -and $LangChoice -ne '2')
if ($LangChoice -eq '1') { $Text = $Lang.CN } else { $Text = $Lang.EN }

# 风险披露
Write-Host "`n============================================================" -ForegroundColor Red
Write-Host $Text.RiskDisclosure -ForegroundColor Red
Write-Host "============================================================" -ForegroundColor Red

$riskConfirm = Read-Host "`n$($Text.RiskConfirm)"
if ($riskConfirm -notmatch 'Y|y') {
    Write-Host "❌ 未确认风险，脚本退出" -ForegroundColor Red
    return
}

# 欢迎界面
Write-Host "`n==========================================" -ForegroundColor Yellow
Write-Log $Text.Welcome "Yellow"
Write-Host "==========================================" -ForegroundColor Yellow

# 创建必要目录：Win7兼容：逐个创建，避免多路径数组的兼容问题
New-Item -Path $BASE_DIR     -ItemType Directory -Force | Out-Null
New-Item -Path $CACHE_DIR    -ItemType Directory -Force | Out-Null
New-Item -Path $TRADER_CACHE -ItemType Directory -Force | Out-Null
New-Item -Path $EA_CACHE     -ItemType Directory -Force | Out-Null
New-Item -Path $TEMP_DIR     -ItemType Directory -Force | Out-Null
New-Item -Path $DESKTOP_DIR  -ItemType Directory -Force | Out-Null
Write-Log ($Text.LogPath -f $LOG_FILE) "Cyan"

# 主菜单
while ($true) {
    Write-Log "`n$($Text.MainMenu)" "Green"
    Write-Log $Text.Menu1 "White"
    Write-Log $Text.Menu2 "White"
    Write-Log $Text.Menu3 "White"
    Write-Log $Text.Menu4 "White"
    Write-Log $Text.Menu5 "White"

    $choice = Read-Host "`n$($Text.SelectTip)"
    switch ($choice) {

        '1' {
            try {
                Tx-Begin

                # ---- 步骤1：选择EA ----
                Write-Log "`n$($Text.EAList)" "Cyan"
                if ($EAs -is [array]) {
                    $eaCount = $EAs.Count
                }
                else {
                    $eaCount = 1
                    $EAs = @($EAs)
                }

                for ($n = 1; $n -le $eaCount; $n++) {
                    $ea = $EAs[$n - 1]
                    $line = "$n. $($ea.Name)"
                    if ($ea.psobject.properties.name -contains 'DisplayName' -and $ea.DisplayName) { $line += " [$(Get-I18n $ea.DisplayName)]" }
                    if ($ea.psobject.properties.name -contains 'Desc' -and $ea.Desc) { $line += "  $(Get-I18n $ea.Desc)" }
                    Write-Log $line "White"
                }
                Write-Log $Text.EASelect "White"
                $selectedEAs = Read-Selection $EAs

                # ---- 步骤2：选择平台 ----
                $commonPlats = @("MT5", "MT4")
                foreach ($ea in $selectedEAs) {
                    $eaPlats = @()
                    if ($ea.psobject.properties.name -contains 'MT5' -and $ea.MT5) { $eaPlats += "MT5" }
                    if ($ea.psobject.properties.name -contains 'MT4' -and $ea.MT4) { $eaPlats += "MT4" }
                    # Win7兼容：-in 改为 -contains
                    $commonPlats = $commonPlats | Where-Object { $eaPlats -contains $_ }
                }

                if ($commonPlats.Count -eq 0) {
                    Write-Log $Text.EANoPlat "Red"; continue
                }
                elseif ($commonPlats.Count -eq 1) {
                    $platName = $commonPlats[0]
                    Write-Log ($Text.PlatAuto -f $platName) "Green"
                }
                else {
                    Write-Log ($Text.PlatSelect -f "MT5", "MT4") "Cyan"
                    $platChoice = Read-Host
                    if ($platChoice -eq '1') { $platName = "MT5" }
                    elseif ($platChoice -eq '2') { $platName = "MT4" }
                    else { Write-Log $Text.InvalidInput "Red"; continue }
                }

                $isMT5 = ($platName -eq "MT5")
                $exe = if ($isMT5) { "terminal64.exe" } else { "terminal.exe" }
                $traderMasterDir = Join-Path $TRADER_CACHE $platName

                # ---- 步骤3：选择服务商 ----
                $availableProviders = @()
                foreach ($p in $Providers) {
                    try {
                        # 使用 Get-Member 获取属性列表
                        $propNames = @($p.psobject.properties | Select-Object -ExpandProperty Name)
                        $hasProperty = $propNames -contains $platName

                        if ($hasProperty) {
                            $val = $p | Select-Object -ExpandProperty $platName
                            if ($val) {
                                $availableProviders += $p
                            }
                        }
                    }
                    catch {
                        # 属性检测失败，跳过该服务商
                    }
                }

                if ($availableProviders.Count -eq 0) {
                    Write-Log ($Text.ProviderNoPlat -f $platName) "Red"; continue
                }
                elseif ($availableProviders.Count -eq 1) {
                    $selectedProvider = $availableProviders[0]
                    Write-Log ($Text.ProviderAuto -f $selectedProvider.Name, $platName) "Green"
                }
                else {
                    Write-Log "`n$($Text.ProviderList -f $platName)" "Cyan"
                    for ($n = 1; $n -le $availableProviders.Count; $n++) {
                        Write-Log "$n. $($availableProviders[$n-1].Name)" "White"
                    }
                    Write-Log $Text.ProviderSelect "White"
                    $selectedProvider = (Read-Selection $availableProviders)[0]
                }

                $downloadUrl = $selectedProvider.$platName
                $installerPath = Join-Path $TEMP_DIR "$platName.exe"

                Write-Log "`n$($Text.RegisterTip)" "Green"
                Write-Log "$($Text.RegisterUrl)$($selectedProvider.RegisterUrl)" "Cyan"
                Write-Log "------------------------------------------------" "Yellow"

                # ---- 步骤4：输入账号 ----
                $Accounts = @()
                Write-Log "`n=== 账号配置 ===" "Cyan"
                while ($true) {
                    $acc = Read-Host $Text.AccInput
                    # Win7兼容：不用 IsNullOrWhiteSpace（需.NET4），改用 Trim()
                    if ($acc.Trim() -eq '') { break }
                    $Accounts += $acc.Trim()
                }
                if ($Accounts.Count -eq 0) {
                    # No accounts entered: only download EA files to cache
                    Save-EAToCache -SelectedEAs $selectedEAs -PlatName $platName
                    Tx-Commit
                    Write-Log $Text.CacheOnlyTip "Green"
                    continue
                }

                # ---- 安装交易程序（缓存复用）----
                Write-Log ($Text.InstallStart -f $platName) "Cyan"
                if (-not (Test-Path $traderMasterDir)) {
                    Write-Log ($Text.DownloadTrader -f $platName) "Cyan"
                    if (-not (Invoke-Download -Url $downloadUrl -Path $installerPath)) {
                        throw "安装包下载失败，无法安装交易程序"
                    }
                    Start-Process $installerPath -ArgumentList "/auto /path:`"$traderMasterDir`"" -Wait
                    Start-Sleep -Seconds 3
                    # Win7兼容：Stop-Process 不支持 -Name 传数组，分两行
                    Stop-Process -Name "terminal"   -Force -ErrorAction SilentlyContinue
                    Stop-Process -Name "terminal64" -Force -ErrorAction SilentlyContinue
                    Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
                    if (-not (Test-Path (Join-Path $traderMasterDir $exe))) {
                        if (Test-Path $traderMasterDir) { Remove-Item $traderMasterDir -Recurse -Force -ErrorAction SilentlyContinue }
                        throw "交易程序安装失败，请重新运行脚本"
                    }

                    # 创建默认快捷方式
                    $defaultLnk = Join-Path $traderMasterDir "Default.lnk"
                    New-Shortcut -Target (Join-Path $traderMasterDir $exe) -Path $defaultLnk -Account "Default"
                }
                else {
                    Write-Log ($Text.CacheHit -f $traderMasterDir) "Green"
                }

                # ---- 下载EA并拷贝到母盘 ----
                Save-EAToCache -SelectedEAs $selectedEAs -PlatName $platName
                Write-Log $Text.CopyEA "Cyan"
                Copy-EAFromCache -TargetDir $traderMasterDir -PlatName $platName

                # ---- 部署各账号（带事务保护）----
                foreach ($acc in $Accounts) {
                    $target = Join-Path $BASE_DIR $acc
                    Write-Log ($Text.DeployAcc -f $acc) "Green"

                    if (Test-Path $target) {
                        $overwriteConfirm = Read-Host ($Text.OverWrite -f $target)
                        if ($overwriteConfirm -notmatch 'Y|y') {
                            Write-Log ($Text.SkipExist -f $target) "Yellow"
                            continue
                        }
                        # Move existing dir to backup instead of deleting (enables rollback restore)
                        Tx-TrackOverwriteDir -Path $target
                    }
                    else {
                        Tx-TrackNew -Type 'Dir' -Path $target
                    }

                    Copy-Item -Path $traderMasterDir -Destination $target -Recurse -Force -ErrorAction Stop
                    $desktopLnk = Join-Path $DESKTOP_DIR "$acc.lnk"
                    New-Shortcut -Target (Join-Path $target $exe) -Path (Join-Path $target "$acc.lnk") -Account $acc
                    New-Shortcut -Target (Join-Path $target $exe) -Path $desktopLnk -Account $acc
                    Tx-TrackNew -Type 'File' -Path $desktopLnk
                }

                Tx-Commit
                Write-Log $Text.Success "Green"
            }
            catch {
                Write-Log ($Text.Error -f $_.Exception.Message) "Red"
                Tx-Rollback
            }
            finally {
                Clear-Temp
            }
        }

        '2' {
            try {
                Tx-Begin

                Write-Log "`n$($Text.NewEATip)" "Cyan"

                # Win7兼容：Get-ChildItem 不支持 -Directory 参数
                # 扫描账号实例目录（BASE_DIR 直接子目录）
                $installedPlats = @()
                Get-ChildItem $BASE_DIR | Where-Object { $_.PSIsContainer } | ForEach-Object {
                    $dir = $_.FullName
                    if (Test-Path (Join-Path $dir "terminal64.exe")) { $installedPlats += "MT5" }
                    elseif (Test-Path (Join-Path $dir "terminal.exe")) { $installedPlats += "MT4" }
                }
                # 同时检查 Trader 缓存目录（全新安装但未部署账号时也能被识别）
                foreach ($plat in @("MT5", "MT4")) {
                    $cExe = if ($plat -eq "MT5") { "terminal64.exe" } else { "terminal.exe" }
                    $cDir = Join-Path $TRADER_CACHE $plat
                    if ((Test-Path (Join-Path $cDir $cExe)) -and ($installedPlats -notcontains $plat)) {
                        $installedPlats += $plat
                    }
                }

                if ($installedPlats.Count -eq 0) {
                    Write-Log $Text.NoInstalledPlat "Red"; continue
                }

                # 直接复用 $EAs 展示选项（与全新安装流程一致，避免二次构建引入的兼容性问题）
                Write-Log "`n$($Text.EAList)" "Cyan"
                if ($EAs -is [array]) {
                    $eaCount = $EAs.Count
                }
                else {
                    $eaCount = 1
                    $EAs = @($EAs)
                }

                for ($n = 1; $n -le $eaCount; $n++) {
                    $ea = $EAs[$n - 1]
                    $line = "$n. $($ea.Name)"
                    if ($ea.psobject.properties.name -contains 'DisplayName' -and $ea.DisplayName) { $line += " [$(Get-I18n $ea.DisplayName)]" }
                    if ($ea.psobject.properties.name -contains 'Desc' -and $ea.Desc) { $line += "  $(Get-I18n $ea.Desc)" }
                    Write-Log $line "White"
                }
                Write-Log $Text.EASelect "White"
                $selectedEAs = Read-Selection $EAs

                foreach ($plat in ($installedPlats | Select-Object -Unique)) {
                    $selectedEAsForPlat = @()
                    foreach ($ea in $selectedEAs) {
                        if ($ea.$plat) { $selectedEAsForPlat += $ea }
                    }
                    if ($selectedEAsForPlat.Count -eq 0) { continue }

                    Save-EAToCache -SelectedEAs $selectedEAsForPlat -PlatName $plat

                    # 更新 Trader 母盘（新部署账号可继承最新EA）
                    $traderPlatDir = Join-Path $TRADER_CACHE $plat
                    if (Test-Path $traderPlatDir) {
                        Copy-EAFromCache -TargetDir $traderPlatDir -PlatName $plat
                    }

                    # 更新所有已部署的账号实例
                    Get-ChildItem $BASE_DIR | Where-Object { $_.PSIsContainer } | ForEach-Object {
                        if ($_.Name -match '^(Cache|Temp|.*_Temp)$') { return }
                        $targetDir = $_.FullName
                        if ((Test-Path (Join-Path $targetDir "terminal64.exe")) -and $plat -eq "MT5") {
                            Copy-EAFromCache -TargetDir $targetDir -PlatName "MT5" -WithTracking
                        }
                        elseif ((Test-Path (Join-Path $targetDir "terminal.exe")) -and $plat -eq "MT4") {
                            Copy-EAFromCache -TargetDir $targetDir -PlatName "MT4" -WithTracking
                        }
                        Write-Log ($Text.EACopySuccess -f $_.Name) "Green"
                    }
                }

                Tx-Commit
                Write-Log $Text.Success "Green"
            }
            catch {
                Write-Log ($Text.Error -f $_.Exception.Message) "Red"
                Tx-Rollback
            }
        }

        '3' {
            # Win7兼容：-notin 改为 -ne -and -ne
            do { $unChoice = Read-Host "`n$($Text.UninstallMenu)" } while ($unChoice -ne '1' -and $unChoice -ne '2' -and $unChoice -ne '3')

            if ($unChoice -eq '1') {
                $acc = Read-Host $Text.Uninput
                $path = Join-Path $BASE_DIR $acc
                $lnk = Join-Path $DESKTOP_DIR "$acc.lnk"
                if (Test-Path $path) { Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue }
                if (Test-Path $lnk) { Remove-Item $lnk  -Force -ErrorAction SilentlyContinue }
                Write-Log "✅ 卸载完成: $acc" "Green"
            }
            if ($unChoice -eq '2') {
                $unAllConfirm = Read-Host $Text.UnAllConfirm
                if ($unAllConfirm -match 'Y|y') {
                    Get-ChildItem $BASE_DIR | Where-Object { $_.PSIsContainer -and $_.Name -notmatch '^(Cache|Temp|.*_Temp)$' } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                    Remove-Item "$DESKTOP_DIR\*.lnk" -Force -ErrorAction SilentlyContinue
                    Write-Log "✅ 全部卸载完成" "Green"
                }
            }
        }

        '4' {
            # List EA files in cache
            $cachedEAs = @(Get-ChildItem -Path $EA_CACHE -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer -and ($_.Extension -eq '.ex5' -or $_.Extension -eq '.ex4') })
            if ($cachedEAs.Count -eq 0) { Write-Log $Text.NoEACache "Yellow"; break }

            Write-Log "`n$($Text.DeleteEAList)" "Cyan"
            for ($n = 1; $n -le $cachedEAs.Count; $n++) {
                Write-Log "$n. $($cachedEAs[$n - 1].Name)" "White"
            }
            Write-Log $Text.DeleteEASelect "White"
            $toDelete = Read-Selection $cachedEAs

            foreach ($f in $toDelete) {
                Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue
                Write-Log ($Text.DeleteEADone -f $f.Name) "Green"

                # Delete from trader master dirs and all account instances
                $dirsToClean = @()
                foreach ($plat in @('MT5', 'MT4')) {
                    $d = Join-Path $TRADER_CACHE $plat
                    if (Test-Path $d) { $dirsToClean = $dirsToClean + @($d) }
                }
                Get-ChildItem $BASE_DIR -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer } | ForEach-Object {
                    $dirsToClean = $dirsToClean + @($_.FullName)
                }
                foreach ($dir in $dirsToClean) {
                    foreach ($mql in @('MQL5', 'MQL4')) {
                        $eaPath = Join-Path $dir "$mql\Experts\$($f.Name)"
                        if (Test-Path $eaPath) {
                            Remove-Item $eaPath -Force -ErrorAction SilentlyContinue
                            Write-Log "  $($f.Name) <- $(Split-Path $dir -Leaf)" "Gray"
                        }
                    }
                }
            }
            Write-Log $Text.Success "Green"
        }

        '5' { exit }
        default { Write-Log $Text.InvalidInput "Red" }
    }
}
