<#
.SYNOPSIS
  Office PDF 工具 —— 一键安装（当前用户，无需管理员）。

.DESCRIPTION
  安装内容：
    1. 程序文件复制到 %LOCALAPPDATA%\OfficePdfTools\app（含自带 Node 运行时）
    2. 生成并信任 localhost 证书（Office 任务窗格必须通过 HTTPS 加载）
    3. 把加载项清单注册到 Office 开发者侧载项 —— 之后每次打开 Office 都会自动出现「PDF 工具」选项卡
    4. 注册登录自启（托盘程序负责拉起本地服务）并立即启动服务
    5. 创建开始菜单快捷方式与「应用和功能」卸载项

.PARAMETER Port
  本地服务端口，默认 3000；若被其它程序占用会自动顺延并在清单中同步。

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File install.ps1
  powershell -ExecutionPolicy Bypass -File install.ps1 -Port 3001 -Silent
#>
[CmdletBinding()]
param(
  [string]$SourceDir = '',
  [string]$InstallDir = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'OfficePdfTools'),
  [int]$Port = 3000,
  [switch]$NoAutoStart,
  [switch]$SkipRegistry,
  [switch]$SkipTls,
  [switch]$SkipShortcuts,
  [switch]$NoStart,
  [switch]$Silent,
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

# 目录与系统程序一律通过 .NET 解析：受限会话里环境变量可能缺失
$SystemDir = [Environment]::GetFolderPath('System')
$UserAppData = [Environment]::GetFolderPath('ApplicationData')
$WscriptExe = Join-Path $SystemDir 'wscript.exe'
$TaskkillExe = Join-Path $SystemDir 'taskkill.exe'
$PowerShellExe = Join-Path $SystemDir 'WindowsPowerShell\v1.0\powershell.exe'
if (-not (Test-Path -LiteralPath $PowerShellExe)) { $PowerShellExe = 'powershell.exe' }

$AppName = 'Office PDF 工具'
$AppId = 'OfficePdfTools'
$Publisher = 'ZizzSaint'
$Version = '1.1.0'

function Say([string]$message, [string]$color = 'Gray') {
  if ($Silent -and $color -eq 'Gray') { return }
  Write-Host $message -ForegroundColor $color
}

function Invoke-Do([string]$desc, [scriptblock]$action) {
  if ($DryRun) {
    Write-Host "  [DryRun] $desc" -ForegroundColor DarkGray
    return
  }
  Write-Host "  · $desc"
  & $action
}

function Stop-ProcessSafe([int]$processId) {
  # taskkill 对已退出的 PID 会往 stderr 写 ERROR，配合 $ErrorActionPreference='Stop' 会变成终止性错误，
  # 这里临时放宽偏好并吞掉输出，保证重复安装/卸载不会因此中断。
  $previous = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  try { & $TaskkillExe /PID $processId /T /F *> $null } catch { }
  finally { $ErrorActionPreference = $previous }
}
Say "=== $AppName 安装程序 v$Version ===" Cyan

# PowerShell 5.1 的怪异行为：[CmdletBinding()] + -File 调用时 $PSScriptRoot 在参数默认值里为空，
# 所以在函数体里再兜底解析一次。
if (-not $SourceDir) { $SourceDir = $PSScriptRoot }
if (-not $SourceDir) { $SourceDir = (Get-Location).Path }

$payload = $SourceDir
if (Test-Path -LiteralPath (Join-Path $SourceDir 'payload')) {
  $payload = Join-Path $SourceDir 'payload'
}
if (-not (Test-Path -LiteralPath (Join-Path $payload 'server\index.js'))) {
  throw "在 $payload 中找不到程序文件（server/index.js），请从解压后的安装包目录运行本脚本。"
}
Say "程序文件: $payload"

$InstallDir = [System.IO.Path]::GetFullPath($InstallDir)
$appDir = Join-Path $InstallDir 'app'
$toolsDir = Join-Path $InstallDir 'tools'
$logDir = Join-Path $InstallDir 'logs'
$outputDir = Join-Path $InstallDir 'output'
$tlsDir = Join-Path $InstallDir 'tls'
$manifestPath = Join-Path $InstallDir 'manifest.xml'
$nodeExe = Join-Path $appDir 'node\node.exe'
Say "安装目录: $InstallDir"

$hasOffice = $false
foreach ($progId in @('Word.Application', 'Excel.Application', 'PowerPoint.Application')) {
  if ([Type]::GetTypeFromProgID($progId)) { $hasOffice = $true; break }
}
if (-not $hasOffice) {
  Say "提示：未检测到桌面版 Word/Excel/PowerPoint，安装完成后加载项不会显示；装好 Office 后首次启动会自动出现。" Yellow
}

function Test-OurService([int]$p) {
  try {
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $r = Invoke-WebRequest -Uri "https://localhost:$p/api/health" -TimeoutSec 3 -UseBasicParsing
    return ($r.Content -match 'office-pdf-tools')
  }
  catch { return $false }
}

function Test-PortBusy([int]$p) {
  $client = New-Object System.Net.Sockets.TcpClient
  try {
    $client.Connect('127.0.0.1', $p)
    $client.Close()
    return $true
  }
  catch { return $false }
  finally { $client.Dispose() }
}

if (Test-OurService $Port) {
  Say "端口 $Port 上已有本程序的服务在运行，将复用该端口。" Yellow
}
else {
  $start = $Port
  while ((Test-PortBusy $Port) -and ($Port -lt ($start + 20))) {
    Say "端口 $Port 被占用，尝试下一个端口。" Yellow
    $Port++
  }
}
Say "服务端口: $Port"

$pidFile = Join-Path $InstallDir 'service.pid'
if (Test-Path -LiteralPath $pidFile) {
  $oldPid = (Get-Content -LiteralPath $pidFile -Raw).Trim()
  if ($oldPid -match '^\d+$') {
    Invoke-Do "停止旧的服务进程 (PID $oldPid)" { Stop-ProcessSafe $oldPid }
  }
}
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
  Where-Object { $_.CommandLine -and $_.CommandLine -like "*$InstallDir*tray.ps1*" } |
  ForEach-Object { $oldPid2 = $_.ProcessId; Invoke-Do "结束旧的托盘进程 (PID $oldPid2)" { Stop-ProcessSafe $oldPid2 } }

Invoke-Do "创建安装目录" { New-Item -ItemType Directory -Force -Path $appDir, $toolsDir, $logDir, $outputDir, $tlsDir | Out-Null }

if ($DryRun) {
  Invoke-Do "复制 $payload → $appDir" { }
}
else {
  Say "  · 复制程序文件（约 230 MB，请稍候）"
  # 排除项必须写全路径：robocopy 的 /XD 按目录名匹配，写 "tmp" 会把 node_modules\tmp 一起排除掉
  $rcArgs = @(
    $payload, $appDir, "/E",
    "/XD", (Join-Path $payload "output"), (Join-Path $payload "logs"), (Join-Path $payload "tmp"),
            (Join-Path $payload ".git"), (Join-Path $payload ".github"),
    "/XF", (Join-Path $payload "service.pid"),
    "/NFL", "/NDL", "/NJH", "/NJS", "/NP", "/R:2", "/W:1"
  )
  & robocopy.exe @rcArgs | Out-Null
  if ($LASTEXITCODE -ge 8) { throw "复制程序文件失败 (robocopy exit $LASTEXITCODE)" }
  "install at $(Get-Date -Format s) payload=$payload port=$Port" | Add-Content -LiteralPath (Join-Path $logDir 'install.log') -Encoding UTF8
}

Invoke-Do "安装托盘与卸载脚本" {
  foreach ($name in @('tray.ps1', 'install.ps1', 'uninstall.ps1')) {
    $src = Join-Path $payload "scripts\$name"
    if (Test-Path -LiteralPath $src) { Copy-Item -LiteralPath $src -Destination $toolsDir -Force }
  }
  $vbs = Join-Path $payload 'scripts\start-hidden.vbs'
  if (Test-Path -LiteralPath $vbs) { Copy-Item -LiteralPath $vbs -Destination $InstallDir -Force }
  $ico = Join-Path $payload 'scripts\app.ico'
  if (Test-Path -LiteralPath $ico) { Copy-Item -LiteralPath $ico -Destination $toolsDir -Force }
}

Invoke-Do "写入端口文件 port.txt = $Port" {
  Set-Content -LiteralPath (Join-Path $InstallDir 'port.txt') -Value $Port -Encoding ASCII
}

$resolvedNode = $nodeExe
if (-not (Test-Path -LiteralPath $resolvedNode)) { $resolvedNode = 'node' }

if ($SkipTls) {
  Say "已跳过证书步骤 (-SkipTls)"
}
elseif (Test-Path -LiteralPath (Join-Path $tlsDir 'localhost.crt')) {
  Say "证书已存在，跳过生成"
}
else {
  Invoke-Do "生成并信任 localhost 证书（可能出现一次系统确认框，请选择「是」）" {
    Push-Location $appDir
    try {
      $result = & $resolvedNode (Join-Path $appDir 'scripts\setup-tls.mjs') $tlsDir 2>&1
      Say "    $result"
    }
    finally { Pop-Location }
  }
}

Invoke-Do "生成加载项清单（端口 $Port）" {
  $xml = Get-Content -LiteralPath (Join-Path $payload 'manifest\manifest.xml') -Raw
  $xml = [regex]::Replace($xml, 'https://localhost:\d+', "https://localhost:$Port")
  $utf8Bom = New-Object System.Text.UTF8Encoding($true)
  [System.IO.File]::WriteAllText($manifestPath, $xml, $utf8Bom)
}

if ($SkipRegistry) {
  Say "已跳过注册表写入 (-SkipRegistry)"
}
else {
  foreach ($ver in @('16.0', '15.0')) {
    Invoke-Do "注册加载项到 Office $ver 开发者侧载项（之后每次打开 Office 自动加载）" {
      $key = "HKCU:\Software\Microsoft\Office\$ver\Wef\Developer"
      New-Item -Path $key -Force | Out-Null
      New-ItemProperty -Path $key -Name $manifestPath -Value $manifestPath -PropertyType String -Force | Out-Null
    }
  }

  if ($NoAutoStart) {
    Say "已跳过登录自启 (-NoAutoStart)"
  }
  else {
    Invoke-Do "注册登录自启（托盘程序）" {
      $runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
      $cmd = "`"$WscriptExe`" `"$(Join-Path $InstallDir 'start-hidden.vbs')`""
      New-ItemProperty -Path $runKey -Name $AppId -Value $cmd -PropertyType String -Force | Out-Null
    }
  }

  Invoke-Do "写入「应用和功能」卸载项" {
    $uninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$AppId"
    New-Item -Path $uninstallKey -Force | Out-Null
    $sizeKb = 0
    if (Test-Path -LiteralPath $appDir) {
      $sum = (Get-ChildItem -LiteralPath $appDir -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
      if ($sum) { $sizeKb = [int]($sum / 1KB) }
    }
    $uninstallCmd = "`"$PowerShellExe`" -NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $toolsDir 'uninstall.ps1')`""
    $props = [ordered]@{
      DisplayName          = $AppName
      DisplayVersion       = $Version
      Publisher            = $Publisher
      InstallLocation      = $InstallDir
      DisplayIcon          = (Join-Path $toolsDir 'app.ico')
      UninstallString      = $uninstallCmd
      QuietUninstallString = "$uninstallCmd -Silent"
      NoModify             = 1
      NoRepair             = 1
      EstimatedSize        = $sizeKb
      URLInfoAbout         = 'https://github.com/ZizzSaint/office-pdf-tools'
    }
    foreach ($k in $props.Keys) {
      New-ItemProperty -Path $uninstallKey -Name $k -Value $props[$k] -Force | Out-Null
    }
  }
}

if ($SkipShortcuts) {
  Say "已跳过快捷方式 (-SkipShortcuts)"
}
else {
  Invoke-Do "创建开始菜单快捷方式" {
    $menuDir = Join-Path $UserAppData 'Microsoft\Windows\Start Menu\Programs\Office PDF 工具'
    New-Item -ItemType Directory -Force -Path $menuDir | Out-Null
    $shell = New-Object -ComObject WScript.Shell
    $icon = Join-Path $toolsDir 'app.ico'
    $pane = "https://localhost:$Port/taskpane.html"

    $lnk = $shell.CreateShortcut((Join-Path $menuDir '打开任务窗格.lnk'))
    $lnk.TargetPath = $pane
    $lnk.IconLocation = $icon
    $lnk.Save()

    $lnk = $shell.CreateShortcut((Join-Path $menuDir '打开输出目录.lnk'))
    $lnk.TargetPath = 'explorer.exe'
    $lnk.Arguments = "`"$outputDir`""
    $lnk.IconLocation = 'shell32.dll,4'
    $lnk.Save()

    $lnk = $shell.CreateShortcut((Join-Path $menuDir '查看日志.lnk'))
    $lnk.TargetPath = 'notepad.exe'
    $lnk.Arguments = "`"$(Join-Path $logDir 'service.log')`""
    $lnk.IconLocation = 'shell32.dll,70'
    $lnk.Save()

    $lnk = $shell.CreateShortcut((Join-Path $menuDir '卸载 Office PDF 工具.lnk'))
    $lnk.TargetPath = $PowerShellExe
    $lnk.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $toolsDir 'uninstall.ps1')`""
    $lnk.IconLocation = $icon
    $lnk.Save()
  }
}

if ($NoStart -or $DryRun) {
  Say "已跳过立即启动服务"
}
else {
  Say "正在启动本地服务 …"
  Start-Process -FilePath $WscriptExe -ArgumentList "`"$(Join-Path $InstallDir 'start-hidden.vbs')`"" -WindowStyle Hidden
  $deadline = (Get-Date).AddSeconds(45)
  $healthy = $false
  while ((Get-Date) -lt $deadline) {
    if (Test-OurService $Port) { $healthy = $true; break }
    Start-Sleep -Milliseconds 900
  }
  if ($healthy) { Say "服务已就绪: https://localhost:$Port/taskpane.html" Green }
  else { Say "服务尚未就绪，请稍后右键托盘图标查看日志，或查看 logs\service.err.log。" Yellow }
}

Say ""
Say "=== 安装完成 ===" Green
Say "  安装目录 : $InstallDir"
Say "  输出目录 : $outputDir"
Say "  服务地址 : https://localhost:$Port/taskpane.html"
Say "  自动加载 : 已注册到 Office 开发者侧载项（16.0 / 15.0）"
if (-not $NoAutoStart) { Say "  开机自启 : 已启用（登录后托盘自动启动服务）" }
Say ""
Say "下一步：完全退出并重新打开 Word / Excel / PowerPoint，" Yellow
Say "       功能区就会出现「PDF 工具」选项卡。" Yellow
Say ""
Say "卸载：开始菜单 → Office PDF 工具 → 卸载 Office PDF 工具" Gray
