<#
.SYNOPSIS
  Office PDF 工具 —— 托盘程序（同时负责拉起并守护本地 HTTPS 服务）。

.DESCRIPTION
  安装版会在用户登录时自动运行本脚本：
    * 先启动 node server/index.js（HTTPS + 固定端口 + 安装目录证书）
    * 再尝试显示通知区域图标（右键：打开任务窗格 / 输出目录 / 日志 / 重启服务 / 退出）
    * 若当前会话无法显示托盘图标（例如无人值守环境），自动退化为无界面守护模式，服务照常运行
#>
[CmdletBinding()]
param(
  [string]$InstallDir = '',
  [int]$Port = 0
)

$ErrorActionPreference = 'Continue'

if (-not $InstallDir) { $InstallDir = Split-Path -Parent $PSScriptRoot }
if (-not $InstallDir) { $InstallDir = (Get-Location).Path }

$installDir = [System.IO.Path]::GetFullPath($InstallDir)
$logDir = Join-Path $installDir 'logs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$trayLog = Join-Path $logDir 'tray.log'

function Write-TrayLog([string]$message) {
  $line = "[" + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + "] " + $message
  Add-Content -LiteralPath $trayLog -Value $line -Encoding UTF8
}

if (-not $Port -or $Port -le 0) {
  $portFile = Join-Path $installDir 'port.txt'
  if (Test-Path -LiteralPath $portFile) {
    $raw = (Get-Content -LiteralPath $portFile -Raw).Trim()
    if ($raw -match '^\d+$') { $Port = [int]$raw }
  }
  if (-not $Port -or $Port -le 0) { $Port = 3000 }
}

$appDir = Join-Path $installDir 'app'
$nodeExe = Join-Path $appDir 'node\node.exe'
if (-not (Test-Path -LiteralPath $nodeExe)) { $nodeExe = 'node' }
$serverJs = Join-Path $appDir 'server\index.js'
$outputDir = Join-Path $installDir 'output'
$tlsKey = Join-Path $installDir 'tls\localhost.key'
$tlsCert = Join-Path $installDir 'tls\localhost.crt'
$pidFile = Join-Path $installDir 'service.pid'
$outLog = Join-Path $logDir 'service.log'
$errLog = Join-Path $logDir 'service.err.log'
$paneUrl = "https://localhost:$Port/taskpane.html"

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
Write-TrayLog "托盘启动 (install=$installDir, port=$Port, node=$nodeExe)"

$script:serverProcess = $null

function Stop-Server {
  if ($script:serverProcess -and -not $script:serverProcess.HasExited) {
    $target = $script:serverProcess.Id
    try {
      $taskkill = Join-Path ([Environment]::GetFolderPath('System')) 'taskkill.exe'
      & $taskkill /PID $target /T /F 2>$null | Out-Null
      Write-TrayLog "已停止服务 PID=$target"
    }
    catch { Write-TrayLog "停止服务失败: $($_.Exception.Message)" }
  }
  $script:serverProcess = $null
  Remove-Item -LiteralPath $pidFile -Force -ErrorAction SilentlyContinue
}

function Start-Server {
  Stop-Server
  if (-not (Test-Path -LiteralPath $serverJs)) {
    Write-TrayLog "未找到 $serverJs"
    return
  }
  foreach ($file in @($outLog, $errLog)) {
    if ((Test-Path -LiteralPath $file) -and (Get-Item -LiteralPath $file).Length -gt 5MB) {
      Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
    }
  }
  $argList = @($serverJs, '--https', '--port', "$Port", '--output', $outputDir)
  if ((Test-Path -LiteralPath $tlsKey) -and (Test-Path -LiteralPath $tlsCert)) {
    $argList += @('--tls-key', $tlsKey, '--tls-cert', $tlsCert)
  }
  try {
    $script:serverProcess = Start-Process -FilePath $nodeExe -ArgumentList $argList -PassThru -WindowStyle Hidden `
      -RedirectStandardOutput $outLog -RedirectStandardError $errLog
    Set-Content -LiteralPath $pidFile -Value $script:serverProcess.Id -Encoding ASCII
    Write-TrayLog "已启动服务 PID=$($script:serverProcess.Id) -> $paneUrl"
  }
  catch { Write-TrayLog "启动服务失败: $($_.Exception.Message)" }
}

function Test-Health {
  # 优先用系统自带的 curl（-k 跳过本地证书校验，避免 PS 5.1 的 TLS/证书坑）
  $curl = Join-Path ([Environment]::GetFolderPath('System')) 'curl.exe'
  if (Test-Path -LiteralPath $curl) {
    try {
      $out = & $curl -s -k --max-time 5 "https://localhost:$Port/api/health" 2>$null
      return ($out -match 'office-pdf-tools')
    }
    catch { return $false }
  }
  try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $res = Invoke-WebRequest -Uri "https://localhost:$Port/api/health" -TimeoutSec 5 -UseBasicParsing
    return ($res.StatusCode -eq 200)
  }
  catch { return $false }
}

function Wait-Healthy([int]$seconds = 60) {
  $deadline = (Get-Date).AddSeconds($seconds)
  while ((Get-Date) -lt $deadline) {
    if (Test-Health) { return $true }
    Start-Sleep -Milliseconds 900
  }
  return $false
}

function Open-Pane { Start-Process $paneUrl | Out-Null }

# ---------------------------------------------------------------- 先启动服务
Start-Server
$healthy = Wait-Healthy
Write-TrayLog ("健康检查: " + $(if ($healthy) { '通过' } else { '未通过' }))

# ---------------------------------------------------------------- 再尝试托盘
$trayReady = $false
try {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing

  $iconPath = Join-Path $installDir 'tools\app.ico'
  $icon = $null
  if (Test-Path -LiteralPath $iconPath) {
    try { $icon = New-Object System.Drawing.Icon($iconPath) } catch { $icon = $null }
  }
  if (-not $icon) { $icon = [System.Drawing.SystemIcons]::Application }

  $notify = New-Object System.Windows.Forms.NotifyIcon
  $notify.Icon = $icon
  $notify.Text = "Office PDF 工具（端口 $Port）"
  $notify.Visible = $true

  $menu = New-Object System.Windows.Forms.ContextMenuStrip
  $menu.Items.Add("打开任务窗格", $null, { Open-Pane }) | Out-Null
  $menu.Items.Add("打开输出目录", $null, { Start-Process explorer.exe $outputDir | Out-Null }) | Out-Null
  $menu.Items.Add("复制服务地址", $null, { [System.Windows.Forms.Clipboard]::SetText($paneUrl) }) | Out-Null
  $menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
  $menu.Items.Add("重启服务", $null, { Start-Server; if (Wait-Healthy) { $notify.ShowBalloonTip(2000, 'Office PDF 工具', '服务已重启', [System.Windows.Forms.ToolTipIcon]::Info) } }) | Out-Null
  $menu.Items.Add("查看日志", $null, { Start-Process notepad.exe $trayLog | Out-Null }) | Out-Null
  $menu.Items.Add("打开安装目录", $null, { Start-Process explorer.exe $installDir | Out-Null }) | Out-Null
  $menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
  $menu.Items.Add("退出", $null, {
    Stop-Server
    $notify.Visible = $false
    [System.Windows.Forms.Application]::ExitThread()
  }) | Out-Null

  $notify.ContextMenuStrip = $menu
  $notify.add_DoubleClick({ Open-Pane })
  if (-not $healthy) {
    $notify.ShowBalloonTip(4000, 'Office PDF 工具', '服务启动较慢或失败，右键托盘图标可查看日志。', [System.Windows.Forms.ToolTipIcon]::Warning)
  }
  $trayReady = $true
}
catch {
  Write-TrayLog "托盘初始化失败，进入无界面守护模式: $($_.Exception.Message)"
}

if ($trayReady) {
  try { [System.Windows.Forms.Application]::Run() }
  finally {
    Stop-Server
    try { $notify.Visible = $false; $notify.Dispose() } catch { }
    Write-TrayLog '托盘退出'
  }
}
else {
  # 无界面守护：服务挂掉就拉起
  while ($true) {
    Start-Sleep -Seconds 30
    if (-not $script:serverProcess -or $script:serverProcess.HasExited) {
      Write-TrayLog '检测到服务已退出，正在重启'
      Start-Server
      Wait-Healthy 20 | Out-Null
    }
  }
}
