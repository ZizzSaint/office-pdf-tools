<#
.SYNOPSIS
  Office PDF 工具 —— 卸载（当前用户）。

.DESCRIPTION
  1. 停止托盘与本地服务
  2. 从 Office 开发者侧载项中移除加载项（Word/Excel/PowerPoint 不再显示该选项卡）
  3. 移除登录自启、开始菜单快捷方式、「应用和功能」卸载项
  4. 删除程序文件（输出目录可用 -KeepOutput 保留）

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File uninstall.ps1
  powershell -ExecutionPolicy Bypass -File uninstall.ps1 -KeepOutput -Silent
#>
[CmdletBinding()]
param(
  [string]$InstallDir = '',
  [switch]$KeepOutput,
  [switch]$KeepShortcuts,
  [switch]$Silent,
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

# 同样用 .NET 取目录，避免依赖环境变量
$SystemDir = [Environment]::GetFolderPath('System')
$TaskkillExe = Join-Path $SystemDir 'taskkill.exe'
$UserAppData = [Environment]::GetFolderPath('ApplicationData')
$UserProfile = [Environment]::GetFolderPath('UserProfile')

$AppName = 'Office PDF 工具'
$AppId = 'OfficePdfTools'

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

if (-not $InstallDir) { $InstallDir = Split-Path -Parent $PSScriptRoot }
if (-not $InstallDir) { $InstallDir = (Get-Location).Path }

function Stop-ProcessSafe([int]$processId) {
  # taskkill 对已退出的 PID 会往 stderr 写 ERROR，配合 $ErrorActionPreference='Stop' 会变成终止性错误，
  # 这里临时放宽偏好并吞掉输出，保证重复安装/卸载不会因此中断。
  $previous = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  try { & $TaskkillExe /PID $processId /T /F *> $null } catch { }
  finally { $ErrorActionPreference = $previous }
}
Say "=== 卸载 $AppName ===" Cyan
$InstallDir = [System.IO.Path]::GetFullPath($InstallDir)
Say "安装目录: $InstallDir"

# 1) 停止进程
$pidFile = Join-Path $InstallDir 'service.pid'
if (Test-Path -LiteralPath $pidFile) {
  $svcPid = (Get-Content -LiteralPath $pidFile -Raw).Trim()
  if ($svcPid -match '^\d+$') {
    Invoke-Do "停止服务进程 (PID $svcPid)" { Stop-ProcessSafe $svcPid }
  }
}
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
  Where-Object { $_.CommandLine -and $_.CommandLine -like "*$InstallDir*tray.ps1*" } |
  ForEach-Object { $trayPid = $_.ProcessId; Invoke-Do "结束托盘进程 (PID $trayPid)" { Stop-ProcessSafe $trayPid } }

# 2) 移除 Office 加载项注册
foreach ($ver in @('16.0', '15.0')) {
  $key = "HKCU:\Software\Microsoft\Office\$ver\Wef\Developer"
  if (Test-Path $key) {
    $props = (Get-Item $key).Property
    foreach ($name in $props) {
      if ($name -like "*$InstallDir*" -or $name -like '*OfficePdfTools*') {
        Invoke-Do "移除加载项注册值: $name" { Remove-ItemProperty -Path $key -Name $name -Force -ErrorAction SilentlyContinue }
      }
    }
    if (-not (Get-Item $key -ErrorAction SilentlyContinue).Property.Count) {
      Invoke-Do "删除空的侧载项注册表键 ($ver)" { Remove-Item -Path $key -Force -ErrorAction SilentlyContinue }
    }
  }
}

# 3) 自启、卸载项、快捷方式
Invoke-Do "移除登录自启" {
  Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name $AppId -Force -ErrorAction SilentlyContinue
}
Invoke-Do "移除「应用和功能」卸载项" {
  Remove-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$AppId" -Recurse -Force -ErrorAction SilentlyContinue
}
if ($KeepShortcuts) {
  Say "已保留开始菜单快捷方式 (-KeepShortcuts)"
}
else {
  Invoke-Do "移除开始菜单快捷方式" {
    Remove-Item -Path (Join-Path $UserAppData 'Microsoft\Windows\Start Menu\Programs\Office PDF 工具') -Recurse -Force -ErrorAction SilentlyContinue
  }
}

# 4) 删除文件
$outputDir = Join-Path $InstallDir 'output'
if ($KeepOutput -and (Test-Path -LiteralPath $outputDir)) {
  $keepDir = Join-Path $UserProfile 'Documents\OfficePdfTools-output'
  Invoke-Do "保留输出文件到 $keepDir" {
    New-Item -ItemType Directory -Force -Path $keepDir | Out-Null
    Copy-Item -Path (Join-Path $outputDir '*') -Destination $keepDir -Recurse -Force -ErrorAction SilentlyContinue
  }
}
Invoke-Do "删除安装目录 $InstallDir" {
  Start-Sleep -Milliseconds 500
  Remove-Item -LiteralPath $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
}

Say ""
Say "=== 卸载完成 ===" Green
Say "重新打开 Word / Excel / PowerPoint 后，「PDF 工具」选项卡将不再出现。" Gray
Say "（localhost 开发证书保留在用户证书存储中，如需彻底清除可执行：" Gray
Say "   certutil -user -delstore Root `"Developer CA for Microsoft Office Add-ins`" ）" Gray
