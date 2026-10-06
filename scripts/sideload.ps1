<#
.SYNOPSIS
  在 Windows 桌面版 Office（Word / Excel / PowerPoint）中侧载本加载项。

.DESCRIPTION
  采用“开发者侧载”注册表方式：在
    HKCU\Software\Microsoft\Office\16.0\Wef\Developer
  下写入一个值，值为 manifest.xml 的完整路径。
  写入后重启 Office，即可在“PDF 工具”选项卡看到按钮（也可在 插入 → 我的加载项 中找到）。

  需要先启动本地服务：npm run https

.PARAMETER ManifestPath
  manifest.xml 路径，默认 <项目根>\manifest\manifest.xml

.PARAMETER Port
  本地服务端口，默认 3000（用于提示与端口检查）

.EXAMPLE
  npm run sideload
  powershell -ExecutionPolicy Bypass -File scripts/sideload.ps1 -Port 3001
#>
[CmdletBinding()]
param(
  [string]$ManifestPath = '',
  [int]$Port = 3000,
  [switch]$SharedFolder
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$root = Split-Path -Parent $PSScriptRoot
if (-not $ManifestPath) { $ManifestPath = Join-Path $root 'manifest\manifest.xml' }
$ManifestPath = (Resolve-Path -LiteralPath $ManifestPath).Path

Write-Host "加载项清单: $ManifestPath" -ForegroundColor Cyan

# 1) 检查清单里的端口是否与本地服务一致
[xml]$manifest = Get-Content -LiteralPath $ManifestPath -Raw
$source = $manifest.OfficeApp.DefaultSettings.SourceLocation.DefaultValue
if ($source -and $source -notmatch ":$Port/") {
  Write-Warning "清单中的服务地址为 $source，但当前端口参数为 $Port。可用 npm run set:port 修改。"
}
Write-Host "任务窗格地址: $source"

# 2) 检查本地服务是否在运行
try {
  $resp = Invoke-WebRequest -Uri "https://localhost:$Port/api/health" -SkipCertificateCheck -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop
  Write-Host "本地服务已就绪 (HTTP $($resp.StatusCode))" -ForegroundColor Green
}
catch {
  try {
    $resp = Invoke-WebRequest -Uri "http://localhost:$Port/api/health" -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop
    Write-Host "本地服务以 HTTP 方式运行；Office 加载项需要 HTTPS，请改用 npm run https" -ForegroundColor Yellow
  }
  catch {
    Write-Warning "无法连接 https://localhost:$Port/api/health，请先在项目目录运行: npm run https"
  }
}

if ($SharedFolder) {
  # 共享文件夹 + 受信任目录 方式
  $share = Join-Path $env:USERPROFILE 'OfficeAddinShare'
  New-Item -ItemType Directory -Force -Path $share | Out-Null
  Copy-Item -LiteralPath $ManifestPath -Destination (Join-Path $share 'manifest.xml') -Force
  Write-Host "已复制到共享目录: $share" -ForegroundColor Green
  Write-Host "接下来在 Office 中：文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录，添加该目录（勾选“显示在菜单中”），"
  Write-Host "然后重启 Office，在 插入 → 我的加载项 → 共享文件夹 中插入本加载项。"
  exit 0
}

# 3) 写注册表（开发者侧载）
$key = 'HKCU:\Software\Microsoft\Office\16.0\Wef\Developer'
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key -Name $ManifestPath -Value $ManifestPath -PropertyType String -Force | Out-Null
Write-Host "已写入注册表: $key" -ForegroundColor Green
Write-Host "  值名: $ManifestPath"
Write-Host "  值数据: $ManifestPath"

Write-Host ""
Write-Host "接下来：" -ForegroundColor Cyan
Write-Host "  1. 完全退出并重新打开 Word / Excel / PowerPoint（托盘里不能有残留进程）"
Write-Host "  2. 新建或打开文档，在功能区找到 “PDF 工具” 选项卡"
Write-Host "  3. 若未出现，检查 文件 → 选项 → 信任中心 → 受信任的加载项目录，确认允许加载项"
Write-Host "  卸载：npm run unsideload"
