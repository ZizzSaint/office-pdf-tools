<#
.SYNOPSIS
  取消侧载（删除开发者注册表项）。
#>
[CmdletBinding()]
param(
  [string]$ManifestPath = ''
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
$root = Split-Path -Parent $PSScriptRoot
if (-not $ManifestPath) { $ManifestPath = Join-Path $root 'manifest\manifest.xml' }
$ManifestPath = (Resolve-Path -LiteralPath $ManifestPath -ErrorAction SilentlyContinue).Path
if (-not $ManifestPath) { $ManifestPath = Join-Path $root 'manifest\manifest.xml' }

$key = 'HKCU:\Software\Microsoft\Office\16.0\Wef\Developer'
if (Test-Path $key) {
  $props = (Get-Item $key).Property
  foreach ($name in $props) {
    if ($name -like '*manifest.xml' -or $name -eq $ManifestPath) {
      Remove-ItemProperty -Path $key -Name $name -Force -ErrorAction SilentlyContinue
      Write-Host "已删除注册表值: $name" -ForegroundColor Green
    }
  }
  if (-not (Get-Item $key).Property.Count) {
    Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
    Write-Host "注册表项已清空并删除: $key"
  }
}
else {
  Write-Host "未找到注册表项 $key（可能尚未侧载）"
}

Write-Host "请重启 Office 使改动生效。"
