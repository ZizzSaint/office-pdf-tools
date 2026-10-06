#Requires -Version 5.1
<#
  Office COM 转 PDF。
  用法: powershell -NoProfile -ExecutionPolicy Bypass -File scripts/office-com.ps1 -InputPath a.docx -OutputPath a.pdf [-App auto|word|excel|powerpoint] [-Sheet "Sheet1"] [-PidFile pid.txt]
  成功输出: {"ok":true,...}  失败输出: {"ok":false,"error":"..."} 并返回退出码 1。
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputPath,
  [ValidateSet('auto', 'word', 'excel', 'powerpoint')][string]$App = 'auto',
  [string]$Sheet = '',
  [switch]$Recalculate,
  [string]$PidFile = ''
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
$ProgressPreference = 'SilentlyContinue'

function New-Result([bool]$ok, [string]$errorMessage) {
  $r = [ordered]@{
    ok      = $ok
    engine  = 'msoffice'
    app     = $script:appKey
    output  = $OutputPath
    error   = $errorMessage
    pid     = $script:officePid
  }
  return ($r | ConvertTo-Json -Compress)
}

function Get-AppKey([string]$file) {
  switch ([System.IO.Path]::GetExtension($file).ToLowerInvariant()) {
    '.doc' { 'word' } '.docx' { 'word' } '.docm' { 'word' } '.dot' { 'word' } '.dotx' { 'word' } '.rtf' { 'word' } '.odt' { 'word' } '.txt' { 'word' }
    '.xls' { 'excel' } '.xlsx' { 'excel' } '.xlsm' { 'excel' } '.xlsb' { 'excel' } '.ods' { 'excel' } '.csv' { 'excel' }
    '.ppt' { 'powerpoint' } '.pptx' { 'powerpoint' } '.pptm' { 'powerpoint' } '.odp' { 'powerpoint' }
    default { '' }
  }
}

function Get-Pids([string]$name) {
  @(Get-Process -Name $name -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
}

function Get-OrCreate-App([string]$progId, [string]$procName, [int[]]$before) {
  # PowerPoint 只允许单实例；若已有实例则复用（既不打断用户，也避免 CO_E_SERVER_EXEC_FAILURE）。
  # 复用时不修改 Visible/DisplayAlerts，也不会在结束时 Quit。
  try {
    $existing = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
    if ($existing) {
      $script:createdApp = $false
      return $existing
    }
  } catch { }
  $created = New-Object -ComObject $progId
  $script:createdApp = $true
  Set-OfficePid $procName $before
  return $created
}

function Set-OfficePid([string]$procName, [int[]]$before) {
  try {
    $after = Get-Pids $procName
    $new = $after | Where-Object { $before -notcontains $_ } | Select-Object -First 1
    if ($new) {
      $script:officePid = [int]$new
      if ($PidFile) { Set-Content -LiteralPath $PidFile -Value $new -Encoding ASCII }
    }
  } catch { }
}

$script:appKey = if ($App -eq 'auto') { Get-AppKey $InputPath } else { $App }
$script:officePid = 0
$script:createdApp = $false

if (-not $script:appKey) {
  Write-Output (New-Result $false "不支持的文件类型: $([System.IO.Path]::GetExtension($InputPath))")
  exit 1
}
if (-not (Test-Path -LiteralPath $InputPath)) {
  Write-Output (New-Result $false "文件不存在: $InputPath")
  exit 1
}

$outDir = Split-Path -Parent $OutputPath
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$officeApp = $null
$officeDoc = $null
$procName = @{ word = 'WINWORD'; excel = 'EXCEL'; powerpoint = 'POWERPNT' }[$script:appKey]

try {
  $before = Get-Pids $procName

  if ($script:appKey -eq 'word') {
    $officeApp = Get-OrCreate-App 'Word.Application' $procName $before
    if ($script:createdApp) {
      $officeApp.Visible = $false
      $officeApp.DisplayAlerts = 0
      try { $officeApp.Options.SaveNormalPrompt = $false } catch { }
    }
    $officeDoc = $officeApp.Documents.Open($InputPath, $false, $true, $false)
    $officeDoc.ExportAsFixedFormat($OutputPath, 17)   # wdExportFormatPDF
    $officeDoc.Close(0)
    $officeDoc = $null
  }
  elseif ($script:appKey -eq 'excel') {
    $officeApp = Get-OrCreate-App 'Excel.Application' $procName $before
    if ($script:createdApp) {
      $officeApp.Visible = $false
      $officeApp.DisplayAlerts = $false
      $officeApp.ScreenUpdating = $false
    }
    $officeDoc = $officeApp.Workbooks.Open($InputPath, 0, $true)
    if ($Recalculate) { try { $officeApp.CalculateFullRebuild() } catch { } }
    if ($Sheet) {
      $ws = $officeDoc.Worksheets.Item($Sheet)
      $ws.ExportAsFixedFormat(0, $OutputPath)   # xlTypePDF
    }
    else {
      $officeDoc.ExportAsFixedFormat(0, $OutputPath)
    }
    $officeDoc.Close($false)
    $officeDoc = $null
  }
  else {
    $officeApp = Get-OrCreate-App 'PowerPoint.Application' $procName $before
    if ($script:createdApp) {
      try { $officeApp.DisplayAlerts = 1 } catch { }   # ppAlertsNone
    }
    $officeDoc = $officeApp.Presentations.Open($InputPath, $true, $false, $false)  # ReadOnly, Untitled:=False, WithWindow:=False
    $officeDoc.SaveAs($OutputPath, 32)   # ppSaveAsPDF
    $officeDoc.Close()
    $officeDoc = $null
  }

  if (-not (Test-Path -LiteralPath $OutputPath)) {
    throw "转换结束但未生成输出文件: $OutputPath"
  }
  Write-Output (New-Result $true '')
  exit 0
}
catch {
  Write-Output (New-Result $false $_.Exception.Message)
  exit 1
}
finally {
  if ($officeDoc) { try { $officeDoc.Close($false) } catch { } }
  if ($officeApp) {
    if ($script:createdApp) { try { $officeApp.Quit() } catch { } }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($officeApp) | Out-Null } catch { }
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
