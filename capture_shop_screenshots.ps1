$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogsDir = Join-Path $ScriptDir "screenshots_logs"
New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null
$RunStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$LogFile = Join-Path $LogsDir "run_$RunStamp.log"

Start-Transcript -LiteralPath $LogFile -Force | Out-Null

try {
  Write-Host "Started:" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
  Write-Host "Script folder:" $ScriptDir

  $Workbook = Join-Path $ScriptDir "Обзор цен 19062026 только цены мониторинг.xlsx"

  if (-not (Test-Path -LiteralPath $Workbook)) {
    $Workbook = Get-ChildItem -LiteralPath $ScriptDir -Filter "*мониторинг.xlsx" |
      Where-Object { -not $_.Name.StartsWith("~$") } |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1 -ExpandProperty FullName
  }

  if (-not $Workbook -or -not (Test-Path -LiteralPath $Workbook)) {
    throw "Monitoring workbook not found near the script."
  }

  $Python = "C:\Users\HONOR\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
  $Node = "C:\Users\HONOR\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
  $Edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

  if (-not (Test-Path -LiteralPath $Python)) { throw "Python not found: $Python" }
  if (-not (Test-Path -LiteralPath $Node)) { throw "Node.js not found: $Node" }
  if (-not (Test-Path -LiteralPath $Edge)) { throw "Microsoft Edge not found: $Edge" }

  $ScreensRoot = Join-Path $ScriptDir "screenshots"
  New-Item -ItemType Directory -Path $ScreensRoot -Force | Out-Null

  $ResolvedScriptDir = (Resolve-Path -LiteralPath $ScriptDir).Path
  $ResolvedScreensRoot = (Resolve-Path -LiteralPath $ScreensRoot).Path
  if (-not $ResolvedScreensRoot.StartsWith($ResolvedScriptDir, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Screenshots folder is outside the script folder, cleanup stopped: $ResolvedScreensRoot"
  }

  Write-Host "Cleaning old screenshots before this run..."
  Get-ChildItem -LiteralPath $ScreensRoot -Directory -Filter "screenshots_*" -ErrorAction SilentlyContinue |
    ForEach-Object {
      if ($_.FullName.StartsWith($ResolvedScreensRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force
      }
    }
  Get-ChildItem -LiteralPath $ScreensRoot -File -Filter "shop_links_*.json" -ErrorAction SilentlyContinue |
    ForEach-Object {
      if ($_.FullName.StartsWith($ResolvedScreensRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $_.FullName -Force
      }
    }

  $LinksJson = Join-Path $ScreensRoot "shop_links_$RunStamp.json"

  Write-Host ""
  Write-Host "Excel:" $Workbook
  Write-Host "Screenshots folder:" $ScreensRoot
  Write-Host "Run log:" $LogFile
  Write-Host ""
  Write-Host "Reading shop links from Excel..."
  & $Python (Join-Path $ScriptDir "extract_shop_links.py") $Workbook $LinksJson
  if ($LASTEXITCODE -ne 0) { throw "Could not read shop links from Excel." }

  Write-Host ""
  Write-Host "Opening pages in Microsoft Edge and saving screenshots..."
  Write-Host "Do not close Edge until the script is done."
  Write-Host ""
  & $Node (Join-Path $ScriptDir "capture_shop_screenshots.mjs") $LinksJson $ScreensRoot $Edge
  if ($LASTEXITCODE -ne 0) { throw "Screenshot capture failed." }

  Write-Host ""
  Write-Host "Done. Screenshots are saved in the screenshots folder."
  Write-Host ""
} catch {
  Write-Host ""
  Write-Host "ERROR:"
  Write-Host $_.Exception.Message
  Write-Host ""
  Write-Host "Log file:" $LogFile
  Write-Host ""
} finally {
  Stop-Transcript | Out-Null
  Read-Host "Press Enter to close"
}

