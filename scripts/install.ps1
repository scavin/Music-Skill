$ErrorActionPreference = "Stop"

$BaseDir  = Join-Path $env:USERPROFILE ".openclaw\music"
$Bin      = Join-Path $BaseDir "go-music-api.exe"
$LogFile  = Join-Path $BaseDir "log.txt"
$PidFile  = Join-Path $BaseDir "pid"
$PortFile = Join-Path $BaseDir "port"
$ApiUrl   = "https://api.github.com/repos/guohuiyuan/go-music-api/releases/latest"

if (-not (Test-Path $BaseDir)) { New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null }

Write-Host "[music] base dir: $BaseDir"

function Test-PidRunning {
    if (-not (Test-Path $PidFile)) { return $false }
    $pid = (Get-Content $PidFile -ErrorAction SilentlyContinue).Trim()
    if (-not $pid) { return $false }
    try {
        $proc = Get-Process -Id ([int]$pid) -ErrorAction Stop
        return ($null -ne $proc)
    } catch {
        return $false
    }
}

function Test-PortBusy([int]$Port) {
    try {
        $conns = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop
        return ($conns.Count -gt 0)
    } catch {
        return $false
    }
}

function Find-FreePort {
    foreach ($p in 8080, 8081, 8090, 18080, 28080) {
        if (-not (Test-PortBusy $p)) { return $p }
    }
    return $null
}

function Test-Health([int]$Port) {
    try {
        $resp = Invoke-WebRequest -Uri "http://localhost:$Port/api/v1/music/search?q=test" -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        return ($resp.StatusCode -eq 200)
    } catch {
        return $false
    }
}

# Clean stale PID
if ((Test-Path $PidFile) -and -not (Test-PidRunning)) {
    Remove-Item $PidFile -Force -ErrorAction SilentlyContinue
}

# Check if already running
if (Test-PidRunning) {
    if (Test-Path $PortFile) {
        $runPort = (Get-Content $PortFile -ErrorAction SilentlyContinue).Trim()
        if ($runPort -and (Test-Health ([int]$runPort))) {
            Write-Host "[music] service already running on port $runPort"
            exit 0
        }
    }
}

# Download binary if not present
if (-not (Test-Path $Bin)) {
    Write-Host "[music] fetching latest release info..."
    $json = Invoke-WebRequest -Uri $ApiUrl -UseBasicParsing -ErrorAction Stop | Select-Object -ExpandProperty Content
    $release = $json | ConvertFrom-Json

    $AssetName = "go-music-api_windows_amd64.zip"
    Write-Host "[music] target: windows-amd64"
    Write-Host "[music] asset: $AssetName"

    $asset = $release.assets | Where-Object { $_.name -eq $AssetName } | Select-Object -First 1
    if (-not $asset) {
        Write-Host "[music] failed to find release asset: $AssetName"
        exit 1
    }
    $downloadUrl = $asset.browser_download_url

    $pkgPath = Join-Path $BaseDir $AssetName
    $extractDir = Join-Path $BaseDir "extract"
    if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
    New-Item -ItemType Directory -Path $extractDir -Force | Out-Null

    Write-Host "[music] downloading: $downloadUrl"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $pkgPath -UseBasicParsing -ErrorAction Stop

    Expand-Archive -Path $pkgPath -DestinationPath $extractDir -Force

    $found = Get-ChildItem -Path $extractDir -Recurse -Filter "go-music-api.exe" | Select-Object -First 1
    if (-not $found) {
        Write-Host "[music] binary not found after extraction"
        exit 1
    }

    # Validate PE header (MZ signature)
    $bytes = [System.IO.File]::ReadAllBytes($found.FullName)
    if ($bytes.Length -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        Write-Host "[music] extracted file is not a valid Windows executable: $($found.FullName)"
        exit 1
    }

    Copy-Item $found.FullName $Bin -Force
    Remove-Item $pkgPath -Force -ErrorAction SilentlyContinue
    Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue
}

# Find free port and start
$port = Find-FreePort
if (-not $port) {
    Write-Host "[music] no free port found"
    exit 1
}

Set-Content -Path $PortFile -Value $port -NoNewline

Write-Host "[music] starting service on port $port..."
$proc = Start-Process -FilePath $Bin -WindowStyle Hidden -RedirectStandardOutput $LogFile -RedirectStandardError (Join-Path $BaseDir "err.txt") -PassThru
Set-Content -Path $PidFile -Value $proc.Id -NoNewline

Start-Sleep -Seconds 2

if (Test-Health $port) {
    Write-Host "[music] installed and running on http://localhost:$port"
    exit 0
}

Write-Host "[music] failed to start, check log: $LogFile"
exit 1
