# morning_startup.ps1 - Bring up the two local tools every morning.
#   * cockpit-php   : MySQL (3306) + PHP built-in server (8580)
#   * rcabinet      : Streamlit app (8501)
# Registered as scheduled task "MorningTools" (At log on + Daily 09:00, catch-up on miss).
# Idempotent: each component is started ONLY if its port is not already listening,
# so running twice (logon + daily) is harmless. ASCII only on purpose (Task Scheduler safe).

$ErrorActionPreference = 'SilentlyContinue'

$Python   = 'C:\Users\ssasa\AppData\Local\Programs\Python\Python312\python.exe'
$Php      = 'C:\xampp\php\php.exe'
$CockpitRoot = 'C:\Users\ssasa\cockpit-php\public'
$RcabinetDir = 'C:\Users\ssasa\tools\rcabinet-checker'
$LogFile  = 'C:\Users\ssasa\tools\morning-startup\logs\morning_startup.log'

function Log([string]$msg) {
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Add-Content -Path $LogFile -Value "$ts  $msg"
}

# Ask the OS listener table directly (covers IPv4 and IPv6). PHP binds ::1 only, so a
# plain TcpClient (IPv4) would falsely report it "down" and spawn duplicates.
function Test-Port([int]$port) {
    $conns = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
    return [bool]$conns
}

function Wait-Port([int]$port, [int]$seconds) {
    for ($i = 0; $i -lt $seconds; $i++) {
        Start-Sleep -Seconds 1
        if (Test-Port $port) { return $true }
    }
    return $false
}

Log "----- morning startup run -----"

# 1) MySQL / MariaDB (XAMPP) on 3306 -- needed by the cockpit and the 10:00 import.
# Prefer the Windows service "mysql" (graceful shutdown at boot). Bare standalone with
# built-in defaults corrupted listings.ibd twice, so fall back to standalone ONLY with my.ini.
if (-not (Test-Port 3306)) {
    $svc = Get-Service -Name 'mysql' -ErrorAction SilentlyContinue
    if ($svc) {
        try { Start-Service -Name 'mysql' -ErrorAction Stop; Log "MySQL service started" } catch { Log "MySQL service start FAILED: $_" }
    } else {
        Start-Process -FilePath 'C:\xampp\mysql\bin\mysqld.exe' `
            -ArgumentList '--defaults-file=C:\xampp\mysql\bin\my.ini', '--standalone' -WindowStyle Hidden
        Log "MySQL standalone started (service missing)"
    }
    Wait-Port 3306 30 | Out-Null
} else {
    Log "MySQL already up (3306)"
}

# 2) PHP built-in server for the cockpit on 8580.
$openCockpit = $false
if (-not (Test-Port 8580)) {
    Start-Process -FilePath $Php `
        -ArgumentList '-S', 'localhost:8580', '-t', $CockpitRoot -WindowStyle Hidden
    if (Wait-Port 8580 15) { Log "cockpit started (8580)"; $openCockpit = $true }
    else { Log "cockpit did NOT come up (8580)" }
} else {
    Log "cockpit already up (8580)"
}

# 3) Streamlit rcabinet-checker on 8501. Headless so it does not pop its own browser tab
# (we control tab opening below). WorkingDirectory must be the app folder so .streamlit/
# secrets.toml and config.toml are picked up.
$openRcabinet = $false
if (-not (Test-Port 8501)) {
    Start-Process -FilePath $Python `
        -ArgumentList '-m', 'streamlit', 'run', 'streamlit_app.py', '--server.headless', 'true', '--server.port', '8501' `
        -WorkingDirectory $RcabinetDir -WindowStyle Hidden
    if (Wait-Port 8501 45) { Log "rcabinet started (8501)"; $openRcabinet = $true }
    else { Log "rcabinet did NOT come up (8501)" }
} else {
    Log "rcabinet already up (8501)"
}

# 4) Open browser tabs ONLY for components we just started (avoid tab spam when already up).
if ($openCockpit)  { Start-Process 'http://localhost:8580/'; Log "opened browser: cockpit" }
if ($openRcabinet) { Start-Process 'http://localhost:8501/'; Log "opened browser: rcabinet" }

Log "----- done -----"
