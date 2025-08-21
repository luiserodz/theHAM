<#
    Sophos Endpoint Silent Uninstall Script
    This script attempts to remove all Sophos components silently so it can be
    deployed through Microsoft Intune.  The original version only targeted a
    single installation path.  The improved version searches for known
    uninstallers and falls back to the registry to remove any remaining Sophos
    products.  Exit codes are returned so Intune can detect success (0) or
    reboot required (3010).
#>

$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Message)
    $timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    Write-Output $logMessage
    # Add-Content -Path "C:\Windows\Temp\SophosUninstall.log" -Value $logMessage
}

function Invoke-Uninstall {
    param(
        [string]$FilePath,
        [string]$Arguments = "",
        [switch]$UseCmd
    )

    try {
        if ($UseCmd) {
            $Arguments = "/c `"$FilePath $Arguments`""
            $FilePath  = "$env:SystemRoot\System32\cmd.exe"
        }

        Write-Log "Running: $FilePath $Arguments"
        $proc = Start-Process -FilePath $FilePath -ArgumentList $Arguments -NoNewWindow -Wait -PassThru
        Write-Log "Exit code: $($proc.ExitCode)"
        return $proc.ExitCode
    }
    catch {
        Write-Log "Failed to run $FilePath : $_"
        return 1
    }
}

Write-Log "Starting Sophos Endpoint uninstallation process"

Get-Service -Name "Sophos*" -ErrorAction SilentlyContinue | Stop-Service -Force -ErrorAction SilentlyContinue
Get-Process -Name "Sophos*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$rebootNeeded = $false

# Known uninstall executables
$uninstallers = @(
    @{ Path = "C:\Program Files\Sophos\Sophos Endpoint Agent\SophosUninstall.exe"; Args = "--quiet" },
    @{ Path = "C:\Program Files (x86)\Sophos\Sophos Endpoint Agent\SophosUninstall.exe"; Args = "--quiet" },
    @{ Path = "C:\Program Files\Sophos\Endpoint Defense\uninstallcli.exe"; Args = "--quiet" },
    @{ Path = "C:\Program Files (x86)\Sophos\Endpoint Defense\uninstallcli.exe"; Args = "--quiet" }
)

foreach ($u in $uninstallers) {
    if (Test-Path $u.Path) {
        $exit = Invoke-Uninstall -FilePath $u.Path -Arguments $u.Args
        if ($exit -eq 3010) { $rebootNeeded = $true }
    }
}

# Registry-based uninstall fall-back
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

foreach ($regPath in $regPaths) {
    Get-ChildItem $regPath -ErrorAction SilentlyContinue | ForEach-Object {
        $item = Get-ItemProperty $_.PSPath
        if ($item.DisplayName -like "*Sophos*") {
            Write-Log "Found registry entry for $($item.DisplayName)"
            $cmd = $item.UninstallString
            if ([string]::IsNullOrWhiteSpace($cmd)) { return }

            if ($cmd -match "msiexec") {
                $args = $cmd.Substring($cmd.IndexOf("msiexec") + 8).Trim()
                if ($args -notmatch "/qn") { $args = "$args /qn /norestart" }
                $exit = Invoke-Uninstall -FilePath "$env:SystemRoot\System32\msiexec.exe" -Arguments $args
            }
            else {
                $exit = Invoke-Uninstall -FilePath $cmd -Arguments "/quiet"
            }

            if ($exit -eq 3010) { $rebootNeeded = $true }
        }
    }
}

if ($rebootNeeded) {
    Write-Log "Uninstallation complete - reboot required"
    exit 3010
}
else {
    Write-Log "Uninstallation complete"
    exit 0
}