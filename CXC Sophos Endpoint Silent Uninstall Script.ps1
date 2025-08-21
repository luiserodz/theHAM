# Sophos Endpoint Silent Uninstall Script
# For deployment via Action1 and Microsoft Intune

# Set error action preference
$ErrorActionPreference = "Stop"

# Define Sophos installation path
$sophosPath = "C:\Program Files\Sophos\Sophos Endpoint Agent"
$uninstallerPath = Join-Path $sophosPath "SophosUninstall.exe"

# Logging function
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    Write-Output $logMessage
    # Optional: Write to file for tracking
    # Add-Content -Path "C:\Windows\Temp\SophosUninstall.log" -Value $logMessage
}

try {
    Write-Log "Starting Sophos Endpoint uninstallation process"
    
    # Check if Sophos is installed
    if (Test-Path $uninstallerPath) {
        Write-Log "Sophos uninstaller found at: $uninstallerPath"
        
        # Check if Sophos services are running
        $sophosServices = Get-Service -Name "Sophos*" -ErrorAction SilentlyContinue
        if ($sophosServices) {
            Write-Log "Found $($sophosServices.Count) Sophos services"
        }
        
        # Execute silent uninstall
        Write-Log "Executing silent uninstall"
        $uninstallProcess = Start-Process -FilePath $uninstallerPath -ArgumentList "--quiet" -Wait -PassThru -NoNewWindow
        
        # Check exit code
        if ($uninstallProcess.ExitCode -eq 0) {
            Write-Log "Sophos uninstalled successfully"
            exit 0
        }
        elseif ($uninstallProcess.ExitCode -eq 3010) {
            Write-Log "Sophos uninstalled successfully - Reboot required"
            exit 3010
        }
        else {
            Write-Log "Uninstall completed with exit code: $($uninstallProcess.ExitCode)"
            exit $uninstallProcess.ExitCode
        }
    }
    else {
        Write-Log "Sophos uninstaller not found at expected location"
        
        # Alternative: Check if Sophos is installed via registry
        $sophosRegistry = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" | 
                         Where-Object { $_.DisplayName -like "*Sophos*Endpoint*" }
        
        if ($sophosRegistry) {
            Write-Log "Sophos found in registry but uninstaller missing"
            exit 1
        }
        else {
            Write-Log "Sophos Endpoint not detected on this system"
            exit 0
        }
    }
}
catch {
    Write-Log "Error during uninstallation: $_"
    exit 1
}