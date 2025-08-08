# Browser Remediation Script for Intune
# Purpose: Installs or updates browsers based on detection results
# Handles user notifications and respects running processes
# Version: 2.0

[CmdletBinding()]
param(
    [Parameter()]
    [switch]$ForceUpdate,
    
    [Parameter()]
    [int]$MaxWaitMinutes = 30,
    
    [Parameter()]
    [switch]$ScheduleRebootIfNeeded,
    
    [Parameter()]
    [string[]]$PreferredBrowsers = @("Edge", "Chrome", "Firefox")
)

$ErrorActionPreference = 'Stop'

# Configuration
$downloadUrls = @{
    Chrome = "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7BBD050720-A9FA-FA92-D25C-940523BD2DC1%7D%26browser%3D0%26usagestats%3D1%26appname%3DGoogle%2520Chrome%26needsadmin%3Dtrue%26ap%3Dx64-stable-statsdef_0%26brand%3DGCEA/dl/chrome/install/googlechromestandaloneenterprise64.msi"
    Edge = "https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/5838c4cb-c4e1-40ec-af9a-0a5f8cd78e24/MicrosoftEdgeEnterpriseX64.msi"
    Firefox = "https://download.mozilla.org/pub/firefox/releases/latest/win64/en-US/Firefox%20Setup%20Latest.msi"
}

$processNames = @{
    Chrome = @("chrome", "GoogleUpdate")
    Edge = @("msedge", "MicrosoftEdge", "MicrosoftEdgeUpdate")
    Firefox = @("firefox", "firefox-updater")
}

# Performance monitoring
$script:performanceMetrics = @{
    StartTime = Get-Date
    Downloads = @()
    Installations = @()
    Notifications = @()
}

function Write-RemediationLog {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$Telemetry
    )
    
    $logPath = "$env:ProgramData\BrowserManager\remediation.log"
    $logDir = Split-Path $logPath -Parent
    
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    # Rotate log if too large
    if ((Test-Path $logPath) -and ((Get-Item $logPath).Length -gt 10MB)) {
        $archivePath = $logPath -replace '\.log

function Show-UserNotification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$Type = "Information",
        [string[]]$Actions = @(),
        [string]$Tag = "BrowserManager",
        [string]$Group = "BrowserUpdates"
    )
    
    try {
        # Primary method: Windows Toast Notifications
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
        
        # Determine icon based on type
        $icon = switch ($Type) {
            "Warning" { "⚠️" }
            "Error" { "❌" }
            "Success" { "✅" }
            default { "ℹ️" }
        }
        
        # Build toast XML
        $titleText = "$icon $Title"
        $template = @"
<toast scenario="reminder" activationType="protocol" launch="action=browserupdate">
    <visual>
        <binding template="ToastGeneric">
            <text>$titleText</text>
            <text>$Message</text>
            <image placement="appLogoOverride" src="https://raw.githubusercontent.com/microsoft/PowerToys/main/src/settings-ui/Settings.UI/Assets/FluentIcons/FluentIconsBrowsers.png" />
        </binding>
    </visual>
"@
        
        if ($Actions.Count -gt 0) {
            $template += "<actions>"
            foreach ($action in $Actions) {
                $actionXml = "<action content='$action' arguments='action=$action' activationType='protocol'/>"
                $template += $actionXml
            }
            $template += "</actions>"
        }
        
        $template += @"
    <audio src="ms-winsoundevent:Notification.Default" />
</toast>
"@
        
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml($template)
        
        $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
        $toast.Tag = $Tag
        $toast.Group = $Group
        
        # Create notifier with app ID
        $appId = "Microsoft.Windows.Explorer"  # Use Explorer as it is always available
        $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId)
        $notifier.Show($toast)
        
        Write-RemediationLog -Message "Toast notification shown: $Title - $Message"
        return $true
    }
    catch {
        Write-RemediationLog -Message "Toast notification failed, trying fallback: $_" -Level "WARNING"
        
        # Fallback 1: Try WScript.Shell popup for user context
        try {
            $wshell = New-Object -ComObject Wscript.Shell
            $iconType = switch ($Type) {
                "Warning" { 48 }  # Exclamation
                "Error" { 16 }    # Critical
                "Success" { 64 }  # Information
                default { 64 }    # Information
            }
            $wshell.Popup($Message, 30, $Title, $iconType) | Out-Null
            Write-RemediationLog -Message "WScript popup shown"
            return $true
        }
        catch {
            Write-RemediationLog -Message "WScript popup failed: $_" -Level "WARNING"
        }
        
        # Fallback 2: Windows Forms balloon tip
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
            
            $notification = New-Object System.Windows.Forms.NotifyIcon
            $notification.Icon = [System.Drawing.SystemIcons]::Information
            $notification.BalloonTipIcon = $Type
            $notification.BalloonTipTitle = $Title
            $notification.BalloonTipText = $Message
            $notification.Visible = $true
            $notification.ShowBalloonTip(30000)
            
            Start-Sleep -Seconds 5
            $notification.Dispose()
            
            Write-RemediationLog -Message "Balloon tip shown"
            return $true
        }
        catch {
            Write-RemediationLog -Message "Balloon tip failed: $_" -Level "WARNING"
        }
        
        # Fallback 3: msg.exe for system context
        try {
            $sessions = @(query user 2>$null | Select-Object -Skip 1 | ForEach-Object {
                $parts = $_ -split '\s+'
                if ($parts.Count -gt 3 -and $parts[3] -eq "Active") { 
                    $parts[2] 
                }
            })
            
            foreach ($sessionId in $sessions) {
                $simpleMessage = "$Title : $Message" -replace "`n", " "
                msg.exe $sessionId /TIME:30 $simpleMessage 2>$null
            }
            Write-RemediationLog -Message "msg.exe notification sent"
            return $true
        }
        catch {
            Write-RemediationLog -Message "All notification methods failed: $_" -Level "ERROR"
            return $false
        }
    }
}

function Test-BrowserRunning {
    param(
        [string]$BrowserName
    )
    
    $processes = $processNames[$BrowserName]
    if (-not $processes) { return $false }
    
    $runningProcesses = @()
    
    foreach ($processName in $processes) {
        $running = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($running) {
            $runningProcesses += $running
        }
    }
    
    if ($runningProcesses.Count -gt 0) {
        # Check if any have active windows (user is actively using)
        $activeWindows = $runningProcesses | Where-Object { $_.MainWindowTitle -ne "" }
        
        Write-RemediationLog -Message "$BrowserName is running with $($runningProcesses.Count) process(es), $($activeWindows.Count) with active windows"
        
        return @{
            IsRunning = $true
            Processes = $runningProcesses
            HasActiveWindows = $activeWindows.Count -gt 0
            ProcessCount = $runningProcesses.Count
        }
    }
    
    return @{
        IsRunning = $false
        Processes = @()
        HasActiveWindows = $false
        ProcessCount = 0
    }
}

function Wait-ForBrowserClose {
    param(
        [string]$BrowserName,
        [int]$MaxWaitMinutes = 30,
        [int]$CheckIntervalSeconds = 30
    )
    
    $endTime = (Get-Date).AddMinutes($MaxWaitMinutes)
    $notificationCount = 0
    $maxNotifications = 3
    
    while ((Get-Date) -lt $endTime) {
        $status = Test-BrowserRunning -BrowserName $BrowserName
        
        if (-not $status.IsRunning) {
            Write-RemediationLog -Message "$BrowserName has been closed by user"
            return $true
        }
        
        # Send periodic reminders
        if ($notificationCount -lt $maxNotifications) {
            $remainingMinutes = [math]::Round(($endTime - (Get-Date)).TotalMinutes)
            
            if ($status.HasActiveWindows) {
                Show-UserNotification -Title "Browser Update Waiting" `
                    -Message "$BrowserName needs to be closed for an important security update. Will wait $remainingMinutes more minutes." `
                    -Type "Warning" `
                    -Actions @("Close $BrowserName Now", "Remind Me Later")
            }
            
            $notificationCount++
        }
        
        Start-Sleep -Seconds $CheckIntervalSeconds
    }
    
    Write-RemediationLog -Message "Timeout waiting for $BrowserName to close" -Level "WARNING"
    return $false
}

function Get-BrowserInstaller {
    param(
        [string]$BrowserName,
        [string]$Url
    )
    
    $tempDir = "$env:TEMP\BrowserInstallers"
    if (-not (Test-Path $tempDir)) {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    }
    
    $installerPath = Join-Path $tempDir "$BrowserName.msi"
    
    # Check if we have a recent cached installer (less than 1 hour old)
    if (Test-Path $installerPath) {
        $fileAge = (Get-Date) - (Get-Item $installerPath).LastWriteTime
        if ($fileAge.TotalHours -lt 1) {
            Write-RemediationLog -Message "Using cached installer for $BrowserName (age: $([math]::Round($fileAge.TotalMinutes)) minutes)"
            return $installerPath
        }
        else {
            Remove-Item $installerPath -Force
        }
    }
    
    try {
        $downloadStart = Get-Date
        Write-RemediationLog -Message "Downloading $BrowserName installer..."
        
        # Test network connectivity first
        $testConnection = Test-NetConnection -ComputerName "www.google.com" -Port 443 -WarningAction SilentlyContinue
        if (-not $testConnection.TcpTestSucceeded) {
            throw "Network connectivity test failed"
        }
        
        # Use BITS for reliable download
        $bitsJob = Start-BitsTransfer -Source $Url -Destination $installerPath `
            -DisplayName "Download $BrowserName" -Priority Normal -Asynchronous
        
        # Wait for download with progress tracking
        $timeout = (Get-Date).AddMinutes(10)
        $lastProgress = 0
        
        while ((Get-Date) -lt $timeout) {
            $job = Get-BitsTransfer -JobId $bitsJob.JobId -ErrorAction SilentlyContinue
            if (-not $job) { break }
            
            if ($job.JobState -eq "Transferred") {
                Complete-BitsTransfer -BitsJob $job
                $downloadTime = ((Get-Date) - $downloadStart).TotalSeconds
                Write-RemediationLog -Message "$BrowserName downloaded successfully in $([math]::Round($downloadTime)) seconds" -Telemetry
                
                # Verify file
                if ((Get-Item $installerPath).Length -lt 1MB) {
                    throw "Downloaded file is too small, likely corrupted"
                }
                
                $script:performanceMetrics.Downloads += @{
                    Browser = $BrowserName
                    Duration = $downloadTime
                    Success = $true
                }
                
                return $installerPath
            }
            elseif ($job.JobState -eq "Error") {
                $errorDetail = $job.ErrorDescription
                Remove-BitsTransfer -BitsJob $job
                throw "BITS transfer failed: $errorDetail"
            }
            elseif ($job.JobState -eq "Transferring") {
                $progress = [math]::Round(($job.BytesTransferred / $job.BytesTotal) * 100)
                if ($progress -ne $lastProgress -and ($progress % 25 -eq 0)) {
                    Write-RemediationLog -Message "$BrowserName download progress: $progress%"
                    $lastProgress = $progress
                }
            }
            
            Start-Sleep -Seconds 5
        }
        
        # Fallback to Invoke-WebRequest
        Write-RemediationLog -Message "BITS timeout, using direct download..." -Level "WARNING"
        
        # Use .NET WebClient for better performance
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        
        # Add progress tracking
        $progressPreference = 'SilentlyContinue'
        $webClient.DownloadFile($Url, $installerPath)
        
        if (Test-Path $installerPath) {
            $downloadTime = ((Get-Date) - $downloadStart).TotalSeconds
            Write-RemediationLog -Message "$BrowserName downloaded successfully (fallback method) in $([math]::Round($downloadTime)) seconds"
            
            $script:performanceMetrics.Downloads += @{
                Browser = $BrowserName
                Duration = $downloadTime
                Success = $true
                Method = "Fallback"
            }
            
            return $installerPath
        }
    }
    catch {
        Write-RemediationLog -Message "Failed to download $BrowserName : $_" -Level "ERROR" -Telemetry
        
        $script:performanceMetrics.Downloads += @{
            Browser = $BrowserName
            Success = $false
            Error = $_.Exception.Message
        }
        
        # Try alternative URLs for specific browsers
        $alternativeUrls = @{
            Firefox = @(
                "https://download-installer.cdn.mozilla.net/pub/firefox/releases/latest/win64/en-US/Firefox%20Setup%20Latest.msi",
                "https://ftp.mozilla.org/pub/firefox/releases/latest/win64/en-US/Firefox%20Setup%20Latest.msi"
            )
            Chrome = @(
                "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
            )
        }
        
        if ($alternativeUrls.ContainsKey($BrowserName)) {
            foreach ($altUrl in $alternativeUrls[$BrowserName]) {
                try {
                    Write-RemediationLog -Message "Trying alternative URL for $BrowserName"
                    Invoke-WebRequest -Uri $altUrl -OutFile $installerPath -UseBasicParsing -TimeoutSec 60
                    if (Test-Path $installerPath) {
                        return $installerPath
                    }
                }
                catch {
                    continue
                }
            }
        }
    }
    
    return $null
}

function Install-Browser {
    param(
        [string]$BrowserName,
        [string]$InstallerPath,
        [bool]$IsUpdate = $false
    )
    
    if (-not (Test-Path $InstallerPath)) {
        Write-RemediationLog -Message "Installer not found: $InstallerPath" -Level "ERROR"
        return $false
    }
    
    $action = if ($IsUpdate) { "update" } else { "install" }
    $installStart = Get-Date
    Write-RemediationLog -Message "Starting $action for $BrowserName..."
    
    try {
        # Verify installer signature
        $signature = Get-AuthenticodeSignature -FilePath $InstallerPath
        if ($signature.Status -ne "Valid") {
            Write-RemediationLog -Message "Invalid signature for $BrowserName installer: $($signature.Status)" -Level "WARNING"
        }
        
        # MSI arguments for silent installation
        $logFile = "$env:ProgramData\BrowserManager\Install_${BrowserName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        $arguments = @(
            "/i",
            "`"$InstallerPath`"",
            "/quiet",
            "/norestart",
            "/l*v",
            "`"$logFile`""
        )
        
        # Add browser-specific parameters
        switch ($BrowserName) {
            "Chrome" {
                $arguments += "REBOOT=ReallySuppress"
                $arguments += "ALLUSERS=1"
                $arguments += "NOGOOGLEUPDATEPING=1"
            }
            "Edge" {
                $arguments += "REBOOT=ReallySuppress"
                $arguments += "DONOTCREATEDESKTOPSHORTCUT=TRUE"
                $arguments += "DONOTCREATETASKBARSHORTCUT=TRUE"
            }
            "Firefox" {
                $arguments += "DESKTOP_SHORTCUT=false"
                $arguments += "INSTALL_MAINTENANCE_SERVICE=false"
                $arguments += "PREVENT_REBOOT_REQUIRED=true"
            }
        }
        
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments `
            -Wait -PassThru -WindowStyle Hidden
        
        $installTime = ((Get-Date) - $installStart).TotalSeconds
        
        $script:performanceMetrics.Installations += @{
            Browser = $BrowserName
            Duration = $installTime
            ExitCode = $process.ExitCode
            Action = $action
        }
        
        if ($process.ExitCode -eq 0) {
            Write-RemediationLog -Message "$BrowserName $action completed successfully in $([math]::Round($installTime)) seconds" -Telemetry
            
            # Verify installation
            Start-Sleep -Seconds 2
            if (Test-BrowserInstalled -BrowserName $BrowserName) {
                Write-RemediationLog -Message "$BrowserName installation verified"
                return $true
            }
            else {
                Write-RemediationLog -Message "$BrowserName installation succeeded but verification failed" -Level "WARNING"
                return $true  # Trust the exit code
            }
        }
        elseif ($process.ExitCode -eq 3010) {
            Write-RemediationLog -Message "$BrowserName $action completed, restart required (3010)" -Level "WARNING" -Telemetry
            
            if ($ScheduleRebootIfNeeded) {
                Set-RebootSchedule -Reason "$BrowserName update requires restart"
            }
            
            return $true
        }
        elseif ($process.ExitCode -eq 1641) {
            Write-RemediationLog -Message "$BrowserName installer initiated a restart (1641)" -Level "WARNING"
            return $true
        }
        else {
            # Parse MSI log for detailed error
            $errorDetail = Get-MSIErrorDetail -LogFile $logFile -ExitCode $process.ExitCode
            Write-RemediationLog -Message "$BrowserName $action failed with exit code: $($process.ExitCode). Detail: $errorDetail" -Level "ERROR" -Telemetry
            return $false
        }
    }
    catch {
        Write-RemediationLog -Message "Failed to $action $BrowserName : $_" -Level "ERROR" -Telemetry
        return $false
    }
}

function Test-BrowserInstalled {
    param(
        [string]$BrowserName
    )
    
    $testPaths = @{
        Chrome = @(
            "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
        )
        Edge = @(
            "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
            "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
        )
        Firefox = @(
            "${env:ProgramFiles}\Mozilla Firefox\firefox.exe",
            "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"
        )
    }
    
    foreach ($path in $testPaths[$BrowserName]) {
        if (Test-Path $path) {
            return $true
        }
    }
    
    return $false
}

function Get-MSIErrorDetail {
    param(
        [string]$LogFile,
        [int]$ExitCode
    )
    
    $commonErrors = @{
        1603 = "Fatal error during installation"
        1618 = "Another installation is in progress"
        1619 = "Installation package could not be opened"
        1620 = "Installation package could not be opened"
        1625 = "Installation prohibited by system policy"
        1633 = "Platform not supported"
        1638 = "Product version already installed"
    }
    
    if ($commonErrors.ContainsKey($ExitCode)) {
        $detail = $commonErrors[$ExitCode]
    }
    else {
        $detail = "Unknown error"
    }
    
    # Try to extract more detail from log
    if (Test-Path $LogFile) {
        try {
            $errorLines = Get-Content $LogFile -Tail 100 | Where-Object { $_ -match "error|failed" }
            if ($errorLines) {
                $detail += ". Log: " + ($errorLines[-1] -replace '\s+', ' ')
            }
        }
        catch {
            # Ignore log parsing errors
        }
    }
    
    return $detail
}

function Set-RebootSchedule {
    param(
        [string]$Reason = "Browser update requires restart",
        [int]$DelayHours = 4
    )
    
    try {
        $rebootTime = (Get-Date).AddHours($DelayHours)
        $taskName = "BrowserManager-ScheduledReboot"
        
        # Remove existing task if present
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        
        # Create scheduled task for reboot
        $action = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument "/r /t 300 /c `"$Reason`""
        $trigger = New-ScheduledTaskTrigger -Once -At $rebootTime
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal | Out-Null
        
        Write-RemediationLog -Message "Reboot scheduled for $rebootTime" -Telemetry
        
        # Notify user
        Show-UserNotification -Title "Restart Required" `
            -Message "Your computer will restart in $DelayHours hours to complete browser updates. Please save your work." `
            -Type "Warning" `
            -Actions @("Restart Now", "OK")
    }
    catch {
        Write-RemediationLog -Message "Failed to schedule reboot: $_" -Level "ERROR"
    }
}

function Update-Browser {
    param(
        [string]$BrowserName,
        [bool]$ForceClose = $false
    )
    
    Write-RemediationLog -Message "Processing update for $BrowserName..."
    
    # Check if browser is running
    $status = Test-BrowserRunning -BrowserName $BrowserName
    
    if ($status.IsRunning -and -not $ForceClose) {
        # If browser has no active windows, we can update silently
        if (-not $status.HasActiveWindows) {
            Write-RemediationLog -Message "$BrowserName is running but has no active windows, proceeding with silent update"
        }
        else {
            # Notify user to close browser
            Show-UserNotification -Title "Browser Update Required" `
                -Message "Please save your work and close $BrowserName to complete an important security update." `
                -Type "Warning" `
                -Actions @("Close $BrowserName", "Wait 30 Minutes") `
                -Tag "$BrowserName-Update"
            
            Write-RemediationLog -Message "Browser has active windows, waiting for user to close..." -Level "WARNING"
            
            # Wait for browser to close
            $closed = Wait-ForBrowserClose -BrowserName $BrowserName -MaxWaitMinutes $MaxWaitMinutes
            
            if (-not $closed) {
                # Final notification before deferring
                Show-UserNotification -Title "Update Postponed" `
                    -Message "$BrowserName update has been postponed. It will be retried on next schedule." `
                    -Type "Warning" `
                    -Tag "$BrowserName-Postponed"
                
                return @{
                    Success = $false
                    Reason = "Browser still running after wait period"
                    Action = "Deferred"
                }
            }
        }
    }
    
    # Download installer
    $installerPath = Get-BrowserInstaller -BrowserName $BrowserName -Url $downloadUrls[$BrowserName]
    
    if (-not $installerPath) {
        return @{
            Success = $false
            Reason = "Download failed"
            Action = "Failed"
        }
    }
    
    # Install/Update
    $success = Install-Browser -BrowserName $BrowserName -InstallerPath $installerPath -IsUpdate $true
    
    # Cleanup (keep cache for 1 hour)
    # Installer cleanup is handled by caching logic in Download-BrowserInstaller
    
    if ($success) {
        Show-UserNotification -Title "Browser Updated Successfully" `
            -Message "$BrowserName has been updated to the latest version. Thank you for your patience." `
            -Type "Success" `
            -Tag "$BrowserName-Success"
        
        return @{
            Success = $true
            Action = "Updated"
        }
    }
    else {
        return @{
            Success = $false
            Reason = "Installation failed"
            Action = "Failed"
        }
    }
}# Download installer
    $installerPath = Download-BrowserInstaller -BrowserName $BrowserName -Url $downloadUrls[$BrowserName]
    
    if (-not $installerPath) {
        return @{
            Success = $false
            Reason = "Download failed"
            Action = "Failed"
        }
    }
    
    # Install/Update
    $success = Install-Browser -BrowserName $BrowserName -InstallerPath $installerPath -IsUpdate $true
    
    # Cleanup
    if (Test-Path $installerPath) {
        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
    }
    
    if ($success) {
        if ($isRunning) {
            Show-UserNotification -Title "Browser Updated" `
                -Message "$BrowserName has been updated. Please restart the browser at your convenience." `
                -Type "Information"
        }
        
        return @{
            Success = $true
            Action = "Updated"
        }
    }
    else {
        return @{
            Success = $false
            Reason = "Installation failed"
            Action = "Failed"
        }
    }
}

function Install-MissingBrowser {
    param(
        [string]$BrowserName
    )
    
    Write-RemediationLog -Message "Installing $BrowserName..."
    
    # Download installer
    $installerPath = Get-BrowserInstaller -BrowserName $BrowserName -Url $downloadUrls[$BrowserName]
    
    if (-not $installerPath) {
        return @{
            Success = $false
            Reason = "Download failed"
            Action = "Failed"
        }
    }
    
    # Install
    $success = Install-Browser -BrowserName $BrowserName -InstallerPath $installerPath -IsUpdate $false
    
    # Cleanup (keep cache for 1 hour)
    # Installer cleanup is handled by caching logic in Get-BrowserInstaller
    
    if ($success) {
        Show-UserNotification -Title "Browser Installed" `
            -Message "$BrowserName has been installed successfully." `
            -Type "Success" `
            -Tag "$BrowserName-Installed"
        
        return @{
            Success = $true
            Action = "Installed"
        }
    }
    else {
        return @{
            Success = $false
            Reason = "Installation failed"
            Action = "Failed"
        }
    }
}

function Start-BrowserRemediation {
    # Load detection state
    $stateFile = "$env:ProgramData\BrowserManager\detection_state.json"
    
    if (-not (Test-Path $stateFile)) {
        Write-RemediationLog -Message "No detection state found, exiting" -Level "ERROR"
        Write-Host "No detection state found"
        exit 1
    }
    
    try {
        $detectionData = Get-Content $stateFile -Raw | ConvertFrom-Json
        
        # Handle both old and new state format
        if ($detectionData.PSObject.Properties.Name -contains "Metadata") {
            # New format with metadata
            $detectionState = $detectionData.Results
            $metadata = $detectionData.Metadata
            Write-RemediationLog -Message "Loaded detection state with metadata from $($metadata.Timestamp)"
        }
        else {
            # Old format - direct array
            $detectionState = $detectionData
            $metadata = @{}
        }
    }
    catch {
        Write-RemediationLog -Message "Failed to parse detection state: $_" -Level "ERROR"
        Write-Host "Failed to parse detection state"
        exit 1
    }
    
    # Filter and prioritize browsers
    $prioritizedState = $detectionState | Where-Object { $_.Action -in @("Install", "Update") } | Sort-Object -Property @(
        @{Expression = {
            switch ($_.Priority) {
                "Critical" { 0 }
                "High" { 1 }
                "Normal" { 2 }
                "Optional" { 3 }
                default { 4 }
            }
        }},
        @{Expression = {
            $browserName = $_.Browser -replace "Microsoft |Mozilla |Google ", ""
            $index = $PreferredBrowsers.IndexOf($browserName)
            if ($index -ge 0) { $index } else { 999 }
        }}
    )
    
    if ($prioritizedState.Count -eq 0) {
        Write-RemediationLog -Message "No browsers need remediation"
        Write-Host "No browsers need remediation"
        exit 0
    }
    
    Write-RemediationLog -Message "Found $($prioritizedState.Count) browsers needing remediation"
    
    $results = @()
    $deferredCount = 0
    
    foreach ($browser in $prioritizedState) {
        # Extract browser name without spaces
        $browserKey = $browser.Browser -replace "Microsoft |Mozilla |Google ", ""
        
        Write-RemediationLog -Message "Processing $($browser.Browser) - Priority: $($browser.Priority), Action: $($browser.Action)"
        
        if ($browser.Action -eq "Install") {
            $result = Install-MissingBrowser -BrowserName $browserKey
            $results += @{
                Browser = $browser.Browser
                Result = $result
                Priority = $browser.Priority
            }
            
            if (-not $result.Success) {
                Write-RemediationLog -Message "Failed to install $($browser.Browser)" -Level "ERROR"
            }
        }
        elseif ($browser.Action -eq "Update") {
            # For critical updates, be more aggressive
            $forceClose = $ForceUpdate -or ($browser.Priority -eq "Critical" -and $browser.Reason -like "*Below minimum*")
            
            $result = Update-Browser -BrowserName $browserKey -ForceClose:$forceClose
            $results += @{
                Browser = $browser.Browser
                Result = $result
                Priority = $browser.Priority
            }
            
            if ($result.Action -eq "Deferred") {
                $deferredCount++
            }
            elseif (-not $result.Success) {
                Write-RemediationLog -Message "Failed to update $($browser.Browser)" -Level "ERROR"
            }
        }
    }
    
    # Save remediation results
    $resultsFile = "$env:ProgramData\BrowserManager\remediation_results.json"
    $resultsData = @{
        Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        ComputerName = $env:COMPUTERNAME
        Results = $results
        Metadata = $metadata
    }
    $resultsData | ConvertTo-Json -Depth 4 | Set-Content -Path $resultsFile -Encoding UTF8
    
    # Save performance metrics
    $metricsFile = "$env:ProgramData\BrowserManager\performance_metrics.json"
    $script:performanceMetrics.EndTime = Get-Date
    $script:performanceMetrics.TotalDuration = ((Get-Date) - $script:performanceMetrics.StartTime).TotalSeconds
    $script:performanceMetrics | ConvertTo-Json -Depth 3 | Set-Content -Path $metricsFile -Encoding UTF8
    
    # Summary
    $installed = ($results | Where-Object { $_.Result.Action -eq "Installed" }).Count
    $updated = ($results | Where-Object { $_.Result.Action -eq "Updated" }).Count
    $deferred = ($results | Where-Object { $_.Result.Action -eq "Deferred" }).Count
    $failed = ($results | Where-Object { $_.Result.Action -eq "Failed" }).Count
    $critical = ($results | Where-Object { $_.Priority -eq "Critical" -and $_.Result.Action -eq "Failed" }).Count
    
    $summaryMessage = "Remediation complete: Installed=$installed, Updated=$updated, Deferred=$deferred, Failed=$failed"
    if ($critical -gt 0) {
        $summaryMessage += " (Critical failures: $critical)"
    }
    Write-RemediationLog -Message $summaryMessage -Telemetry
    
    # Send summary notification
    if ($installed -gt 0 -or $updated -gt 0) {
        $successMessage = "Successfully processed $($installed + $updated) browser(s)"
        if ($deferred -gt 0) {
            $successMessage += ". $deferred update(s) postponed (browser in use)"
        }
        
        Show-UserNotification -Title "Browser Management Complete" `
            -Message $successMessage `
            -Type "Success" `
            -Tag "BrowserManager-Summary"
    }
    
    # Handle critical failures specially
    if ($critical -gt 0) {
        $criticalBrowsers = $results | Where-Object { 
            $_.Priority -eq "Critical" -and $_.Result.Action -eq "Failed" 
        } | ForEach-Object { $_.Browser }
        
        Write-RemediationLog -Message "Critical browser failures: $($criticalBrowsers -join ', ')" -Level "ERROR" -Telemetry
        
        Show-UserNotification -Title "Critical Browser Update Failed" `
            -Message "Failed to update critical browsers: $($criticalBrowsers -join ', '). Please contact IT support." `
            -Type "Error" `
            -Tag "BrowserManager-Critical"
        
        Write-Host "Critical browser operations failed"
        exit 1
    }
    elseif ($failed -gt 0) {
        Write-Host "Some browser operations failed"
        exit 1
    }
    elseif ($deferred -gt 0) {
        Write-Host "Some updates deferred - browsers in use"
        exit 0  # Don't fail for deferred updates
    }
    else {
        Write-Host "All browser operations completed successfully"
        exit 0
    }
}

# Cleanup function for old logs and temp files
function Clear-OldFiles {
    param(
        [int]$DaysToKeep = 30
    )
    
    try {
        $cutoffDate = (Get-Date).AddDays(-$DaysToKeep)
        
        # Clean old logs
        $logDir = "$env:ProgramData\BrowserManager"
        if (Test-Path $logDir) {
            Get-ChildItem -Path $logDir -Filter "*.log" | 
                Where-Object { $_.LastWriteTime -lt $cutoffDate } | 
                Remove-Item -Force -ErrorAction SilentlyContinue
        }
        
        # Clean old installers
        $tempDir = "$env:TEMP\BrowserInstallers"
        if (Test-Path $tempDir) {
            Get-ChildItem -Path $tempDir -Filter "*.msi" | 
                Where-Object { $_.LastWriteTime -lt (Get-Date).AddHours(-2) } | 
                Remove-Item -Force -ErrorAction SilentlyContinue
        }
        
        Write-RemediationLog -Message "Cleaned up old files"
    }
    catch {
        Write-RemediationLog -Message "Cleanup failed: $_" -Level "WARNING"
    }
}

# Main execution
try {
    Write-RemediationLog -Message "=== Browser Remediation Started ===" -Telemetry
    Write-RemediationLog -Message "Parameters: ForceUpdate=$ForceUpdate, MaxWait=$MaxWaitMinutes min, ScheduleReboot=$ScheduleRebootIfNeeded"
    
    # Check for admin privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $isAdmin) {
        Write-RemediationLog -Message "Script requires administrator privileges" -Level "ERROR"
        Write-Host "Administrator privileges required"
        exit 1
    }
    
    # Perform cleanup of old files
    Clear-OldFiles -DaysToKeep 30
    
    # Run remediation
    Start-BrowserRemediation
}
catch {
    Write-RemediationLog -Message "Remediation script failed: $_" -Level "ERROR" -Telemetry
    Write-Host "Remediation failed: $_"
    
    # Send failure notification
    Show-UserNotification -Title "Browser Management Error" `
        -Message "An error occurred during browser management. IT has been notified." `
        -Type "Error" `
        -Tag "BrowserManager-Error"
    
    exit 1
}
finally {
    Write-RemediationLog -Message "=== Browser Remediation Completed ==="
    
    # Report performance metrics if verbose
    if ($VerbosePreference -eq 'Continue') {
        $metrics = Get-Content "$env:ProgramData\BrowserManager\performance_metrics.json" -Raw | ConvertFrom-Json
        Write-Verbose "Total execution time: $([math]::Round($metrics.TotalDuration)) seconds"
    }
}, "_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        Move-Item -Path $logPath -Destination $archivePath -Force
        
        # Keep only last 5 logs
        Get-ChildItem -Path $logDir -Filter "remediation_*.log" | 
            Sort-Object CreationTime -Descending | 
            Select-Object -Skip 5 | 
            Remove-Item -Force
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $Message" | Add-Content -Path $logPath -Encoding UTF8
    
    # Send telemetry for important events
    if ($Telemetry) {
        try {
            $eventLog = "Application"
            $source = "BrowserManager"
            
            if (-not [System.Diagnostics.EventLog]::SourceExists($source)) {
                [System.Diagnostics.EventLog]::CreateEventSource($source, $eventLog)
            }
            
            $eventType = switch ($Level) {
                "ERROR" { [System.Diagnostics.EventLogEntryType]::Error }
                "WARNING" { [System.Diagnostics.EventLogEntryType]::Warning }
                default { [System.Diagnostics.EventLogEntryType]::Information }
            }
            
            [System.Diagnostics.EventLog]::WriteEntry($source, $Message, $eventType, 1000)
        }
        catch {
            # Silently fail telemetry
        }
    }
}

function Show-UserNotification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$Type = "Information",
        [string[]]$Actions = @(),
        [string]$Tag = "BrowserManager",
        [string]$Group = "BrowserUpdates"
    )
    
    try {
        # Primary method: Windows Toast Notifications
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
        
        # Determine icon based on type
        $icon = switch ($Type) {
            "Warning" { "⚠️" }
            "Error" { "❌" }
            "Success" { "✅" }
            default { "ℹ️" }
        }
        
        # Build toast XML
        $template = @"
<toast scenario="reminder" activationType="protocol" launch="action=browserupdate">
    <visual>
        <binding template="ToastGeneric">
            <text>$icon $Title</text>
            <text>$Message</text>
            <image placement="appLogoOverride" src="https://raw.githubusercontent.com/microsoft/PowerToys/main/src/settings-ui/Settings.UI/Assets/FluentIcons/FluentIconsBrowsers.png" />
        </binding>
    </visual>
"@
        
        if ($Actions.Count -gt 0) {
            $template += "<actions>"
            foreach ($action in $Actions) {
                $template += "<action content='$action' arguments='action=$action' activationType='protocol'/>"
            }
            $template += "</actions>"
        }
        
        $template += @"
    <audio src="ms-winsoundevent:Notification.Default" />
</toast>
"@
        
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml($template)
        
        $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
        $toast.Tag = $Tag
        $toast.Group = $Group
        
        # Create notifier with app ID
        $appId = "Microsoft.Windows.Explorer"  # Use Explorer as it's always available
        $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId)
        $notifier.Show($toast)
        
        Write-RemediationLog -Message "Toast notification shown: $Title - $Message"
        return $true
    }
    catch {
        Write-RemediationLog -Message "Toast notification failed, trying fallback: $_" -Level "WARNING"
        
        # Fallback 1: Try WScript.Shell popup for user context
        try {
            $wshell = New-Object -ComObject Wscript.Shell
            $iconType = switch ($Type) {
                "Warning" { 48 }  # Exclamation
                "Error" { 16 }    # Critical
                "Success" { 64 }  # Information
                default { 64 }    # Information
            }
            $wshell.Popup($Message, 30, $Title, $iconType) | Out-Null
            Write-RemediationLog -Message "WScript popup shown"
            return $true
        }
        catch {
            Write-RemediationLog -Message "WScript popup failed: $_" -Level "WARNING"
        }
        
        # Fallback 2: Windows Forms balloon tip
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
            
            $notification = New-Object System.Windows.Forms.NotifyIcon
            $notification.Icon = [System.Drawing.SystemIcons]::Information
            $notification.BalloonTipIcon = $Type
            $notification.BalloonTipTitle = $Title
            $notification.BalloonTipText = $Message
            $notification.Visible = $true
            $notification.ShowBalloonTip(30000)
            
            Start-Sleep -Seconds 5
            $notification.Dispose()
            
            Write-RemediationLog -Message "Balloon tip shown"
            return $true
        }
        catch {
            Write-RemediationLog -Message "Balloon tip failed: $_" -Level "WARNING"
        }
        
        # Fallback 3: msg.exe for system context
        try {
            $sessions = @(query user 2>$null | Select-Object -Skip 1 | ForEach-Object {
                $parts = $_ -split '\s+'
                if ($parts[3] -eq "Active") { $parts[2] }
            })
            
            foreach ($sessionId in $sessions) {
                $simpleMessage = "$Title : $Message" -replace "`n", " "
                msg.exe $sessionId /TIME:30 $simpleMessage 2>$null
            }
            Write-RemediationLog -Message "msg.exe notification sent"
            return $true
        }
        catch {
            Write-RemediationLog -Message "All notification methods failed: $_" -Level "ERROR"
            return $false
        }
    }
}

function Test-BrowserRunning {
    param(
        [string]$BrowserName
    )
    
    $processes = $processNames[$BrowserName]
    if (-not $processes) { return $false }
    
    foreach ($processName in $processes) {
        $running = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($running) {
            Write-RemediationLog -Message "$BrowserName is running (Process: $processName)"
            return $true
        }
    }
    
    return $false
}

function Download-BrowserInstaller {
    param(
        [string]$BrowserName,
        [string]$Url
    )
    
    $tempDir = "$env:TEMP\BrowserInstallers"
    if (-not (Test-Path $tempDir)) {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    }
    
    $installerPath = Join-Path $tempDir "$BrowserName.msi"
    
    try {
        Write-RemediationLog -Message "Downloading $BrowserName installer..."
        
        # Use BITS for reliable download
        $bitsJob = Start-BitsTransfer -Source $Url -Destination $installerPath `
            -DisplayName "Download $BrowserName" -Priority Normal -Asynchronous
        
        # Wait for download with timeout
        $timeout = (Get-Date).AddMinutes(10)
        while ((Get-Date) -lt $timeout) {
            $job = Get-BitsTransfer -JobId $bitsJob.JobId -ErrorAction SilentlyContinue
            if (-not $job) { break }
            
            if ($job.JobState -eq "Transferred") {
                Complete-BitsTransfer -BitsJob $job
                Write-RemediationLog -Message "$BrowserName downloaded successfully"
                return $installerPath
            }
            elseif ($job.JobState -eq "Error") {
                Remove-BitsTransfer -BitsJob $job
                throw "BITS transfer failed"
            }
            
            Start-Sleep -Seconds 5
        }
        
        # Fallback to Invoke-WebRequest
        Write-RemediationLog -Message "BITS timeout, using direct download..." -Level "WARNING"
        Invoke-WebRequest -Uri $Url -OutFile $installerPath -UseBasicParsing
        
        if (Test-Path $installerPath) {
            Write-RemediationLog -Message "$BrowserName downloaded successfully (fallback method)"
            return $installerPath
        }
    }
    catch {
        Write-RemediationLog -Message "Failed to download $BrowserName : $_" -Level "ERROR"
        
        # Try alternative download for Firefox
        if ($BrowserName -eq "Firefox") {
            try {
                $altUrl = "https://download-installer.cdn.mozilla.net/pub/firefox/releases/latest/win64/en-US/Firefox%20Setup%20Latest.msi"
                Invoke-WebRequest -Uri $altUrl -OutFile $installerPath -UseBasicParsing
                if (Test-Path $installerPath) {
                    return $installerPath
                }
            }
            catch {
                Write-RemediationLog -Message "Alternative Firefox download also failed" -Level "ERROR"
            }
        }
    }
    
    return $null
}

function Install-Browser {
    param(
        [string]$BrowserName,
        [string]$InstallerPath,
        [bool]$IsUpdate = $false
    )
    
    if (-not (Test-Path $InstallerPath)) {
        Write-RemediationLog -Message "Installer not found: $InstallerPath" -Level "ERROR"
        return $false
    }
    
    $action = if ($IsUpdate) { "update" } else { "install" }
    Write-RemediationLog -Message "Starting $action for $BrowserName..."
    
    try {
        # MSI arguments for silent installation
        $arguments = @(
            "/i",
            "`"$InstallerPath`"",
            "/quiet",
            "/norestart",
            "/l*v",
            "`"$env:ProgramData\BrowserManager\$BrowserName_$action.log`""
        )
        
        # Add browser-specific parameters
        switch ($BrowserName) {
            "Chrome" {
                $arguments += "REBOOT=ReallySuppress"
                $arguments += "ALLUSERS=1"
            }
            "Edge" {
                $arguments += "REBOOT=ReallySuppress"
                $arguments += "DONOTCREATEDESKTOPSHORTCUT=TRUE"
            }
            "Firefox" {
                $arguments += "DESKTOP_SHORTCUT=false"
                $arguments += "INSTALL_MAINTENANCE_SERVICE=false"
            }
        }
        
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments `
            -Wait -PassThru -WindowStyle Hidden
        
        if ($process.ExitCode -eq 0) {
            Write-RemediationLog -Message "$BrowserName $action completed successfully"
            return $true
        }
        elseif ($process.ExitCode -eq 3010) {
            Write-RemediationLog -Message "$BrowserName $action completed, restart required" -Level "WARNING"
            return $true
        }
        else {
            Write-RemediationLog -Message "$BrowserName $action failed with exit code: $($process.ExitCode)" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-RemediationLog -Message "Failed to $action $BrowserName : $_" -Level "ERROR"
        return $false
    }
}

function Update-Browser {
    param(
        [string]$BrowserName,
        [bool]$ForceClose = $false
    )
    
    Write-RemediationLog -Message "Processing update for $BrowserName..."
    
    # Check if browser is running
    $isRunning = Test-BrowserRunning -BrowserName $BrowserName
    
    if ($isRunning -and -not $ForceClose) {
        # Notify user to close browser
        Show-UserNotification -Title "Browser Update Required" `
            -Message "Please close $BrowserName to complete the update. The update will be installed automatically once closed." `
            -Type "Warning"
        
        Write-RemediationLog -Message "Browser is running, waiting for user to close..." -Level "WARNING"
        
        # Schedule for next run
        return @{
            Success = $false
            Reason = "Browser running - user notified"
            Action = "Deferred"
        }
    }
    


function Install-MissingBrowser {
    param(
        [string]$BrowserName
    )
    
    Write-RemediationLog -Message "Installing $BrowserName..."
    
    # Download installer
    $installerPath = Get-BrowserInstaller -BrowserName $BrowserName -Url $downloadUrls[$BrowserName]
    
    if (-not $installerPath) {
        return @{
            Success = $false
            Reason = "Download failed"
            Action = "Failed"
        }
    }
    
    # Install
    $success = Install-Browser -BrowserName $BrowserName -InstallerPath $installerPath -IsUpdate $false
    
    # Cleanup
    if (Test-Path $installerPath) {
        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
    }
    
    if ($success) {
        Show-UserNotification -Title "Browser Installed" `
            -Message "$BrowserName has been installed successfully." `
            -Type "Information"
        
        return @{
            Success = $true
            Action = "Installed"
        }
    }
    else {
        return @{
            Success = $false
            Reason = "Installation failed"
            Action = "Failed"
        }
    }
}

function Start-BrowserRemediation {
    # Load detection state
    $stateFile = "$env:ProgramData\BrowserManager\detection_state.json"
    
    if (-not (Test-Path $stateFile)) {
        Write-RemediationLog -Message "No detection state found, exiting" -Level "ERROR"
        Write-Host "No detection state found"
        exit 1
    }
    
    try {
        $detectionState = Get-Content $stateFile -Raw | ConvertFrom-Json
    }
    catch {
        Write-RemediationLog -Message "Failed to parse detection state: $_" -Level "ERROR"
        Write-Host "Failed to parse detection state"
        exit 1
    }
    
    $results = @()
    $overallSuccess = $true
    
    foreach ($browser in $detectionState) {
        if ($browser.Action -eq "Install") {
            # Extract browser name without spaces
            $browserKey = $browser.Browser -replace "Microsoft |Mozilla |Google ", ""
            $result = Install-MissingBrowser -BrowserName $browserKey
            $results += @{
                Browser = $browser.Browser
                Result = $result
            }
            
            if (-not $result.Success) {
                Write-RemediationLog -Message "Failed to install $($browser.Browser)" -Level "ERROR"
            }
        }
        elseif ($browser.Action -eq "Update") {
            # Extract browser name without spaces
            $browserKey = $browser.Browser -replace "Microsoft |Mozilla |Google ", ""
            $result = Update-Browser -BrowserName $browserKey
            $results += @{
                Browser = $browser.Browser
                Result = $result
            }
            
            if (-not $result.Success -and $result.Action -ne "Deferred") {
                $overallSuccess = $false
            }
        }
    }
    
    # Save remediation results
    $resultsFile = "$env:ProgramData\BrowserManager\remediation_results.json"
    $results | ConvertTo-Json -Depth 3 | Set-Content -Path $resultsFile -Encoding UTF8
    
    # Summary
    $installed = ($results | Where-Object { $_.Result.Action -eq "Installed" }).Count
    $updated = ($results | Where-Object { $_.Result.Action -eq "Updated" }).Count
    $deferred = ($results | Where-Object { $_.Result.Action -eq "Deferred" }).Count
    $failed = ($results | Where-Object { $_.Result.Action -eq "Failed" }).Count
    
    Write-RemediationLog -Message "Remediation complete: Installed=$installed, Updated=$updated, Deferred=$deferred, Failed=$failed"
    
    if ($failed -gt 0) {
        Write-Host "Some browser operations failed"
        exit 1
    }
    elseif ($deferred -gt 0) {
        Write-Host "Some updates deferred - browsers in use"
        exit 0  # Don't fail for deferred updates
    }
    else {
        Write-Host "All browser operations completed successfully"
        exit 0
    }
}

# Main execution
try {
    Write-RemediationLog -Message "=== Browser Remediation Started ==="
    
    # Check for admin privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $isAdmin) {
        Write-RemediationLog -Message "Script requires administrator privileges" -Level "ERROR"
        Write-Host "Administrator privileges required"
        exit 1
    }
    
    Start-BrowserRemediation
}
catch {
    Write-RemediationLog -Message "Remediation script failed: $_" -Level "ERROR"
    Write-Host "Remediation failed: $_"
    exit 1
}
finally {
    Write-RemediationLog -Message "=== Browser Remediation Completed ==="
}