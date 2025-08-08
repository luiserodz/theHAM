# Browser Detection Script for Intune
# Purpose: Detects if browsers need installation or updates
# Returns: Exit 0 if compliant, Exit 1 if remediation needed
# Version: 2.0

[CmdletBinding()]
param(
    [Parameter()]
    [switch]$SkipOnlineCheck,

    [Parameter()]
    [int]$MaxVersionsBehind = 2,

    [Parameter()]
    [string[]]$RequiredBrowsers = @("Edge"),  # Browsers that MUST be installed

    [Parameter()]
    [string[]]$OptionalBrowsers = @("Chrome", "Firefox")  # Browsers to update if present
)

$ErrorActionPreference = 'Stop'

# Configuration
$browsers = @{
    Chrome = @{
        Name = "Google Chrome"
        RegistryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
        )
        VersionAPI = "https://versionhistory.googleapis.com/v1/chrome/platforms/win64/channels/stable/versions"
        ProcessName = "chrome"
        MinVersion = "120.0.0.0"
        ExecutablePaths = @(
            "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
            "${env:LocalAppData}\Google\Chrome\Application\chrome.exe"
        )
    }
    Edge = @{
        Name = "Microsoft Edge"
        RegistryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge",
            "HKLM:\SOFTWARE\Microsoft\Edge\Version",
            "HKCU:\SOFTWARE\Microsoft\Edge\Version"
        )
        VersionAPI = "https://edgeupdates.microsoft.com/api/products"
        ProcessName = "msedge"
        MinVersion = "120.0.0.0"
        ExecutablePaths = @(
            "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
            "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
        )
    }
    Firefox = @{
        Name = "Mozilla Firefox"
        RegistryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox*",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox*",
            "HKLM:\SOFTWARE\Mozilla\Mozilla Firefox"
        )
        VersionAPI = "https://product-details.mozilla.org/1.0/firefox_versions.json"
        ProcessName = "firefox"
        MinVersion = "120.0"
        ExecutablePaths = @(
            "${env:ProgramFiles}\Mozilla Firefox\firefox.exe",
            "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe",
            "${env:LocalAppData}\Mozilla Firefox\firefox.exe"
        )
    }
}

function Write-DetectionLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $logPath = "$env:ProgramData\BrowserManager\detection.log"
    $logDir = Split-Path $logPath -Parent

    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    if ((Test-Path $logPath) -and ((Get-Item $logPath).Length -gt 5MB)) {
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $archivePath = $logPath -replace '\.log$', "_" + $timestamp + ".log"
        Move-Item -Path $logPath -Destination $archivePath -Force

        Get-ChildItem -Path $logDir -Filter "detection_*.log" |
            Sort-Object CreationTime -Descending |
            Select-Object -Skip 5 |
            Remove-Item -Force
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $Message" | Add-Content -Path $logPath -Encoding UTF8
}

function Get-InstalledBrowserVersion {
    param(
        [string]$BrowserName,
        [array]$RegistryPaths,
        [array]$ExecutablePaths
    )

    foreach ($path in $RegistryPaths) {
        try {
            if ($path -like "*\\*") {
                $basePath = Split-Path $path -Parent
                $pattern = Split-Path $path -Leaf

                if (Test-Path $basePath) {
                    $items = Get-ChildItem $basePath -ErrorAction SilentlyContinue |
                        Where-Object { $_.PSChildName -like $pattern }

                    foreach ($item in $items) {
                        $regKey = Get-ItemProperty -Path $item.PSPath -ErrorAction SilentlyContinue
                        if ($regKey.DisplayVersion) {
                            Write-DetectionLog -Message "Found $BrowserName version in registry: $($regKey.DisplayVersion)"
                            return $regKey.DisplayVersion
                        }
                    }
                }
            }
            else {
                if (Test-Path $path) {
                    $regKey = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
                    if ($regKey.DisplayVersion) {
                        Write-DetectionLog -Message "Found $BrowserName version in registry: $($regKey.DisplayVersion)"
                        return $regKey.DisplayVersion
                    }
                    if ($regKey.Version) {
                        Write-DetectionLog -Message "Found $BrowserName version in registry: $($regKey.Version)"
                        return $regKey.Version
                    }
                }
            }
        }
        catch {
            continue
        }
    }

    foreach ($exePath in $ExecutablePaths) {
        if (Test-Path $exePath) {
            try {
                $fileVersion = (Get-Item $exePath).VersionInfo.ProductVersion
                if ($fileVersion) {
                    Write-DetectionLog -Message "Found $BrowserName version from executable: $fileVersion"
                    return $fileVersion
                }
            }
            catch {
                continue
            }
        }
    }

    return $null
}

function Get-LatestBrowserVersion {
    param(
        [string]$BrowserName,
        [string]$VersionAPI,
        [switch]$UseCache
    )

    $cacheFile = "$env:ProgramData\BrowserManager\version_cache.json"
    if ($UseCache -and (Test-Path $cacheFile)) {
        try {
            $cache = Get-Content $cacheFile -Raw | ConvertFrom-Json
            $cacheAge = (Get-Date) - [DateTime]$cache.Timestamp
            if ($cacheAge.TotalHours -lt 24 -and $cache.$BrowserName) {
                Write-DetectionLog -Message "Using cached version for $BrowserName : $($cache.$BrowserName)"
                return $cache.$BrowserName
            }
        } catch {
        }
    }

    if ($script:SkipOnlineCheck) {
        Write-DetectionLog -Message "Skipping online version check for $BrowserName"
        return $null
    }

    try {
        $testConnection = Test-NetConnection -ComputerName "www.microsoft.com" -Port 443 -WarningAction SilentlyContinue -InformationLevel Quiet
        if (-not $testConnection) {
            Write-DetectionLog -Message "Network connectivity test failed" -Level "WARNING"
            return $null
        }

        $response = Invoke-RestMethod -Uri $VersionAPI -Method Get -UseBasicParsing -TimeoutSec 10
        $version = $null

        switch ($BrowserName) {
            "Chrome" {
                if ($response -and $response[0].version) {
                    $version = $response[0].version
                }
            }
            "Edge" {
                $stable = $response | Where-Object { $_.Product -eq "Stable" }
                if ($stable -and $stable.Releases) {
                    $latestRelease = $stable.Releases |
                        Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" } |
                        Select-Object -First 1
                    if ($latestRelease) {
                        $version = $latestRelease.ProductVersion
                    }
                }
            }
            "Firefox" {
                if ($response.LATEST_FIREFOX_VERSION) {
                    $version = $response.LATEST_FIREFOX_VERSION
                }
            }
        }

        if ($version) {
            try {
                if (Test-Path $cacheFile) {
                    $cache = Get-Content $cacheFile -Raw | ConvertFrom-Json
                } else {
                    $cache = @{}
                }
                $cache | Add-Member -NotePropertyName $BrowserName -NotePropertyValue $version -Force
                $cache | Add-Member -NotePropertyName "Timestamp" -NotePropertyValue (Get-Date).ToString() -Force
                $cache | ConvertTo-Json | Set-Content $cacheFile -Encoding UTF8
            } catch {
            }
            return $version
        }
    }
    catch {
        Write-DetectionLog -Message "Failed to get latest version for $BrowserName : $_" -Level "ERROR"
        if ($UseCache -and (Test-Path $cacheFile)) {
            try {
                $cache = Get-Content $cacheFile -Raw | ConvertFrom-Json
                if ($cache.$BrowserName) {
                    Write-DetectionLog -Message "Using expired cache for $BrowserName : $($cache.$BrowserName)" -Level "WARNING"
                    return $cache.$BrowserName
                }
            } catch {
            }
        }
        return $null
    }

    return $null
}

function Compare-Versions {
    param(
        [string]$CurrentVersion,
        [string]$LatestVersion,
        [string]$MinimumVersion,
        [int]$MaxVersionsBehind = 2
    )

    try {
        $cleanCurrent = $CurrentVersion -replace '[^\d.]', ''
        $cleanLatest = if ($LatestVersion) { $LatestVersion -replace '[^\d.]', '' } else { $null }
        $cleanMinimum = $MinimumVersion -replace '[^\d.]', ''

        $currentParts = $cleanCurrent -split '\.' | Select-Object -First 4
        while ($currentParts.Count -lt 4) { $currentParts += "0" }
        $current = [Version]($currentParts -join '.')

        $minimumParts = $cleanMinimum -split '\.' | Select-Object -First 4
        while ($minimumParts.Count -lt 4) { $minimumParts += "0" }
        $minimum = [Version]($minimumParts -join '.')

        if ($current -lt $minimum) {
            return @{
                NeedsUpdate = $true
                Reason = "Below minimum version"
                Current = $CurrentVersion
                Latest = $LatestVersion
                Minimum = $MinimumVersion
            }
        }

        if ($cleanLatest) {
            $latestParts = $cleanLatest -split '\.' | Select-Object -First 4
            while ($latestParts.Count -lt 4) { $latestParts += "0" }
            $latest = [Version]($latestParts -join '.')

            if ($current -lt $latest) {
                $majorDiff = $latest.Major - $current.Major
                $minorDiff = if ($majorDiff -eq 0) { $latest.Minor - $current.Minor } else { 100 }

                if ($majorDiff -gt 0 -or $minorDiff -gt $MaxVersionsBehind) {
                    return @{
                        NeedsUpdate = $true
                        Reason = "Update available (${majorDiff} major, ${minorDiff} minor versions behind)"
                        Current = $CurrentVersion
                        Latest = $LatestVersion
                        VersionsBehind = @{
                            Major = $majorDiff
                            Minor = $minorDiff
                        }
                    }
                }
            }
        }

        return @{
            NeedsUpdate = $false
            Current = $CurrentVersion
            Latest = $LatestVersion
        }
    }
    catch {
        Write-DetectionLog -Message "Version comparison failed: $_" -Level "ERROR"
        return @{
            NeedsUpdate = $false
            Error = $_.Exception.Message
        }
    }
}

function Test-BrowserCompliance {
    $results = @()
    $needsRemediation = $false
    $criticalIssues = @()

    Write-DetectionLog -Message "Starting compliance check with MaxVersionsBehind=$MaxVersionsBehind"
    Write-DetectionLog -Message "Required browsers: $($RequiredBrowsers -join ', ')"
    Write-DetectionLog -Message "Optional browsers: $($OptionalBrowsers -join ', ')"

    foreach ($browserKey in $browsers.Keys) {
        $browser = $browsers[$browserKey]
        $isRequired = $RequiredBrowsers -contains $browserKey
        $isOptional = $OptionalBrowsers -contains $browserKey

        if (-not $isRequired -and -not $isOptional) {
            Write-DetectionLog -Message "Skipping $($browser.Name) - not in required or optional list"
            continue
        }

        Write-DetectionLog -Message "Checking $($browser.Name) (Required: $isRequired)..."

        $installedVersion = Get-InstalledBrowserVersion -BrowserName $browserKey -RegistryPaths $browser.RegistryPaths -ExecutablePaths $browser.ExecutablePaths

        if (-not $installedVersion) {
            if ($isRequired) {
                Write-DetectionLog -Message "$($browser.Name) is REQUIRED but not installed" -Level "ERROR"
                $results += @{
                    Browser = $browser.Name
                    Status = "Not Installed"
                    Action = "Install"
                    Priority = "Critical"
                }
                $needsRemediation = $true
                $criticalIssues += "$($browser.Name) not installed (required)"
            } else {
                Write-DetectionLog -Message "$($browser.Name) is not installed (optional)"
                $results += @{
                    Browser = $browser.Name
                    Status = "Not Installed"
                    Action = "Skip"
                    Priority = "Optional"
                }
            }
            continue
        }

        Write-DetectionLog -Message "$($browser.Name) installed version: $installedVersion"

        $latestVersion = Get-LatestBrowserVersion -BrowserName $browserKey -VersionAPI $browser.VersionAPI -UseCache

        if ($latestVersion) {
            Write-DetectionLog -Message "$($browser.Name) latest version: $latestVersion"
            $comparison = Compare-Versions -CurrentVersion $installedVersion -LatestVersion $latestVersion -MinimumVersion $browser.MinVersion -MaxVersionsBehind $MaxVersionsBehind
            if ($comparison.NeedsUpdate) {
                $priority = if ($comparison.Reason -like "*Below minimum*") { "Critical" } elseif ($isRequired) { "High" } else { "Normal" }
                Write-DetectionLog -Message "$($browser.Name) needs update: $($comparison.Reason)" -Level "WARNING"
                $results += @{
                    Browser = $browser.Name
                    Status = "Outdated"
                    Current = $comparison.Current
                    Latest = $comparison.Latest
                    Action = "Update"
                    Reason = $comparison.Reason
                    Priority = $priority
                }
                $needsRemediation = $true
                if ($priority -eq "Critical") { $criticalIssues += "$($browser.Name) $($comparison.Reason)" }
            } else {
                Write-DetectionLog -Message "$($browser.Name) is up to date"
                $results += @{
                    Browser = $browser.Name
                    Status = "Compliant"
                    Version = $installedVersion
                    Latest = $latestVersion
                }
            }
        } else {
            Write-DetectionLog -Message "Could not determine latest version for $($browser.Name)" -Level "WARNING"
            $comparison = Compare-Versions -CurrentVersion $installedVersion -LatestVersion $null -MinimumVersion $browser.MinVersion -MaxVersionsBehind $MaxVersionsBehind
            if ($comparison.NeedsUpdate) {
                $results += @{
                    Browser = $browser.Name
                    Status = "Below Minimum"
                    Current = $installedVersion
                    Minimum = $browser.MinVersion
                    Action = "Update"
                    Priority = "Critical"
                }
                $needsRemediation = $true
                $criticalIssues += "$($browser.Name) below minimum version"
            } else {
                $results += @{
                    Browser = $browser.Name
                    Status = "Version Check Failed"
                    Version = $installedVersion
                    Note = "Unable to verify latest version"
                }
            }
        }
    }

    $metadata = @{
        Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        ComputerName = $env:COMPUTERNAME
        UserName = $env:USERNAME
        RequiredBrowsers = $RequiredBrowsers
        OptionalBrowsers = $OptionalBrowsers
        CriticalIssues = $criticalIssues
        NeedsRemediation = $needsRemediation
    }

    $stateFile = "$env:ProgramData\BrowserManager\detection_state.json"
    $state = @{
        Metadata = $metadata
        Results = $results
    }
    $state | ConvertTo-Json -Depth 4 | Set-Content -Path $stateFile -Encoding UTF8

    $summary = "Detection complete. Remediation needed: $needsRemediation"
    if ($criticalIssues.Count -gt 0) {
        $summary += ". Critical issues: $($criticalIssues -join '; ')"
    }
    Write-DetectionLog -Message $summary

    if ($needsRemediation) {
        Write-Host "Browsers need installation or updates"
        if ($criticalIssues.Count -gt 0) {
            Write-Host "Critical: $($criticalIssues -join '; ')"
        }
        exit 1
    } else {
        Write-Host "All browsers are compliant"
        exit 0
    }
}

try {
    Write-DetectionLog -Message "=== Browser Detection Started ==="
    Write-DetectionLog -Message "Script Version: 2.0"
    Write-DetectionLog -Message "Parameters: SkipOnlineCheck=$SkipOnlineCheck, MaxVersionsBehind=$MaxVersionsBehind"

    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($currentUser)
    $isAdmin = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-DetectionLog -Message "Running without administrator privileges - some checks may be limited" -Level "WARNING"
    }

    $logDir = "$env:ProgramData\BrowserManager"
    if (Test-Path $logDir) {
        $oldLogs = Get-ChildItem -Path $logDir -Filter "*.log" |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) }
        if ($oldLogs.Count -gt 0) {
            $oldLogs | Remove-Item -Force -ErrorAction SilentlyContinue
            Write-DetectionLog -Message "Removed $($oldLogs.Count) old log files"
        }
    }

    if (-not $SkipOnlineCheck) {
        try {
            $networkTest = Test-NetConnection -ComputerName "www.microsoft.com" -Port 443 -WarningAction SilentlyContinue -InformationLevel Quiet
            if (-not $networkTest) {
                Write-DetectionLog -Message "Network connectivity test failed - will skip online version checks" -Level "WARNING"
                $script:SkipOnlineCheck = $true
            }
        } catch {
            Write-DetectionLog -Message "Network test failed: $_ - will skip online version checks" -Level "WARNING"
            $script:SkipOnlineCheck = $true
        }
    }

    Test-BrowserCompliance
}
catch {
    Write-DetectionLog -Message "Detection script failed: $_" -Level "ERROR"
    Write-DetectionLog -Message "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
    Write-Host "Detection failed: $_"

    $errorState = @{
        Error = $_.Exception.Message
        StackTrace = $_.ScriptStackTrace
        Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }
    $errorFile = "$env:ProgramData\BrowserManager\detection_error.json"
    $errorState | ConvertTo-Json | Set-Content -Path $errorFile -Encoding UTF8
    exit 1
}
finally {
    Write-DetectionLog -Message "=== Browser Detection Completed ==="
}
