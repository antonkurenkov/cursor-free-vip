# set color theme
$Theme = @{
    Primary   = 'Cyan'
    Success   = 'Green'
    Warning   = 'Yellow'
    Error     = 'Red'
    Info      = 'White'
}

# ASCII Logo
$Logo = @"
   ██████╗██╗   ██╗██████╗ ███████╗ ██████╗ ██████╗      ██████╗ ██████╗  ██████╗   
  ██╔════╝██║   ██║██╔══██╗██╔════╝██╔═══██╗██╔══██╗     ██╔══██╗██╔══██╗██╔═══██╗  
  ██║     ██║   ██║██████╔╝███████╗██║   ██║██████╔╝     ██████╔╝██████╔╝██║   ██║  
  ██║     ██║   ██║██╔══██╗╚════██║██║   ██║██╔══██╗     ██╔═══╝ ██╔══██╗██║   ██║  
  ╚██████╗╚██████╔╝██║  ██║███████║╚██████╔╝██║  ██║     ██║     ██║  ██║╚██████╔╝  
   ╚═════╝ ╚═════╝ ╚═╝  ╚═╝╚══════╝ ╚═════╝ ╚═╝  ╚═╝     ╚═╝     ╚═╝  ╚═╝ ╚═════╝  
"@

# Beautiful Output Function
function Write-Styled {
    param (
        [string]$Message,
        [string]$Color = $Theme.Info,
        [string]$Prefix = "",
        [switch]$NoNewline
    )
    $symbol = switch ($Color) {
        $Theme.Success { "[OK]" }
        $Theme.Error   { "[X]" }
        $Theme.Warning { "[!]" }
        default        { "[*]" }
    }
    
    $output = if ($Prefix) { "$symbol $Prefix :: $Message" } else { "$symbol $Message" }
    if ($NoNewline) {
        Write-Host $output -ForegroundColor $Color -NoNewline
    } else {
        Write-Host $output -ForegroundColor $Color
    }
}

# Get version number function
function Get-LatestVersion {
    try {
        $latestRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/SHANMUGAM070106/cursor-free-vip/releases/latest"
        return @{
            Version = $latestRelease.tag_name.TrimStart('v')
            Assets = $latestRelease.assets
        }
    } catch {
        Write-Styled $_.Exception.Message -Color $Theme.Error -Prefix "Error"
        throw "Cannot get latest version"
    }
}

# Show Logo
Write-Host $Logo -ForegroundColor $Theme.Primary
$releaseInfo = Get-LatestVersion
$version = $releaseInfo.Version
Write-Host "Version $version" -ForegroundColor $Theme.Info
Write-Host "Created by YeongPin`n" -ForegroundColor $Theme.Info

# Set TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Function to extract ZIP file
function Expand-ZipFile {
    param(
        [string]$ZipFile,
        [string]$Destination
    )
    
    try {
        Write-Styled "Extracting ZIP file..." -Color $Theme.Primary -Prefix "Extract"
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path $Destination)) {
            New-Item -ItemType Directory -Path $Destination -Force | Out-Null
        }
        
        # Extract ZIP
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipFile, $Destination)
        
        Write-Styled "Extraction completed" -Color $Theme.Success -Prefix "Extract"
        return $true
    }
    catch {
        Write-Styled "Failed to extract ZIP: $($_.Exception.Message)" -Color $Theme.Error -Prefix "Error"
        return $false
    }
}

# Main installation function
function Install-CursorFreeVIP {
    Write-Styled "Start downloading Cursor Free VIP" -Color $Theme.Primary -Prefix "Download"
    
    try {
        # Get latest version
        Write-Styled "Checking latest version..." -Color $Theme.Primary -Prefix "Update"
        $releaseInfo = Get-LatestVersion
        $version = $releaseInfo.Version
        Write-Styled "Found latest version: $version" -Color $Theme.Success -Prefix "Version"
        
        # Find corresponding resources - теперь ищем ZIP файл
        $asset = $releaseInfo.Assets | Where-Object { $_.name -eq "CursorFreeVIP_${version}_windows.zip" }
        if (!$asset) {
            Write-Styled "File not found: CursorFreeVIP_${version}_windows.zip" -Color $Theme.Error -Prefix "Error"
            Write-Styled "Available files:" -Color $Theme.Warning -Prefix "Info"
            $releaseInfo.Assets | ForEach-Object {
                Write-Styled "- $($_.name)" -Color $Theme.Info
            }
            throw "Cannot find target file"
        }
        
        # Check if Downloads folder already exists for the corresponding version
        $DownloadsPath = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
        $downloadPath = Join-Path $DownloadsPath "CursorFreeVIP_${version}_windows.zip"
        $extractPath = Join-Path $DownloadsPath "CursorFreeVIP_${version}"
        
        if (Test-Path $downloadPath) {
            Write-Styled "Found existing ZIP file" -Color $Theme.Success -Prefix "Found"
            Write-Styled "Location: $downloadPath" -Color $Theme.Info -Prefix "Location"
        } else {
            Write-Styled "No existing ZIP file found, starting download..." -Color $Theme.Primary -Prefix "Download"
            
            # Create WebClient and add progress event
            $webClient = New-Object System.Net.WebClient
            $webClient.Headers.Add("User-Agent", "PowerShell Script")

            # Define progress variables
            $Global:downloadedBytes = 0
            $Global:totalBytes = 0
            $Global:lastProgress = 0
            $Global:lastBytes = 0
            $Global:lastTime = Get-Date

            # Download progress event
            $eventId = [guid]::NewGuid()
            Register-ObjectEvent -InputObject $webClient -EventName DownloadProgressChanged -Action {
                $Global:downloadedBytes = $EventArgs.BytesReceived
                $Global:totalBytes = $EventArgs.TotalBytesToReceive
                $progress = [math]::Round(($Global:downloadedBytes / $Global:totalBytes) * 100, 1)
                
                # Only update display when progress changes by more than 1%
                if ($progress -gt $Global:lastProgress + 1) {
                    $Global:lastProgress = $progress
                    $downloadedMB = [math]::Round($Global:downloadedBytes / 1MB, 2)
                    $totalMB = [math]::Round($Global.totalBytes / 1MB, 2)
                    
                    # Calculate download speed
                    $currentTime = Get-Date
                    $timeSpan = ($currentTime - $Global:lastTime).TotalSeconds
                    if ($timeSpan -gt 0) {
                        $bytesChange = $Global:downloadedBytes - $Global:lastBytes
                        $speed = $bytesChange / $timeSpan
                        
                        # Choose appropriate unit based on speed
                        $speedDisplay = if ($speed -gt 1MB) {
                            "$([math]::Round($speed / 1MB, 2)) MB/s"
                        } elseif ($speed -gt 1KB) {
                            "$([math]::Round($speed / 1KB, 2)) KB/s"
                        } else {
                            "$([math]::Round($speed, 2)) B/s"
                        }
                        
                        Write-Host "`rDownloading: $downloadedMB MB / $totalMB MB ($progress%) - $speedDisplay" -NoNewline -ForegroundColor Cyan
                        
                        # Update last data
                        $Global:lastBytes = $Global:downloadedBytes
                        $Global:lastTime = $currentTime
                    }
                }
            } | Out-Null

            # Download completed event
            Register-ObjectEvent -InputObject $webClient -EventName DownloadFileCompleted -Action {
                Write-Host "`r" -NoNewline
                Write-Styled "Download completed!" -Color $Theme.Success -Prefix "Complete"
                Unregister-Event -SourceIdentifier $eventId
            } | Out-Null

            # Start download
            $webClient.DownloadFileAsync([Uri]$asset.browser_download_url, $downloadPath)

            # Wait for download to complete
            while ($webClient.IsBusy) {
                Start-Sleep -Milliseconds 100
            }
            
            Write-Styled "File location: $downloadPath" -Color $Theme.Info -Prefix "Location"
        }
        
        # Extract ZIP file
        Write-Styled "Extracting files..." -Color $Theme.Primary -Prefix "Extract"
        if (Expand-ZipFile -ZipFile $downloadPath -Destination $extractPath) {
            Write-Styled "Files extracted to: $extractPath" -Color $Theme.Success -Prefix "Extract"
            
            # Look for executable in extracted files
            $exeFiles = Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse -File | Select-Object -First 1
            if ($exeFiles) {
                $exePath = $exeFiles.FullName
                Write-Styled "Found executable: $(Split-Path $exePath -Leaf)" -Color $Theme.Success -Prefix "Found"
                Write-Styled "Starting program..." -Color $Theme.Primary -Prefix "Launch"
                
                # Check if running with administrator privileges
                $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                
                if (-not $isAdmin) {
                    Write-Styled "Requesting administrator privileges..." -Color $Theme.Warning -Prefix "Admin"
                    
                    # Create new process with administrator privileges
                    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $startInfo.FileName = $exePath
                    $startInfo.UseShellExecute = $true
                    $startInfo.Verb = "runas"
                    
                    try {
                        [System.Diagnostics.Process]::Start($startInfo)
                        Write-Styled "Program started with admin privileges" -Color $Theme.Success -Prefix "Launch"
                        return
                    }
                    catch {
                        Write-Styled "Failed to start with admin privileges. Starting normally..." -Color $Theme.Warning -Prefix "Warning"
                        Start-Process $exePath
                        return
                    }
                }
                
                # If already running with administrator privileges, start directly
                Start-Process $exePath
            } else {
                Write-Styled "No executable found in extracted files. Opening extraction folder..." -Color $Theme.Warning -Prefix "Warning"
                Start-Process $extractPath
            }
        }
    }
    catch {
        Write-Styled $_.Exception.Message -Color $Theme.Error -Prefix "Error"
        throw
    }
}

# Execute installation
try {
    Install-CursorFreeVIP
}
catch {
    Write-Styled "Download failed" -Color $Theme.Error -Prefix "Error"
    Write-Styled $_.Exception.Message -Color $Theme.Error
}
finally {
    Write-Host "`nPress any key to exit..." -ForegroundColor $Theme.Info
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
