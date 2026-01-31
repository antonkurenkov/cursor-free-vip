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

# Function to extract ZIP file with overwrite support
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
        } else {
            Write-Styled "Destination folder already exists, cleaning old files..." -Color $Theme.Warning -Prefix "Cleanup"
        }
        
        # Extract ZIP with overwrite support
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        
        # Open ZIP archive
        $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipFile)
        
        try {
            foreach ($entry in $zip.Entries) {
                $entryPath = Join-Path $Destination $entry.FullName
                $entryDir = Split-Path $entryPath -Parent
                
                # Create directory if it doesn't exist
                if (-not (Test-Path $entryDir)) {
                    New-Item -ItemType Directory -Path $entryDir -Force | Out-Null
                }
                
                # Skip directories
                if (-not [string]::IsNullOrEmpty($entry.Name)) {
                    # Extract file with overwrite
                    try {
                        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $entryPath, $true)
                    }
                    catch {
                        Write-Styled "Warning: Could not extract $($entry.Name): $($_.Exception.Message)" -Color $Theme.Warning -Prefix "Warning"
                    }
                }
            }
        }
        finally {
            $zip.Dispose()
        }
        
        Write-Styled "Extraction completed" -Color $Theme.Success -Prefix "Extract"
        return $true
    }
    catch {
        Write-Styled "Failed to extract ZIP: $($_.Exception.Message)" -Color $Theme.Error -Prefix "Error"
        return $false
    }
}

# Alternative function using Shell.Application (more compatible)
function Expand-ZipFileCompatible {
    param(
        [string]$ZipFile,
        [string]$Destination
    )
    
    try {
        Write-Styled "Extracting ZIP file (compatible method)..." -Color $Theme.Primary -Prefix "Extract"
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path $Destination)) {
            New-Item -ItemType Directory -Path $Destination -Force | Out-Null
        }
        
        # Use Shell.Application for extraction (supports overwrite)
        $shell = New-Object -ComObject Shell.Application
        $zip = $shell.NameSpace($ZipFile)
        
        # Copy all items with overwrite flag (16 = 4 (NoProgressDialog) + 16 (YesToAll))
        $flags = 20  # 4 + 16
        
        foreach ($item in $zip.Items()) {
            $shell.NameSpace($Destination).CopyHere($item, $flags)
            # Small delay to avoid issues
            Start-Sleep -Milliseconds 100
        }
        
        # Release COM object
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        
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
        
        # Clean old extraction if exists and user wants to
        if (Test-Path $extractPath) {
            Write-Styled "Previous extraction found. Delete old files? (Y/N)" -Color $Theme.Warning -Prefix "Cleanup"
            $response = Read-Host "Press Y to delete or N to keep"
            if ($response -eq 'Y' -or $response -eq 'y') {
                try {
                    Remove-Item -Path $extractPath -Recurse -Force -ErrorAction Stop
                    Write-Styled "Old files deleted successfully" -Color $Theme.Success -Prefix "Cleanup"
                }
                catch {
                    Write-Styled "Could not delete old files: $($_.Exception.Message)" -Color $Theme.Error -Prefix "Warning"
                }
            }
        }
        
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
                $progress = [math]::Round(($Global:downloadedBytes / $Global.totalBytes) * 100, 1)
                
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
        
        # Try first method, if fails try second method
        $success = $false
        
        try {
            # First try with .NET method
            if (Expand-ZipFile -ZipFile $downloadPath -Destination $extractPath) {
                $success = $true
            }
        }
        catch {
            Write-Styled "First extraction method failed, trying alternative..." -Color $Theme.Warning -Prefix "Warning"
            
            # Try with Shell.Application method
            if (Expand-ZipFileCompatible -ZipFile $downloadPath -Destination $extractPath) {
                $success = $true
            }
        }
        
        if ($success) {
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
                # Also look for .bat files
                $batFiles = Get-ChildItem -Path $extractPath -Filter "*.bat" -Recurse -File | Select-Object -First 1
                if ($batFiles) {
                    Write-Styled "Found batch file: Launcher.bat" -Color $Theme.Success -Prefix "Found"
                    Write-Styled "Opening folder with instructions..." -Color $Theme.Info -Prefix "Info"
                    
                    # Show message about Launcher.bat
                    Write-Host "`n" -NoNewline
                    Write-Host "="*50 -ForegroundColor Cyan
                    Write-Host "INSTRUCTIONS:" -ForegroundColor Yellow
                    Write-Host "- Open the folder: $extractPath" -ForegroundColor White
                    Write-Host "- Run 'Launcher.bat' as Administrator" -ForegroundColor White
                    Write-Host "- Follow the on-screen instructions" -ForegroundColor White
                    Write-Host "="*50 -ForegroundColor Cyan
                    Write-Host "`n"
                    
                    # Open folder
                    Start-Process $extractPath
                } else {
                    Write-Styled "No executable or batch file found. Opening extraction folder..." -Color $Theme.Warning -Prefix "Warning"
                    Start-Process $extractPath
                }
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
