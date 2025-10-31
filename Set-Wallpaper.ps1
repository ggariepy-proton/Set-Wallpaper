# Set-Wallpaper.ps1
[CmdletBinding()]
param(
    [switch]$TestMode = $false
)

Function AddTextToImage {
    # Original code from http://www.ravichaganti.com/blog/?p=1012
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true)][String] $sourcePath,
        [Parameter(Mandatory=$true)][String] $destPath,
        [Parameter(Mandatory=$true)][String] $Title,
        [switch]$TestMode = $false
    )
 
    [Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null

    Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EntryType Information -EventId 2000 -Message "AddTextToImage: Get input image from $sourcePath"
    if ($TestMode) { Write-Host "AddTextToImage: Get input image from $sourcePath" }
    $srcImg = [System.Drawing.Image]::FromFile($sourcePath)

    Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EntryType Information -EventId 2001 -Message "AddTextToImage: Create labeled output at $destPath"
    if ($TestMode) { Write-Host "AddTextToImage: Create labeled output at $destPath" }
    $imgFile = new-object System.Drawing.Bitmap([int]($srcImg.width)),([int]($srcImg.height))
 
    $Image = [System.Drawing.Graphics]::FromImage($imgFile)
    $Image.SmoothingMode = "AntiAlias"
     
    $Rectangle = New-Object Drawing.Rectangle 0, 0, $srcImg.Width, $srcImg.Height
    $Image.DrawImage($srcImg, $Rectangle, 200, 200, $srcImg.Width, $srcImg.Height, ([Drawing.GraphicsUnit]::Pixel))

    Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EntryType Information -EventId 2002 -Message "AddTextToImage: Draw title: $Title"
    if ($TestMode) { Write-Host "AddTextToImage: Draw title: $Title" }
    $Font = new-object System.Drawing.Font("Verdana", 200)
    $Brush = New-Object Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 255, 255,255))
    $Image.DrawString($Title, $Font, $Brush, 10, 10)


    Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EntryType Information -EventId 2003 -Message "AddTextToImage: Save and close the file [$destPath]"
    if ($TestMode) { Write-Host "AddTextToImage: Save and close the file [$destPath]" }
    $imgFile.save($destPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    $imgFile.Dispose()
    $srcImg.Dispose()
}

$logFileExists = Get-EventLog -List | Where-Object { $_.LogDisplayName -eq "Set-Wallpaper" }

if (-not $logFileExists) {
    New-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper"
    Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EntryType Information -EventId 9999 -Message "Set-Wallpaper event log created."
}

# Import the FP.SetWallpaper module if not already loaded
$moduleExists = Get-Module -ListAvailable -Name FP.SetWallpaper

if ($null -eq $moduleExists) {
    Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EntryType Information -EventId 9998 -Message "Installing FP.SetWallpaper module."
    Install-Module FP.SetWallpaper -Force -AllowClobber -Scope CurrentUser -AcceptLicense
}
else {
    Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EntryType Information -EventId 1007 -Message "FP.SetWallpaper module already installed."
}

Import-Module FP.SetWallpaper

# Get the user's temp directory and set up logging
$path_to_file = "C:\Windows\Temp\SatelliteImagesTemp"
Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EntryType Information -EventId 1008 -Message "Script started. Using temp directory: $path_to_file"

# Create the directory if it doesn't exist
if (-not (Test-Path $path_to_file)) {
    Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EntryType Information -EventId 1009 -Message "Creating directory: $path_to_file"
    New-Item -ItemType Directory -Path $path_to_file | Out-Null
}

$urls = @(
    'https://cdn.star.nesdis.noaa.gov/GOES19/ABI/CONUS/GEOCOLOR/10000x6000.jpg', # North America
    'https://cdn.star.nesdis.noaa.gov/GOES19/ABI/FD/GEOCOLOR/10848x10848.jpg')   # Full Disk - Americas
    
$stayRunning = $true

while ($stayRunning -eq $true) {
    Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EventId 1000 -Message "Starting wallpaper update loop."
    $monitors = Get-Monitor
    $monitor_count = 0
    foreach ($url in $urls) {
        $regex = $url -match('\w+.jpg$')
        if ($regex) {
            $file_name = $Matches.0
            Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EntryType Information -EventId 1000 -Message "Getting new image and saving to [$file_name.tmp] in $path_to_file"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest $url -OutFile "$path_to_file\$file_name.tmp" -TimeoutSec 30 -SkipCertificateCheck -SkipHttpErrorCheck
            $ProgressPreference = 'Continue'
        }

        $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss" -AsUTC
        AddTextToImage -sourcePath "$path_to_file\$file_name.tmp" -destPath "$path_to_file\$file_name" -Title "$date UTC" -TestMode:$TestMode
        Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EventId 1001 -Message "Setting wallpaper [$path_to_file\$file_name] for monitor $monitor_count"
        Set-Wallpaper -InputObject $monitors[$monitor_count++] -LiteralPath "$path_to_file\$file_name" -Force 2>&1	
        Start-Sleep -Seconds 10
        Remove-Item "$path_to_file\$file_name"
        Remove-Item "$path_to_file\$file_name.tmp" -ErrorAction SilentlyContinue
        if ($monitor_count -ge $monitors.length) {
            Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EventId 1002 -Message "All monitors set, breaking out of loop."
            break
        }
    }
    Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EventId 1003 -Message "Sleeping for $(if ($TestMode) { '1 minute' } else { '15 minutes' }) before next update."
    if ($TestMode) {
        Start-Sleep -Seconds 60  # 1 minute in test mode
    } else {
        Start-Sleep -Seconds 900  # 15 minutes in normal mode
    }
    Write-EventLog -LogName "Set-Wallpaper" -Source "Set-Wallpaper" -EventId 1004 -Message "Waking up for next wallpaper update."
}