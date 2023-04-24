# Bing Wallpapers
# Fetch the Bing wallpaper image of the day
# <https://github.com/timothymctim/Bing-wallpapers>
#
# Copyright (c) 2015 Tim van de Kamp
# License: MIT license

Param(
    # Get the Bing image of this country
    [ValidateSet('auto', 'ar-XA', 'bg-BG', 'cs-CZ', 'da-DK', 'de-AT',
    'de-CH', 'de-DE', 'el-GR', 'en-AU', 'en-CA', 'en-GB', 'en-ID',
    'en-IE', 'en-IN', 'en-MY', 'en-NZ', 'en-PH', 'en-SG', 'en-US',
    'en-XA', 'en-ZA', 'es-AR', 'es-CL', 'es-ES', 'es-MX', 'es-US',
    'es-XL', 'et-EE', 'fi-FI', 'fr-BE', 'fr-CA', 'fr-CH', 'fr-FR',
    'he-IL', 'hr-HR', 'hu-HU', 'it-IT', 'ja-JP', 'ko-KR', 'lt-LT',
    'lv-LV', 'nb-NO', 'nl-BE', 'nl-NL', 'pl-PL', 'pt-BR', 'pt-PT',
    'ro-RO', 'ru-RU', 'sk-SK', 'sl-SL', 'sv-SE', 'th-TH', 'tr-TR',
    'uk-UA', 'zh-CN', 'zh-HK', 'zh-TW')][string]$locale = 'auto',

    # Download the latest $files wallpapers
    [int]$files = 3,

    # Resolution of the image to download
    [ValidateSet('auto', '800x600', '1024x768', '1280x720', '1280x768',
    '1366x768', '1920x1080', '1920x1200', '720x1280', '768x1024',
    '768x1280', '768x1366', '1080x1920')][string]$resolution = 'auto',

    # Destination folder to download the wallpapers to
    [string]$downloadFolder = "$([Environment]::GetFolderPath("MyPictures"))\Wallpapers"
)

Function Set-WallPaper {
 
    <#
     
        .SYNOPSIS
        Applies a specified wallpaper to the current user's desktop
        
        .PARAMETER Image
        Provide the exact path to the image
     
        .PARAMETER Style
        Provide wallpaper style (Example: Fill, Fit, Stretch, Tile, Center, or Span)
      
        .EXAMPLE
        Set-WallPaper -Image "C:\Wallpaper\Default.jpg"
        Set-WallPaper -Image "C:\Wallpaper\Background.jpg" -Style Fit
      
    #>
     
    param (
        [parameter(Mandatory = $True)]
        # Provide path to image
        [string]$Image,
        # Provide wallpaper style that you would like applied
        [parameter(Mandatory = $False)]
        [ValidateSet('Fill', 'Fit', 'Stretch', 'Tile', 'Center', 'Span')]
        [string]$Style
    )
     
    $WallpaperStyle = Switch ($Style) {
      
        "Fill" { "10" }
        "Fit" { "6" }
        "Stretch" { "2" }
        "Tile" { "0" }
        "Center" { "0" }
        "Span" { "22" }
      
    }
     
    If ($Style -eq "Tile") {
     
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 1 -Force
     
    }
    Else {
     
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 0 -Force
     
    }
     
    Add-Type -TypeDefinition @" 
    using System; 
    using System.Runtime.InteropServices;
      
    public class Params
    { 
        [DllImport("User32.dll",CharSet=CharSet.Unicode)] 
        public static extern int SystemParametersInfo (Int32 uAction, 
                                                       Int32 uParam, 
                                                       String lpvParam, 
                                                       Int32 fuWinIni);
    }
"@ 
      
    $SPI_SETDESKWALLPAPER = 0x0014
    $UpdateIniFile = 0x01
    $SendChangeEvent = 0x02
      
    $fWinIni = $UpdateIniFile -bor $SendChangeEvent
      
    $ret = [Params]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $Image, $fWinIni)
}
# Max item count: the number of images we'll query for
[int]$maxItemCount = [System.Math]::max(1, [System.Math]::max($files, 8))
# URI to fetch the image locations from
if ($locale -eq 'auto') {
    $market = ""
} else {
    $market = "&mkt=$locale"
}
[string]$hostname = "https://www.bing.com"
[string]$uri = "$hostname/HPImageArchive.aspx?format=xml&idx=0&n=$maxItemCount$market"

# Get the appropiate screen resolution
if ($resolution -eq 'auto') {
    Add-Type -AssemblyName System.Windows.Forms
    $primaryScreen = [System.Windows.Forms.Screen]::AllScreens | Where-Object {$_.Primary -eq 'True'}
    if ($primaryScreen.Bounds.Width -le 1024) {
        $resolution = '1024x768'
    } elseif ($primaryScreen.Bounds.Width -le 1280) {
        $resolution = '1280x720'
    } elseif ($primaryScreen.Bounds.Width -le 1366) {
        $resolution = '1366x768'
    } elseif ($primaryScreen.Bounds.Height -le 1080) {
        $resolution = '1920x1080'
    } else {
        $resolution = '1920x1200'
    }
}

# Check if download folder exists and otherwise create it
if (!(Test-Path $downloadFolder)) {
    New-Item -ItemType Directory $downloadFolder
}

$request = Invoke-WebRequest -Uri $uri -UseBasicParsing
[xml]$content = $request.Content

$items = New-Object System.Collections.ArrayList
foreach ($xmlImage in $content.images.image) {
    [datetime]$imageDate = [datetime]::ParseExact($xmlImage.startdate, 'yyyyMMdd', $null)
    [string]$imageUrl = "$hostname$($xmlImage.urlBase)_$resolution.jpg"

    # Add item to our array list
    $item = New-Object System.Object
    $item | Add-Member -Type NoteProperty -Name date -Value $imageDate
    $item | Add-Member -Type NoteProperty -Name url -Value $imageUrl
    $null = $items.Add($item)
}

# Keep only the most recent $files items to download
if (!($files -eq 0) -and ($items.Count -gt $files)) {
    # We have too many matches, keep only the most recent
    $items = $items|Sort date
    while ($items.Count -gt $files) {
        # Pop the oldest item of the array
        $null, $items = $items
    }
}

Write-Host "Downloading images..."
$client = New-Object System.Net.WebClient
foreach ($item in $items) {
    $baseName = $item.date.ToString("yyyy-MM-dd")
    $destination = "$downloadFolder\$baseName.jpg"
    $url = $item.url

    # Download the enclosure if we haven't done so already
    if (!(Test-Path $destination)) {
        Write-Debug "Downloading image to $destination"
        $client.DownloadFile($url, "$destination")
    }
}
Write-Host "Setting up the wallpaper..."
Set-WallPaper -Image $destination -Style Fill

if ($files -gt 0) {
    # We do not want to keep every file; remove the old ones
    Write-Host "Cleaning the directory..."
    $i = 1
    Get-ChildItem -Filter "????-??-??.jpg" $downloadFolder | Sort -Descending FullName | ForEach-Object {
        if ($i -gt $files) {
            # We have more files than we want, delete the extra files
            $fileName = $_.FullName
            Write-Debug "Removing file $fileName"
            Remove-Item "$fileName"
        }
        $i++
    }
}
