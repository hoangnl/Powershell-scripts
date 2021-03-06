$Market = "en-US"
$Resolution = "1920x1080"
$ImageFileName = "wallpaper.jpg"
$DownloadDirectory = "$env:USERPROFILE\Pictures\BingWallpaper"
$BingImageFullPath = "$($DownloadDirectory)\$($ImageFileName)"


New-Item -ItemType directory -Force -Path $DownloadDirectory | Out-Null

[xml]$Bingxml = (New-Object System.Net.WebClient).DownloadString("http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1&mkt=$($Market)");
$ImageUrl = "http://www.bing.com$($Bingxml.images.image.urlBase)_$($Resolution).jpg";

if ((Test-Path "$BingImageFullPath") -And ((Get-ChildItem "$BingImageFullPath").LastWriteTime.ToShortDateString() -eq (get-date).ToShortDatesTring())){
    Write-Host -ForegroundColor Green "You already have today's Bing image in: $DownloadDirectory"   
}
else {
    #Invoke-WebRequest -UseBasicParsing -Uri $ImageUrl -OutFile "$BingImageFullPath";
    (New-Object System.Net.WebClient).DownloadFile($ImageUrl, "$BingImageFullPath")
    Write-Host -ForegroundColor Green "Today's Bing image downloaded to: $DownloadDirectory" 
}

While (!(Test-Path "$BingImageFullPath")) {
    Write-Host -ForegroundColor Yellow "Waiting for Bing image to finish downloading..."
    Start-Sleep -Seconds 10
}
Add-Type @"
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
namespace Wallpaper
{
   public enum Style : int
   {
       Tile, Center, Stretch, NoChange
   }
   public class Setter {
      public const int SetDesktopWallpaper = 20;
      public const int UpdateIniFile = 0x01;
      public const int SendWinIniChange = 0x02;
      [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
      private static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
      public static void SetWallpaper ( string path, Wallpaper.Style style ) {
         SystemParametersInfo( SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange );
         RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true);
         switch( style )
         {
            case Style.Stretch :
               key.SetValue(@"WallpaperStyle", "2") ; 
               key.SetValue(@"TileWallpaper", "0") ;
               break;
            case Style.Center :
               key.SetValue(@"WallpaperStyle", "1") ; 
               key.SetValue(@"TileWallpaper", "0") ; 
               break;
            case Style.Tile :
               key.SetValue(@"WallpaperStyle", "1") ; 
               key.SetValue(@"TileWallpaper", "1") ;
               break;
            case Style.NoChange :
               break;
         }
         key.Close();
      }
   }
}
"@

#[Wallpaper.Setter]::SetWallpaper( '', 0 )
Write-Host $BingImageFullPath
[Wallpaper.Setter]::SetWallpaper( "$BingImageFullPath", 3)

Function Set-WallPaper($Value)

{
 Set-ItemProperty -path 'HKCU:\Control Panel\Desktop\' -name wallpaper -value $value
  Set-ItemProperty -path 'HKCU:\Control Panel\Desktop\' -name wallpaper -value $value
 rundll32.exe user32.dll, UpdatePerUserSystemParameters
 rundll32.exe user32.dll, UpdatePerUserSystemParameters 
rundll32.exe user32.dll, UpdatePerUserSystemParameters 
}
Set-WallPaper -value $BingImageFullPath