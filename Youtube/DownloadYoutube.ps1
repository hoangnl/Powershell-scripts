$Path = "C:\Users\HOANG\Desktop\"
$Playlisturl = "http://www.youtube.com/playlist?list=PL1058E06599CCF54D"
$VideoUrls= (invoke-WebRequest -uri $Playlisturl).Links | ? {$_.HREF -like "/watch*"} | `
? innerText -notmatch ".[0-9]:[0-9]." | ? {$_.innerText.Length -gt 3} | Select innerText, `
@{Name="URL";Expression={'http://www.youtube.com' + $_.href}} | ? innerText -notlike "*Play all*"
 
$VideoUrls
 
ForEach ($video in $VideoUrls){
    Write-Host ("Downloading " + $video.innerText)
    C:\Users\HOANG\Desktop\Git\Powershell\youtube-dl.exe $video.URL --output ($Path + "%(title)s.%(ext)s")
}