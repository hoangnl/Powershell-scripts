$Path = "C:\Users\HOANG\Desktop\API\"
[String]$Folder = "InBox"
[String]$Test ="Unread"
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNameSpace("MAPI")
$NameSpace.Folders.Item(1)
$Email = $NameSpace.Folders.Item(1).Folders.Item($Folder).Items
$body = $Email | foreach {
  if($_.Unread -eq $True) {
   $_.body
  }
}
write-host "----------------"
$body | % {
    #Write-Host $_
    $message = $_.Substring(0,3)
    #Write-Host $message
    $MessageFolder = ($Path + $message)
    If((test-path $MessageFolder) -eq 0)
    {
        New-Item -ItemType Directory -Force -Path $MessageFolder
    }
    $array = $_.Split(“`n”)
    $s = $array[1]
    $sender = ""
    if($message -ne "FFM")
    {
        $sender = $s.SubString($s.IndexOf("-") + 1, $s.LastIndexOf("/") - $s.IndexOf("-") - 1)
    }
    else
    {
        $index = $s.IndexOf('/', $s.IndexOf('/') + 1);
        #Write-Host "index {0}" - f $index
        $sender = $s.SubString($s.IndexOf("/") + 1, $index - $s.IndexOf("/") - 1)
    }
    #Write-Host $sender
    $SenderFolder = ($MessageFolder + "\" + $sender)
    If((test-path $SenderFolder) -eq 0)
    {
        New-Item -ItemType Directory -Force -Path $SenderFolder
    }
    $_ >> ($SenderFolder + "\" + ("{0:yyyyMMddHHmmss}" -f (get-date)) + ".txt") 
}

