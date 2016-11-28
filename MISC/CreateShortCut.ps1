$Shell = New-Object -ComObject ("WScript.Shell")
$ShortCut = $Shell.CreateShortcut($env:USERPROFILE + "\Desktop\Notepad++.lnk")
$ShortCut.TargetPath="C:\Program Files (x86)\Notepad++\notepad++.exe"
#$ShortCut.Arguments="-arguementsifrequired"
$ShortCut.WorkingDirectory = "C:\Program Files (x86)\Notepad++";
$ShortCut.WindowStyle = 1;
$ShortCut.Hotkey = "CTRL+SHIFT+F";
$ShortCut.IconLocation = "notepad++.exe, 0";
$ShortCut.Description = "Your Custom Shortcut Description";
$ShortCut.Save()