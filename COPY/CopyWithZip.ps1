$CopyList = @(
"E:\desktop\NSW\06.SOURCE\CNV\Community\NSW.Dtos\bin\Debug\NSW.Dtos.dll",
"E:\desktop\NSW\06.SOURCE\CNV\Community\NSW.DataAccess\bin\Debug\NSW.DataAccess.dll",
"E:\desktop\NSW\06.SOURCE\CNV\Community\NSW.Proxy\bin\Debug\NSW.Proxy.dll",
"E:\desktop\NSW\06.SOURCE\CNV\Synchronize\NSW.Synchronize\bin\NSW.Synchronize.dll",
"E:\desktop\NSW\06.SOURCE\CNV\Synchronize\NSW.Publish\bin\NSW.Publish.dll",
"E:\desktop\NSW\06.SOURCE\CNV\NSW.Web\bin\NSW.Web.dll",
"E:\desktop\NSW\06.SOURCE\CNV\Community\NSW.Library\bin\Debug\NSW.Library.dll",
"E:\desktop\NSW\06.SOURCE\CNV\NSW.Web\UserControls"
)
$Destination = "E:\desktop\Deploy"
$winrar = "c:\program files\winrar\Rar.exe" 
$UnRAR = "c:\program files\winrar\UnRAR.exe" 
get-childitem -path $Destination -recurse | Remove-Item -Recurse -Force
$CopyList| % {
    
    if((Test-Path $_) -ne 0){
        if((Get-Item $_) -is [System.IO.DirectoryInfo]){
            $folder = (Get-Item $_).Name
            &$winrar a -r -ep1 ($_ + ".rar") $_
            Copy-Item -Path ($_ + ".rar") -Destination $Destination
            Write-Host ($Destination + "\" + $folder + ".rar")
            &$unrar x -y ($Destination + "\" + $folder + ".rar") ($Destination)
            get-childitem -path $Destination -recurse -include "*.cs" | Remove-Item -Force
            #get-childitem -path $Destination -recurse -include "*.rar" | Remove-Item -Force
        }
        else {
            Copy-Item -Path $_ -Destination $Destination
        }
    }
    
}
Read-Host "Complete. Enter to exit"