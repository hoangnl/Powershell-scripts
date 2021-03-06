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
get-childitem -path $Destination -recurse | Remove-Item -Recurse -Force
$CopyList| % {
    
    if((Test-Path $_) -ne 0){
        if((Get-Item $_) -is [System.IO.DirectoryInfo]){
            $folder = (Get-Item $_).Name
            $sourcePath = $_  
            Get-ChildItem $sourcePath -Recurse -Include '*.ascx' | Foreach-Object `
            {
                $destDir = Split-Path ($_.FullName -Replace [regex]::Escape($sourcePath), ($Destination + "\" + $folder))
                if (!(Test-Path $destDir))
                {
                    New-Item -ItemType directory $destDir | Out-Null
                }
                Copy-Item $_ -Destination $destDir
            }
        }
        else {
            Copy-Item -Path $_ -Destination $Destination
        }
    }
    
}
Read-Host "Complete. Enter to exit"