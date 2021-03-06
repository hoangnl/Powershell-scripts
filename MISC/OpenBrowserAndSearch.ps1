 Function Start-Google
{
<#
.Synopsis
Searches the Googes
.DESCRIPTION
Lets you quickly start a search Google from within Powershell
.EXAMPLE
Start-Google -search PowerShell
#>
                      [CmdletBinding()]
              Param ( [Parameter(Mandatory=$false,
                      ValueFromPipelineByPropertyName=$true,
                      Position=0)]
                      $Search = "The Overnight Admin",
                      [Parameter(Mandatory=$false,
                      ValueFromPipelineByPropertyName=$true,
                      Position=0)]
                      $google = "https://www.google.com/search?q="
              )
         Begin {
                    $ie = new-object -com internetexplorer.application
         }
    Process {
                $Search | ForEach-Object { $google = $google + "$_+"}
    }
 End
     {
         $url = $google.Substring(0,$google.Length-1)
         $ie.navigate( $url )
         $ie.visible = $true
 }
}
Start-Google -Search "Enter Hilarious Search Here"