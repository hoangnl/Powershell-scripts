#*********************************************************************************************************** 
# This script installs SharePoint 2010 Prerequisites and Add required Windows Server 8 Beta Features/Roles 
# Do not run this script in any place other than within a Windows Server 8 Beta environment 
#  
# Run this script as Administrator
# 
# Step 1: Add Windows Server 8 Beta Roles/Features  
# Step 2: Modify IIS 8 Settings
# Step 3: Downloads Pre-Reqs from Microsoft Downloads site 
# Step 4: Install Pre-Reqs 
#
# Example of a download directory: C:\downloads
# *Do not include trailing '\'
#
# Don't forget to:
# Set-ExecutionPolicy RemoteSigned
# If you have not done so already within you Windows Server 8 Beta instance
#******************************************************************************************************* 
param([string]$PreReqInstallerDir = $(Read-Host -Prompt "Please Enter a directory path where I can store the downloaded Prerequisite files")) 
 
# Import Required Modules
Import-Module BitsTransfer 
Import-Module ServerManager
 
# Specify download url's for SharePoint 2010 prerequisites
$DownloadUrls = (                          
                    "http://download.microsoft.com/download/C/9/F/C9F6B386-824B-4F9E-BD5D-F95BB254EC61/Redist/amd64/Microsoft%20Sync%20Framework/Synchronization.msi", #Microsoft Sync Framework v1.0
                    "http://download.microsoft.com/download/0/A/2/0A28BBFA-CBFA-4C03-A739-30CCA5E21659/FilterPack64bit.exe", #Microsoft Office 2010 Filter Packs
                    "http://download.microsoft.com/download/c/c/4/cc4dcac6-ea60-4868-a8e0-62a8510aa747/MSChart.exe", #Microsoft Chart Controls for Microsoft .NET Framework 3.5
                    "http://download.microsoft.com/download/3/5/5/35522a0d-9743-4b8c-a5b3-f10529178b8a/sqlncli.msi", #Microsoft SQL Server 2008 Native Client
                    "http://download.microsoft.com/download/A/D/0/AD021EF1-9CBC-4D11-AB51-6A65019D4706/SQLSERVER2008_ASADOMD10.msi", #Microsoft SQL Server 2008 Analysis Services ADOMD.NET                   
                    "http://download.microsoft.com/download/1/0/F/10F1C44B-6607-41ED-9E82-DF7003BFBC40/1033/x64/rsSharePoint.msi", #SQL Server 2008 R2 Reporting Services Add-in for Microsoft SharePoint Technologies 2010
                    "http://download.microsoft.com/download/8/D/F/8DFE3CE7-6424-4801-90C3-85879DE2B3DE/Platform/x64/SpeechPlatformRuntime.msi", #Microsoft Server Speech Platform
                    "http://download.microsoft.com/download/E/0/3/E033A120-73D0-4629-8AED-A1D728CB6E34/SR/MSSpeech_SR_en-US_TELE.msi"  #Speech recognition language for English       
                ) 
 
 
function AddWindowsFeatures() 
{ 
    Write-Host ""
    Write-Host "====================================================================="
    Write-Host "Step 1 of 4: Enabling required Windows Roles/Features." 
    Write-Host ""
    Write-Host "This may take a few minutes. Please wait..."
    Write-Host ""
    Write-Host "Note: You'll receive a warning after this is done if"
    Write-Host "      Windows automatic updates is not enabled."
    Write-Host "=====================================================================" 
     
    $ReturnCode = 0   
    
      
    # Note: You can use the Get-WindowsFeature cmdlet (its in the ServerManager
    #       module) to get a listing of all features and roles. This is how I 
    #        figured out the Role/Feature names below
    $WindowsFeatures = @(
    			"Web-Server",
    			"Web-Common-Http",
    			"Web-Default-Doc",
    			"Web-Dir-Browsing",
    			"Web-Http-Errors",
    			"Web-Static-Content",
    			"Web-App-Dev",
    			"Web-Asp-Net",
    			"Web-Net-Ext",
    			"Web-ISAPI-Ext",
    			"Web-ISAPI-Filter",
    			"Web-Health",
    			"Web-Http-Logging",
    			"Web-Log-Libraries",
    			"Web-Request-Monitor",
    			"Web-Http-Tracing",
    			"Web-Custom-Logging",
    			"Web-Scripting-Tools",
    			"Web-Security",
    			"Web-Basic-Auth",
    			"Web-Windows-Auth",
    			"Web-Digest-Auth",
    			"Web-Filtering",
    			"Web-Performance",
    			"Web-Stat-Compression",
    			"Web-Dyn-Compression",
    			"Web-Mgmt-Tools",
    			"Web-Mgmt-Console",
    			"Web-Mgmt-Compat",
    			"Web-Metabase",
    			"Web-WMI",
    			"WAS",
    			"WAS-Process-Model",
    			"WAS-NET-Environment",
    			"WAS-Config-APIs",
    			"NET-Framework-Features",
    			"NET-Framework-Core",
    			"NET-HTTP-Activation",
    			"NET-Non-HTTP-Activ",
			"Windows-Identity-Foundation"
    )
    Try 
    { 
	# Create PowerShell to execute 
        $myCommand = 'Add-WindowsFeature ' + [string]::join(",",$WindowsFeatures)

	# Execute $myCommand
        Invoke-Expression $myCommand                  
        
        Write-Host " - Done enabling all required Windows Roles/Features" 
    } 
    Catch 
    { 
        $ReturnCode = -1 
        Write-Warning "Error when Adding Windows Features." 
        Write-Error $_ 
        break 
    } 
     
    return $ReturnCode  
} 
 


function ModifyIIS8Settings() {


    Write-Host ""
    Write-Host "====================================================================="
    Write-Host "Step 2 of 4: Modify IIS 8 Settings" 
    Write-Host "=====================================================================" 
    $ReturnCode = 0

    Try 
    { 

	Import-Module WebAdministration
	Set-WebConfigurationProperty '/system.applicationHost/applicationPools/applicationPoolDefaults' -Name managedRuntimeVersion -value v2.0
	Write-Host " - Changed Application Pool Default .NET Framework Version to v2.0" 
    } 
    Catch 
    { 
         $ReturnCode = -1 
         Write-Warning "Error while changing the managedRuntimeVersion for IIS 8 to v2.0" 
         Write-Error $_ 
         break 
    }     
    
    return $ReturnCode 

}


function DownLoadPreRequisites() 
{ 

    Write-Host ""
    Write-Host "====================================================================="
    Write-Host "Step 3 of 4: Downloading SharePoint 2010 Prerequisites Please wait..." 
    Write-Host "====================================================================="
     
    $ReturnCode = 0 
 
    ForEach ($DownLoadUrl in $DownloadUrls) 
    { 
        ## Get the file name based on the portion of the URL after the last slash 
        $FileName = $DownLoadUrl.Split('/')[-1] 
        Try 
        { 
            ## Check if destination file already exists 
            If (!(Test-Path "$PreReqInstallerDir\$FileName")) 
            { 
                ## Begin download 
                Start-BitsTransfer -Source $DownLoadUrl -Destination $PreReqInstallerDir\$fileName -DisplayName "Downloading `'$FileName`' to $PreReqInstallerDir" -Priority High -Description "From $DownLoadUrl..." -ErrorVariable err 
                If ($err) {Throw ""} 
            } 
            Else 
            { 
                Write-Host " - File $FileName already exists, skipping..." 
            } 
        } 
        Catch 
        { 
            $ReturnCode = -1 
            Write-Warning " - An error occurred downloading `'$FileName`'" 
            Write-Error $_ 
            break 
        } 
    } 
    Write-Host " - Done downloading Prerequisites required for SharePoint 2010" 
     
    return $ReturnCode 
} 
 
 
function InstallPreReqs() 
{ 

    $ReturnCode = 0

    Write-Host ""
    Write-Host "====================================================================="
    Write-Host "Step 4 of 4: Installing Prerequisites required for SharePoint 2010" 
    Write-Host ""
    Write-Host "You'll be prompted to install each prerequisite. "
    Write-Host "Please follow the screen prompts for each one accordingly." 
    Write-Host ""
    Write-Host "Keep an eye on your taskbar as some installs "
    Write-Host "may not open and blink in the taskbar" 
    Write-Host "=====================================================================" 
     
     
    #Install rest of the Prerequisites that were downloaded 
    ForEach ($DownLoadUrl in $DownloadUrls) 
    { 
        ## Get the file name based on the portion of the URL after the last slash 
        $FileName = $DownLoadUrl.Split('/')[-1] 
        Try 
        { 
            ## Check if destination file already exists 
            If (Test-Path "$PreReqInstallerDir\$FileName") 
            { 
                $StartInfo = New-Object -TypeName System.Diagnostics.ProcessStartInfo  
                $StartInfo.FileName = "$PreReqInstallerDir\$FileName"  
                $InstallProcesss = [System.Diagnostics.Process]::Start($StartInfo); 
             
                While(-not($InstallProcesss.HasExited)) 
                { 
                    Start-Sleep -Seconds 1         
                } 
                Write-Host " - $FileName processed" 
            } 
        } 
        Catch 
        { 
            $ReturnCode = -1 
            Write-Warning " - An error occurred while installing`'$FileName`'" 
            Write-Error $_ 
            break 
        }     
    } 
    return $ReturnCode 
} 
 
function CheckProvidedDownloadPath()
{


    $ReturnCode = 0

    Try 
    { 
        # Check if destination path exists 
        If (Test-Path $PreReqInstallerDir) 
        { 
           # Remove trailing slash if it is present
           $script:PreReqInstallerDir = $PreReqInstallerDir.TrimEnd('\')
	   $ReturnCode = 0
        }
        Else {

	   $ReturnCode = -1
           Write-Host ""
	   Write-Warning "Your specified download path does not exist. Please verify your download path then run this script again."
           Write-Host ""
        } 


    } 
    Catch 
    { 
         $ReturnCode = -1 
         Write-Warning "An error has occurred when checking your specified download path" 
         Write-Error $_ 
         break 
    }     
    
    return $ReturnCode 

}


 
function InstallFeaturesAndPreReqs() 
{ 

    $rc = 0 
    
    $rc = CheckProvidedDownloadPath

    # Step 1 - Add the Windows Features/Roles
    if($rc -ne -1) 
    { 
        $rc = AddWindowsFeatures 
    } 
     
    # Step 2 - Change IIS .NET Application Pools Default to v2.0
    if($rc -ne -1) 
    { 
        $rc = ModifyIIS8Settings 
    }

    # Step 3 - Download Pre-Reqs  
    if($rc -ne -1) 
    { 
        $rc = DownLoadPreRequisites 
    } 
     
    # Step 4 - Install the Pre-Reqs 
    if($rc -ne -1) 
    { 
        $rc = InstallPreReqs 
    } 

    if($rc -ne -1)
    {

        Write-Host ""
        Write-Host "================================================================"
        Write-Host "Script execution is now complete!"
        Write-Host ""
        Write-Host "For complete details on the SharePoint 2010 installation procedure on Windows Server 8 Beta, visit:"
        Write-Host "http://craiglussier.com/2012/03/01/install-sharepoint-2010-on-windows-server-8-beta/"
        Write-Host ""
        Write-Host ""
    }


} 
 
InstallFeaturesAndPreReqs