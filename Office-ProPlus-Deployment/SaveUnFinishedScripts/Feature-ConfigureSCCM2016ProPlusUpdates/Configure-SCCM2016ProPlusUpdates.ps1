Function Run-365OfficeUpdateSetup {
<#
.SYNOPSIS
This function will deploy a specified Office 365 update using SCCM.

.DESCRIPTION
Suplpying at least a SharePath and a Branch this function will download the branch updates
to a shared folder. A deployment packaged and device collection will be created. The deployment 
package will then be deployed to the collection.

.PARAMETER SharePath
Required. The folder path of the share to download the updates to or if it does not exist will be
created and shared.

.PARAMETER Branch
Required. The branch number of the update to download and deploy.

.PARAMETER ShareName
The name of a share to download the updates to. If not specified the folder in SharePath will be
the ShareName.

.PARAMETER SiteServer
The name of the SCCM Site Server. If not specified the local host will be considered the site server.

.PARAMETER SiteCode
The name of the Site Code.

.PARAMETER UNCSourcePath
The UNC path to the shared folder.

.PARAMETER Description
The description of the deployment package.

.PARAMETER DistributionPoint
The name of the distribution point.

.PARAMETER Bitness
The bit of Office to deploy.

.PARAMETER DeploymentPackageName
The name of the deployment package.

.PARAMETER CollectionName
The name of the collection that will be created.
#>
Param(
    [Parameter(Mandatory=$True)]
    [string]$SharePath,
    [Parameter(Mandatory=$True)]      
	[string]$Branch, #ex:"FRCB1602",
	[string]$ShareName,        
	[string]$SiteServer = $env:COMPUTERNAME,
	[string]$SiteCode,           
	[string]$UNCSourcePath,      
	[string]$Description = "Deploys office 365 latest updates",       
	[string]$DistributionPoint,     
	[string]$Bitness = "32",       
	[string]$DeploymentPackageName = "O365 Clients - English",             
	[string]$CollectionName = "Office 2016 Clients",    
    [string]$Path = $SharePath      
)

    Import-Module ((Split-Path $env:SMS_ADMIN_UI_PATH)+"\ConfigurationManager.psd1")
    $PSD = Get-PSDrive -PSProvider CMSite
    CD "$($PSD):"

    if(!$SiteCode){
        $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
    }

    if(!$DistributionPoint){
        $DistributionPoint = (Get-CMDistributionPointInfo).Name
    }

    if(!$UNCSourcePath){
        $drive = (Split-Path $SharePath -Qualifier).Replace(':','')
        $leaf = Split-Path $SharePath -NoQualifier
        $UNCSourcePath = "\\$SiteServer$leaf"
    }
    
    #Create a shared folder and set the permission
    CreateFolder -SharePath $SharePath -ShareName $ShareName
    
    #Create the deployment package
    write-host "Creating the deployment package..."
    Setup-NewSoftwareUpdateDeploymentPackage -SiteServer $SiteServer -Name $DeploymentPackageName -Description $Description -SourcePath $UNCSourcePath -DistributionPoint $DistributionPoint

    #Get updates from catalog and import them
    Write-Host "Collecting the update display name from the catalog..."
    Write-Host ""
    $global:ouidp = $Branch
    $global:bitness = $Bitness
    $global:sun = Get-CMSoftwareUpdate | where {$_.LocalizedDisplayName -like "*$ouidp*" -and $_.LocalizedDisplayName -like "*$bitness*"}
       
    $SharePath = $SharePath + "\" + "en-us"
    Write-Host "Downloading the updates..."
    Write-Host ""
    Download-SoftwareUpdateFiles -Branch $Branch -Path $SharePath -Bitness $Bitness -SiteCode $SiteCode -SiteServer $SiteServer -DeploymentPackageName $DeploymentPackageName | Out-Null 

    #Create the device collection
    Create-DeviceCollection -CollectionName $CollectionName -SiteCode $SiteCode
 
    #Deploy package
    Write-Host "Deploying the updates..."
    Write-Host ""
    Deploy-SoftwareUpdates -Branch $Branch -CollectionName $CollectionName -Bitness $Bitness
 
    #Check Status
    Write-Host "Checking deployment status..."
    Check-DeploymentStatus -Branch $Branch -Bitness $Bitness
}

Function CreateFolder{
<# 
.SYNOPSIS 
    Creates a folder and sets permission
.DESCRIPTION 
    Creates a folder, sets proper permissions, and shares the folder
.PARAMETER FolderName
    Specify the name of the folder to be created
.PARAMETER FolderPath
    Specify the path that the folder will be created in
.EXAMPLE 
    CreateFolder -FolderName Downloaded-Files -FolderPath \\Server1\Updates
    Will create a folder named "Downloaded-Files" in the directory "\\Server1\Updates" with proper permissions and sharing
#> 
Param(
    [Parameter(Mandatory=$True)]  
	[String]$SharePath,   
	[string]$ShareName
)

    if(!$SharePath){
        $drive = Read-Host "Please enter the drive letter to create the share on"
        if(!$ShareName){
            $checkSharename = Get-ChildItem $drive | Where-Object {$_.Name -eq "Packages"}
            if(!$checkSharename){
                $ShareName = "Packages"
                $global:SharePath = $drive + $ShareName
            }
            else{
                $ShareName = Read-Host "A share already exists in $drive called Packages. Enter yes to continue in this share, or enter the name of the folder to be created and shared"
                if(($ShareName -eq "yes") -or ($ShareName -eq 'y')){
                    $ShareName = "Packages"
                    $global:SharePath = $drive + $ShareName                  
                }
                elseif(($ShareName -ne "yes") -or ($ShareName -ne "y")){
                    $ShareName = $ShareName
                    $global:SharePath = $drive + $ShareName
                }
            }
        }
        else{
            $global:SharePath = $drive + $ShareName
        }     
    }
    else{
        $ShareName = $SharePath.Split("\")[-1] 
    }

    if(!(Test-Path $SharePath)){
        New-Item "$SharePath\en-us" -Type Directory | Out-Null
        $checkShareName = Get-Item $SharePath | Where-Object {$_.Name -eq $ShareName}
        if($checkSharename){
            Write-Host ""
            Write-host "The folder $ShareName has been created."
            Write-Host ""
        }
        $shareString = $ShareName.ToString()+'='+$SharePath
        net share $ShareString '/Grant:Authenticated Users,Change'
      
        $acl = Get-Acl $SharePath
        $permission = "Everyone","FullControl","ContainerInherit,ObjectInherit","None","Allow"
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
        $acl.SetAccessRule($accessRule)
        $acl | Set-Acl $SharePath
    } 
}

Function Download-SoftwareUpdateFiles{
<# 
.SYNOPSIS 
    Downloads files for Software Update
.DESCRIPTION 
    Downloads files for Software Update
.PARAMETER Branch
    Specify the name of the Office 365 branch update to download, ie "FRCB1601"
.PARAMETER SiteCode
    Specify the 3 character server code that the collection is being added to
.PARAMETER Path
    Specify the path for the files to be downloaded to
.Parameter DeploymentPackageName
    Specify the name of the deployment package the update files will be added to
.EXAMPLE 
    Download-SoftwareUpdateFiles -Branch -SiteCode S01 -Path \\Server1\Updates\en-us
    Downloads files for the Office 365 update "FRCB1601" on the Server "S01" to the path \\Server1\Updates\en-us
#> 
Param(
    [Parameter(Mandatory=$true)]  
	[string]$Branch,
    [Parameter(Mandatory=$True)]	
	[String]$Path = $NULL,	
	[String]$Bitness, 
	[String]$SiteCode, 
	[String]$SiteServer = $env:COMPUTERNAME, 
    [String]$DeploymentPackageName,
    [String]$UNCSourcePath
)
    if(!$Bitness){
        $Bitness = "32"
    }

    if(!$SiteCode){
        $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
    }

    $hash = @{
        #open a CIM sesison to the ConfigMgr Server
        cimsession = New-CimSession -ComputerName $SiteServer
        NameSpace = 'Root\SMS\Site_S01' #point to the SMS Namespace
        ErrorAction = 'Stop'
    }

    Import-Module ((Split-Path $env:SMS_ADMIN_UI_PATH)+"\ConfigurationManager.psd1")
    $PSD = Get-PSDrive -PSProvider CMSite
    CD "$($PSD):"

    ##Get the  Display Name of the update
    if(!$sun){
        Write-Host "Collecting the update display name from the catalog..."
        $ouidp = $Branch
        $bitness = $Bitness
        $sun = Get-CMSoftwareUpdate | where {$_.LocalizedDisplayName -like "*$ouidp*" -and $_.LocalizedDisplayName -like "*$bitness*"}
    }

    $officeupdatename = $sun.LocalizedDisplayName

    [System.Xml.XmlDocument]$SoftwareUpdateXml = New-Object System.Xml.XmlDocument

    $SoftwareUpdateXml.LoadXml($sun.SDMPackageXML) | Out-Null

    $SaveFolderPaths = $SoftwareUpdateXml.GetElementsByTagName('Updates')

    foreach($tempPath in $SaveFolderPaths){
        if($tempPath.SoftwareUpdateReference){
            #Write-Host $tempPath.SoftwareUpdateReference.LogicalName
            [String]$SaveFolderPath = $tempPath.SoftwareUpdateReference.LogicalName
            $SaveFolderPath = $SaveFolderPath.Replace("SUM_","")
        }
    }

    $KBID = $sun.CI_ID

    [string]$KBID = $KBID
    
    [array]$CIIDs = @()

    if(!$SiteCode){
        $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
    }
     
    $CIID = (gwmi -ns root\sms\site_$($SiteCode) -class SMS_SoftwareUpdate | where {$_.CI_ID -eq $KBID}).CI_ID
    $CIIDs += $CIID
       
    $DownloadInfo = foreach ($CI_ID in $CIIDs){
        $contentID = Get-CimInstance -Query "Select * from SMS_CITOContent Where CI_ID=$CI_ID"  @hash
        #Filter out the English Local and ContentID's not targeted to a particular Language
        $contentID = $contentID  | Where {($_.ContentLocales -contains "Locale:0")}

        foreach ($id in $contentID){
            $ContentIDManger = $id.ContentID
            $contentFile = Get-CimInstance -Query "Select * from SMS_CIContentfiles WHERE ContentID=$($ID.ContentID)AND (LanguageID=1033 OR LanguageID=0)" @hash
            [pscustomobject]@{Source = $contentFile.SourceURL ;
                              ContentID = $contentFile.ContentID ;
                              FileHash = $contentFile.FileHash ;
                              FileVersion = $contentFile.FileVersion ;
                              LanguageID = $contentFile.LanguageID ;
                              IsSigned = $contentFile.IsSigned ;
                              Destination = $Path;
            }
        }    
    }

    $UpdateContentIDs = @()
    $UpdateContentSourcePaths = @()

    foreach($tempVar in $DownloadInfo){
        foreach($tempVar2 in $tempVar.Source){
        [string]$tempString = $tempVar2
        $tempString =  $tempString.Substring($tempString.LastIndexOf(("office")))
        $tempString = [string]$tempVar.Destination + "\" + $SaveFolderPath + "\" + $tempString
        $tempString = $tempString.Replace("/","\")    
        $tempDir = Split-Path -Path $tempString
        if (!(Test-Path $tempDir)){
            New-Item $tempDir -Type Directory
            }
            if(Test-Path -Path $tempString){
                if($tempString.Contains("v32")){
                    Rename-Item $tempString v32.cab
                }
                elseif($tempString.Contains("v64")){
                    Rename-Item $tempString v64.cab
                }
            }

        Start-BitsTransfer -Destination $tempString -Source $tempVar2
        
        }
    }

    $UpdateContentIDs += [UInt32]$ContentIDManger
    $UpdateContentSourcePaths += $UNCSourcePath
    
    Add-Content -SiteCode $SiteCode -SiteServer $SiteServer -ContentIDs $UpdateContentIDs -ContentSourcePath $UpdateContentSourcePaths -DeploymentPackageName $DeploymentPackageName
}

 Function Setup-NewSoftwareUpdateDeploymentPackage{
 <# 
.SYNOPSIS 
    Create a Deployment Package in Configuration Manager 2012. 
.DESCRIPTION 
    Use this script if you need to create a Deployment Package in Configuration Manager 2012.  
.PARAMETER SiteServer 
    Primary Site server name with SMS Provider installed 
.PARAMETER Name 
    Name of the Deployment Package 
.PARAMETER Description 
    Description of the Deployment Package 
.PARAMETER SourcePath 
    UNC path to the source location where downloaded patches will be stored 
.EXAMPLE 
    .\New-CMDeploymentPackage.ps1 -SiteServer CM01 -Name "Critical and Security Patches" -SourcePath "\\CAS01\Source$\SUM\ADRs\CS" -Description "Contains Critical and Security patches" 
    Create a Deployment Package called 'Critical and Security Patches', specifying a source path and description on a Primary Site server called 'CM01': 
#> 
[CmdletBinding(SupportsShouldProcess=$true)] 
param( 
    [parameter(Mandatory=$true)] 
    [string]$SourcePath,
        
    [string]$SiteServer = $env:COMPUTERNAME,
    
    [string]$SiteCode, 
     
    [string]$Name = "O365 Clients - English", 
     
    [string]$Description = "Deploys office 365 latest updates", 

    [string]$DistributionPoint
) 
Begin{
    if(!$SiteCode){
        $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
    }
    
    if(!$DistributionPoint){
        $DistributionPoint = (Get-CMDistributionPointInfo).Name
    } 
} 
Process { 
    function Get-DuplicateInfo { 
        $IsDuplicatePkg = $false 
        $EnumDeploymentPackages = Get-CimInstance -CimSession $CimSession -Namespace "root\SMS\site_$($SiteCode)" -ClassName SMS_SoftwareUpdatesPackage -ErrorAction SilentlyContinue -Verbose:$false 
        foreach ($Pkgs in $EnumDeploymentPackages) { 
            if ($Pkgs.PkgSourcePath -like "$($SourcePath)") { 
                $IsDuplicatePkg = $true 
            } 
        } 
        return $IsDuplicatePkg 
    } 
    function Remove-CimSessions { 
        foreach ($Session in $(Get-CimSession -ComputerName $SiteServer -ErrorAction SilentlyContinue -Verbose:$false)) { 
            if ($Session.TestConnection()) { 
                Write-Verbose -Message "Closing CimSession against '$($Session.ComputerName)'" 
                Remove-CimSession -CimSession $Session -ErrorAction SilentlyContinue -Verbose:$false 
            } 
        } 
    } 
    try { 
        Write-Verbose -Message "Establishing a Cim session against '$($SiteServer)'" 
        $CimSession = New-CimSession -ComputerName $SiteServer -Verbose:$false 
        # Check if there's an existing Deployment Package with the same name 
        if ((Get-CimInstance -CimSession $CimSession -Namespace "root\SMS\site_$($SiteCode)" -ClassName SMS_SoftwareUpdatesPackage -Filter "Name like '$($Name)'" -ErrorAction SilentlyContinue -Verbose:$false | Measure-Object).Count -eq 0) { 
            # Check if there's an existing Deployment Package with the same source path 
            if ((Get-DuplicateInfo) -eq $false) { 
                $CimProperties = @{ 
                    "Name" = "$($Name)" 
                    "PkgSourceFlag" = [UInt32]2 
                    "PkgSourcePath" = "$($SourcePath)" 
                    "SourceVersion" = [UInt32]2
                } 
                if ($PSBoundParameters["Description"]) { 
                    $CimProperties.Add("Description",$Description) 
                } 
                $CMDeploymentPackage = New-CimInstance -CimSession $CimSession -Namespace "root\SMS\site_$($SiteCode)" -ClassName SMS_SoftwareUpdatesPackage -Property $CimProperties -Verbose:$false -ErrorAction Stop
                $PSObject = [PSCustomObject]@{ 
                    "Name" = $CMDeploymentPackage.Name 
                    "Description" = $CMDeploymentPackage.Description 
                    "PackageID" = $CMDeploymentPackage.PackageID 
                    "PkgSourcePath" = $CMDeploymentPackage.PkgSourcePath 
                    "SourceVersion" = $CMDeploymentPackage.SourceVersion
                    #"DistributionPoint" = $CMDeploymentPackage.DistributionPoint
                    
                } 
                             
                Write-Output $PSObject 
                
            } 
            else { 
                Write-Warning -Message "A Deployment Package with the specified source path already exists" 
            } 
        } 
        else { 
            Write-Warning -Message "A Deployment Package with the name '$($Name)' already exists" 
        } 
    } 
    catch [Exception] { 
        Remove-CimSessions 
        Throw $_.Exception.Message 
    }
    Add-DistributionPoint -SiteCode $SiteCode -DistributionPoint $DistributionPoint -DeploymentPackageName $Name -SiteServer $SiteServer
} 
End{ 
    # Remove active Cim session established to $SiteServer 
    Remove-CimSessions 
}
}

 Function Create-DeviceCollection{
 <# 
.SYNOPSIS 
    Creates a device collection
.DESCRIPTION 
    Creates a device collection
.PARAMETER CollectionName
    Specify the name of the device collection that is being created
.PARAMETER SiteCode
    Specify the 3 character server code that the collection is being added to
.EXAMPLE 
    CreateCollection -CollectionName "Office 365 updates" -SiteCode S01
    Creates a collection named "Office 365 updates" on the server "S01"
#>  
Param( 
    [string]$CollectionName = "Office 2016 Clients",

    [string]$SiteCode
)
    Write-Host "Creating the device collection for Office 2016 clients..."
  
    if(!$SiteCode){
        $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
    }

    Import-Module ((Split-Path $env:SMS_ADMIN_UI_PATH)+"\ConfigurationManager.psd1")
    $PSD = Get-PSDrive -PSProvider CMSite
    CD "$($PSD):"

    $CMCollection = ([WMIClass]”root\sms\site_$($SiteCode):SMS_Collection”).CreateInstance()

    $CMCollection.Name = $CollectionName        
    $CMCollection.LimitToCollectionID = “SMS00001”
    $CMCollection.RefreshType = 2
    $CMCollection.Put() | Out-Null

    $CMRule = ([WMIClass]”root\sms\site_S01:SMS_CollectionRuleQuery”).CreateInstance()

    $CMRule.QueryExpression=”Select SMS_R_System.ResourceId, SMS_R_System.ResourceType, SMS_R_System.Name, SMS_R_System.SMSUniqueIdentifier, SMS_R_System.ResourceDomainORWorkgroup, SMS_R_System.Client from  SMS_R_System inner join SMS_G_System_ADD_REMOVE_PROGRAMS_64 on SMS_G_System_ADD_REMOVE_PROGRAMS_64.ResourceID = SMS_R_System.ResourceId inner join SMS_G_System_ADD_REMOVE_PROGRAMS on SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceId = SMS_R_System.ResourceId where SMS_G_System_ADD_REMOVE_PROGRAMS_64.DisplayName like `"Office 16%`" or SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName like `"Office 16%`"”

    $CMRule.RuleName = “Office 2016 Query”

    $CMCollection.AddMembershipRule($CMRule) | Out-Null

    $CMSchedule = ([WMIClass]"root\sms\site_S01:SMS_ST_RecurInterval").CreateInstance()

    $CMSchedule.DaySpan = “1”

    $CMSchedule.StartTime = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime((Get-Date).ToString())

    $CMCollection.RefreshSchedule=$CMSchedule

    $CMCollection.RefreshType = 6

    $CMCollection.Put() | Out-Null

    #Check if the collection was created successfully
    $checkCollection = Get-CMCollection -Name $CollectionName
    if(!$checkCollection){
        Write-Warning "Failed to create collection $CollectionName"
    }
    else{
        Write-Host ""
        Write-Host "The device collection $CollectionName has been created successfully."
        Get-CMCollection -Name $CollectionName | select Name,MemberCount,CollectionId | ft -AutoSize
    }
 }

 Function Deploy-SoftwareUpdates{
 <# 
.SYNOPSIS 
    Deploy Software update
.DESCRIPTION 
    Deploys a WSUS software update
.PARAMETER Branch
    Specify the branch ID for the update looks like "FRCB1601"
.PARAMETER CollectionName
    Specify the name of the device collection the update should be deployed to
.PARAMETER Bitness
    Specify 32 vs 64 bit
.EXAMPLE 
    DeploySoftwareUpdates -Branch FRCB1601 -Bitness 64 -CollectionName "Office 365 updates"
    Deploys the 64 bit version of the software update "FRCB1601" to the collection "Office 365 updates"
#>  
Param(
    [Parameter(Mandatory=$true)]  
	[string]$Branch,
	
	[String]$CollectionName = "Office 2016 Clients",
	
	[String]$Bitness = "32"
)
    Import-Module ((Split-Path $env:SMS_ADMIN_UI_PATH)+"\ConfigurationManager.psd1")
    $PSD = Get-PSDrive -PSProvider CMSite
    CD "$($PSD):"

    ##Get the  Display Name of the update
    if(!$sun){
        Write-Host "Collecting the update display name from the catalog..."
        $ouidp = $Branch
        $bitness = $Bitness
        $sun = Get-CMSoftwareUpdate | where {$_.LocalizedDisplayName -like "*$ouidp*" -and $_.LocalizedDisplayName -like "*$bitness*"}
    }

    $officeupdatename = $sun.LocalizedDisplayName
        
    Start-CMSoftwareUpdateDeployment -SoftwareUpdateName $officeupdatename -CollectionName $CollectionName        
 }

 Function Check-DeploymentStatus{
 <# 
.SYNOPSIS 
    Checks status of a Update that was deployed
.DESCRIPTION 
    Checks status of a Update that was deployed
.PARAMETER Branch
    Specify the branch ID for the update looks like "FRCB1601"
.PARAMETER Bitness
    Specify 32 vs 64 bit
.EXAMPLE 
    CheckDeploymentStatus -Branch FRCB1601 -Bitness 64
    Checks the status of the 64 bit version of the Office 365 update "FRCB1601"
#>  
Param(
    [Parameter(Mandatory=$true)]     
	[string]$Branch,
	
	[String]$Bitness = "32"
)

    Import-Module ((Split-Path $env:SMS_ADMIN_UI_PATH)+"\ConfigurationManager.psd1")
    $PSD = Get-PSDrive -PSProvider CMSite
    CD "$($PSD):"

    ## get the  Display Name of the update
    if(!$sun){
        Write-Host "Collecting the update display name from the catalog..."
        $ouidp = $Branch
        $bitness = $Bitness
        $sun = Get-CMSoftwareUpdate | where {$_.LocalizedDisplayName -like "*$ouidp*" -and $_.LocalizedDisplayName -like "*$bitness*"}
    }

    $officeupdatename = $sun.LocalizedDisplayName

    Get-CMDeployment -SoftwareName $officeupdatename
 }

Function Add-Content {
Param(   
	[string]$SiteCode,
	
	[String]$DistributionPoint,
 
	[String]$DeploymentPackageName = "O365 Clients - English",
    
	[String]$SiteServer = $env:COMPUTERNAME, 
   
	[array]$ContentIDs, 
  
	[array]$ContentSourcePath  
)

    if(!$SiteCode){
        $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
    }
    
    if(!$DistributionPoint){
        $DistributionPoint = (Get-CMDistributionPointInfo).Name
    }
                   
    $PackageID = (Get-WmiObject -Namespace root/SMS/site_$($SiteCode) -ComputerName $SiteServer -Query "SELECT * FROM SMS_SoftwareUpdatesPackage WHERE Name='$DeploymentPackageName'").PackageID
    #$PackageID

    $DeployPackage = (Get-WmiObject -Namespace root/SMS/site_$($SiteCode) -ComputerName $SiteServer -Query "SELECT * FROM SMS_SoftwareUpdatesPackage WHERE Name='$DeploymentPackageName'")

    $DeployPackage.AddUpdateContent($ContentIDs,$ContentSourcePath,$true)                                       
}

Function Add-DistributionPoint{
Param(   
	[string]$SiteCode,
	
	[String]$DistributionPoint,
  
	[String]$DeploymentPackageName,
    
	[String]$SiteServer
)

    Import-Module ((Split-Path $env:SMS_ADMIN_UI_PATH)+"\ConfigurationManager.psd1")
    $PSD = Get-PSDrive -PSProvider CMSite
    CD "$($PSD):"
        
    $PackageID = (Get-WmiObject -Namespace root/SMS/site_$($SiteCode) -ComputerName $SiteServer -Query "SELECT * FROM SMS_SoftwareUpdatesPackage WHERE Name='$DeploymentPackageName'").PackageID
            
    #echo "This is a Package" 
    Start-CMContentDistribution -DeploymentPackageId  $PackageID -DistributionPointName $DistributionPoint                         
}