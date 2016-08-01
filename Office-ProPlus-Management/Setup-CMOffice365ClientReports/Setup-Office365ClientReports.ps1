function Edit-ConfigurationMofFile{
<#
.Synopsis
Edits the configuration.mof file.

.DESCRIPTION
Appends the configuration.mof file in the Configuration Manager installation directory with the content of configuration.txt.

.PARAMETER configMofTxt
The name of the configuration.txt file.

.EXAMPLE
Edit-ConfigurationMofFile
Appends the content inside of configuration.txt into the configuration.mof file.
#>
Param(
    [Parameter(Mandatory=$false)]
    $configMofTxt = "configuration.txt",

    [System.Management.Automation.PSCredential]$Credentials
)

Begin{
    $ConfigMofFileName = "configuration.mof"
    $configMofLocation = "inboxes\clifiles.src\hinv"
    
    
    $HKLM = [UInt32] "0x80000002"
    $installKey = 'SOFTWARE\Microsoft\SMS'

    if($Credentials) {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $env:COMPUTERNAME -Credential $Credentials
    } 
    else{
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $env:COMPUTERNAME
    }
}

Process{

    #Check uninstall path for ConfigMgr version   
    $ConfigMgrVersion = $regProv.EnumKey($HKLM, $installKey)
    foreach ($key in $ConfigMgrVersion.sNames) {
        if($key -match 'Setup'){
            $path = Join-Path $installKey $key
            $configItems = $regProv.EnumValues($HKLM, $path)
            foreach($item in $configItems.sNames){
                if($item -eq 'Installation Directory'){
                    $installationDirectory = $regProv.GetStringValue($HKLM, $path, $item).sValue
                }
            }
        }
    }

    $ConfigurationMofPath = Join-Path $installationDirectory $configMofLocation
    $configMofFilePath = $ConfigurationMofPath + "\" + $ConfigMofFileName
    $backupMofFilePath = $configMofFilePath + ".backup"
  
    #Make a backup of the mof file
    Copy-Item -Path $configMofFilePath -Destination $backupMofFilePath

    #Append the mof file
    Add-Content -Path $configMofFilePath -Value (Get-Content $configMofTxt)
}
}

function Import-CMReports{
<#
.SYNOPSIS
    Import all reports (.rdl files) in a specific folder to a Reporting Service point
.DESCRIPTION
    Use this script to import all the reports (.rdl files) in the specified source path folder to a Reporting Service point in ConfigMgr 2012
.PARAMETER ReportServer
    Site Server where SQL Server Reporting Services are installed
.PARAMETER SiteCode
    SiteCode of the Reporting Service point
.PARAMETER RootFolderName
    Should only be specified if the default 'ConfigMgr_<sitecode>' folder is not used and a custom folder was created
.PARAMETER FolderName
    If specified, search is restricted to within this folder if it exists
.PARAMETER SourcePath
    Path to where .rdl files eligible for import are located. If no SourcePath is specific the script will use the current location as the SourcePath.
.PARAMETER Credential
    PSCredential object created with Get-Credential or specify an username
.PARAMETER Force
    Will create a folder specified in the FolderName parameter if an existing folder is not present. Will be created in the root

.EXAMPLE
    .\Import-CMReports.ps1 -ReportServer CM01 -SiteCode PS1 -FolderName "Custom Reports" -SourcePath "C:\Import\RDL" -Force
    Import all the reports in 'C:\Import\RDL' to a folder called 'Custom Reports' on a report server called 'CM01'. 
    If the folder doesn't exist, it will be created in the root path:
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [parameter()]
    [string]$ReportServer = $env:COMPUTERNAME,

    [parameter()]
    [string]$SiteCode,

    [parameter()]
    [string]$RootFolderName = "ConfigMgr",

    [parameter()]
    [string]$FolderName = "Custom Reports",

    [parameter()]
    [string]$SourcePath,

    [Parameter()]
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [parameter()]
    [switch]$Force
)
Begin {

    if(!$SourcePath){
        $SourcePath = (Get-Location).Path + "\Reports"
    }

    #Build the Uri
    $SSRSUri = "http://$($ReportServer)/ReportServer/ReportService2010.asmx"
    #Get the site code if one is not specified
    if (!$SiteCode) {
       $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
    }
    # Build the default or custom ConfigMgr path for a Reporting Service point
    if ($RootFolderName -like "ConfigMgr") {
        $SSRSRootFolderName = -join ("/","$($RootFolderName)","_",$($SiteCode))
    }
    else {
        $SSRSRootFolderName = -join ("/","$($RootFolderName)")
    }
    #Build Server path
    if ($PSBoundParameters["FolderName"]) {
        $SSRSRootPath = -join ($SSRSRootFolderName,"/",$FolderName)
    }
    else {
        $SSRSRootPath = $SSRSRootFolderName
    }
    # Configure arguments being passed to the New-WebServiceProxy cmdlet by splatting
    $ProxyArgs = [ordered]@{
        "Uri" = $SSRSUri
        "UseDefaultCredential" = $true
    }
    if ($Credential -ne [System.Management.Automation.PSCredential]::Empty) {
        $ProxyArgs.Remove("UseDefaultCredential")
        $ProxyArgs.Add("Credential", $Credential)
    }
    else {
        Write-Verbose -Message "Credentials was not provided, using default"
    }
}
Process {
    try { 
        # Set up a WebServiceProxy
        $WebServiceProxy = New-WebServiceProxy @ProxyArgs -ErrorAction Stop
        if ($PSBoundParameters["FolderName"]) {
            Write-Verbose -Message "FolderName was specified"
            if ($WebServiceProxy.ListChildren($SSRSRootFolderName, $true) | Select-Object ID, Name, Path, TypeName | Where-Object { ($_.TypeName -eq "Folder") -and ($_.Name -like "$($FolderName)") }) {
                Create-Report -FilePath $SourcePath -ServerPath $SSRSRootPath
            }
            else {
                if(!($WebServiceProxy.ListChildren("/",$true) | Where-Object {$_.Path -like "$SSRSRootFolderName/$FolderName*"})){
                    if($PSCmdlet.ShouldProcess("Folder: $($FolderName)","Create")) {
                        Write-Host "Creating folder $($FolderName)..."
                        $TypeName = $WebServiceProxy.GetType().Namespace
                        $Property = New-Object -TypeName (-join ($TypeName,".Property"))
                        $Property.Name = "$($FolderName)"
                        $Property.Value = "$($FolderName)"
                        $Properties = New-Object -TypeName (-join ($TypeName,".Property","[]")) 1
                        $Properties[0] = $Property
                        $WebServiceProxy.CreateFolder($FolderName,"$($SSRSRootFolderName)",$Properties) | Out-Null 
                    }
                    Create-Report -FilePath $SourcePath -ServerPath $SSRSRootPath
                }
                else {
                    Write-Warning -Message "Unable to find a folder matching '$($FolderName)'"
                }
            }
        }
        else {
            Create-Report -FilePath $SourcePath -ServerPath $SSRSRootPath
        }
    }
    catch [Exception] {
        Throw $_.Exception.Message
    }
}
End{    
    Write-Progress -Activity "Importing Reports" -Completed -ErrorAction SilentlyContinue 
}
}

function Create-Report {
Param(
    [parameter(Mandatory=$true)]
    [string]$FilePath,
    [parameter(Mandatory=$true)]
    [string]$ServerPath
)
    
    $RDLFiles = Get-ChildItem -Path $FilePath -Filter "*.rdl"
    $RDLFilesCount = ($RDLFiles | Measure-Object).Count
    if(($RDLFiles | Measure-Object).Count -ge 1) {
        foreach ($RDLFile in $RDLFiles) {               
            $ProgressCount++
            Write-Progress -Activity "Importing Reports" -Id 1 -Status "$($ProgressCount) / $($RDLFilesCount)" -CurrentOperation "$($RDLFile.Name)" -PercentComplete (($ProgressCount / $RDLFilesCount) * 100)
            if($PSCmdlet.ShouldProcess("Report: $($RDLFile.BaseName)","Validate")) {
                $ValidateReportName = $WebServiceProxy.ListChildren($SSRSRootPath, $true) | Where-Object { ($_.TypeName -like "Report") -and ($_.Name -like "$($RDLFile.BaseName)") }
            }
            if($ValidateReportName -eq $null) {
                if($PSCmdlet.ShouldProcess("Report: $($RDLFile.BaseName)","Create")) {
                    $RDLFileName = [System.IO.Path]::GetFileNameWithoutExtension($RDLFile.Name)
                    $ByteStream = Get-Content -Path $RDLFile.FullName -Encoding Byte
                    $Warnings = @()
                    Write-Verbose -Message "Importing report '$($RDLFileName)'"
                    $WebServiceProxy.CreateCatalogItem("Report",$RDLFileName,$SSRSRootPath,$true,$ByteStream,$null,[ref]$Warnings) | Out-Null
                }
                #Get name of default ConfigMgr data source
                $DefaultCMDataSource = $WebServiceProxy.ListChildren($SSRSRootFolderName, $true) | Where-Object { $_.TypeName -like "DataSource" } | Select-Object -First 1
                if($DefaultCMDataSource -ne $null) {
                    if($PSCmdlet.ShouldProcess("DataSource: $($DefaultCMDataSource.Name)","Set")) {
                        $CurrentReport = $WebServiceProxy.ListChildren($SSRSRootFolderName, $true) | Where-Object { ($_.TypeName -like "Report") -and ($_.Name -like "$($RDLFileName)") -and ($_.CreationDate -ge (Get-Date).AddMinutes(-5)) }
                        $CurrentReportDataSource = $WebServiceProxy.GetItemDataSources($CurrentReport.Path)
                        $DataSourceType = $WebServiceProxy.GetType().Namespace
                        $ArrayItems = 1 # Means how many objects should be in the array
                        $DataSourceArray = New-Object -TypeName (-join ($DataSourceType,".DataSource","[]")) $ArrayItems
                        $DataSourceArray[0] = New-Object -TypeName (-join ($DataSourceType,".DataSource"))
                        $DataSourceArray[0].Name = $CurrentReportDataSource.Name
                        $DataSourceArray[0].Item = New-Object -TypeName (-join ($DataSourceType,".DataSourceReference"))
                        $DataSourceArray[0].Item.Reference = $DefaultCMDataSource.Path
                        Write-Verbose -Message "Changing data source for report '$($RDLFileName)'"
                        $WebServiceProxy.SetItemDataSources($CurrentReport.Path, $DataSourceArray)
                    }
                }
                else{
                    Write-Warning -Message "Unable to determine default ConfigMgr data source, will not edit data source for report '$($RDLFileName)'"
                }
            }
            else{
                Write-Warning -Message "A report with the name '$($RDLFile.BaseName)' already exists, skipping import"
            }
        }
    }
    else{
        Write-Warning -Message "No .rdl files was found in the specified path"
    }
}
