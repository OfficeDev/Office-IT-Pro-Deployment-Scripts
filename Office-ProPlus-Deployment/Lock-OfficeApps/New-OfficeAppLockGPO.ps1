Function New-OfficeAppLockGPO{
<#
.SYNOPSIS
Creates a group policy that will lock the specified version of Office.

.DESCRIPTION
This function will create a GPO that will block Office applications older than 2016
from opening by providing the version/s of Office. If the GPO name and WMI filter name are not
provided a default name will be created. The WMI filter will be automatically linked
to the corresponding GPO.

.PARAMETER GpoName
Name of the new GPO.

.PARAMETER OfficeVersion
The version of Office to block.

.EXAMPLE
New-OfficeAppLockGPO -GpoName "Lock Office 2010,2013" -OfficeVersion Office2010,Office2013
A GPO and WMi filter called "Lock Office 2010,2013" will be created and linked. When applied to
the appropriate OU any computer with Office 2010 and 2016 or 2013 and 2016 will be prevented
from opening the version older than 2016.

.NOTE
This function will dot source the Manage-OfficeWmiFilters.ps1 script. Be sure both scripts
are saved in the same folder.
#>

    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $GpoName = $null,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $WmiFilterName = $GpoName,

        [ValidateSet("Office2003", "Office2007", "Office2010","Office2013")]
        [string[]] $OfficeVersion
    )

    Import-Module -Name grouppolicy

    . .\Manage-OfficeWmiFilters.ps1

    $dateconv = Get-Date -Format G
    $date = (Get-date $dateconv).TofileTime()
   
    if(!($GpoName)){
    
        $GpoName = @("LockOffice2003","LockOffice2007","LockOffice2010","LockOffice2013")   
        $officeNumbers = @("11","12","14","15")
        $gpoCounter = 0

        foreach($Gpo in $GpoName){

            New-GPO -Name $Gpo

            $appStrings = @("C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\WINWORD.EXE",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\WINWORD.EXE",
                            "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\OUTLOOK.EXE",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\OUTLOOK.EXE",
                            "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\ONENOTE.EXE",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\ONENOTE.EXE",
                            "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\POWERPNT.EXE",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\POWERPNT.EXE",
                            "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\MSACCESS.EXE",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\MSACCESS.EXE",
                            "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\EXCEL.EXE",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\EXCEL.EXE"
                            "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\INFOPATH.EXE",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\INFOPATH.EXE"
                            "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\WINPROJ.EXE",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\WINPROJ.EXE"
                            "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\MSPUB.EXE",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\MSPUB.EXE"
                            "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\SPDESIGN.EXE",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\SPDESIGN.EXE"
                            "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\GROOVE.EXE",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\GROOVE.EXE"
                            "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\VISLIB.DLL",
                            "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\VISLIB.DLL")

            $appLocations = @("%HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRoot%",
                              "%HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir%")

            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\Certificates" -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\CRLs" -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\CTLs" -Type String -Value ""
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\Certificates" -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\CRLs" -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\CTLs" -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName DefaultLevel -Type DWord -Value 262144 | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName ExecutableTypes -Type MultiString -Value "ADE\0ADP\0BAS\0BAT\0CHM\0CMD\0COM\0CPL\0CRT\0EXE\0HLP\0HTA\0INF\0INS\0ISP\0LNK\0MDB\0MDE\0MSC\0MSI\0MSP\0MST\0OCX\0PCD\0PIF\0REG\0SCR\0SHS\0URL\0VB\0WSC" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName PolicyScope -Type DWord -Value 0 | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName TransparentEnabled -Type DWord -Value 1 | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName Description -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName LastModified -Type QWord -Value $date | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null

            foreach($app in $appStrings)
            {
                $guid = ([system.guid]::NewGuid())
                $guidString = "{$($guid.ToString())}"

                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName LastModified -Type QWord -Value $date | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName Description -Type String -Value "" | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName ItemData -Type ExpandString -Value $app | Out-Null   
            }

            foreach($loc in $appLocations)
            {
                $guid = ([system.guid]::NewGuid())
                $guidString = "{$($guid.ToString())}"
    
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName Description -Type String -Value "" | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName ItemData -Type ExpandString -Value "{$($loc.ToString())}" | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName LastModified -Type QWord -Value $date | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null
            }

            Write-Host "The Group Policy $Gpo has been created"

            $gpoCounter = $gpoCounter + 1
        }
    }
    
    else{
    
        New-GPO -Name $GpoName

        $appLocations = @("%HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRoot%",
                          "%HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir%")

        foreach($loc in $appLocations)
        {
            $guid = ([system.guid]::NewGuid())
            $guidString = "{$($guid.ToString())}"
    
            Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName Description -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName ItemData -Type ExpandString -Value "{$($loc.ToString())}" | Out-Null
            Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName LastModified -Type QWord -Value $date | Out-Null
            Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null       
        }

        if($OfficeVersion -eq "Office2003"){

            $officeNumber = '11'
            SetGpoPolValues
        }

        if($OfficeVersion -eq "Office2007"){

            $officeNumber = '12'
            SetGpoPolValues
        }

        if($OfficeVersion -eq "Office2010"){

            $officeNumber = '14'
            SetGpoPolValues
        }
                    
        if($OfficeVersion -eq "Office2013"){

            $officeNumber = '15'
            SetGpoPolValues
        }             
    }

        if($OfficeVersion -contains "Office2003")        
        {
            if($OfficeVersion -contains "Office2007")
            {
                if($OfficeVersion -contains "Office2010")
                {
                    if($OfficeVersion -contains "Office2013")
                    {
                        $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "11.0%" OR Version LIKE "12.%" OR Version LIKE "14.0%" OR Version LIKE "15.0%"'
                    }
                    else
                    {
                        $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "11.0%" OR Version LIKE "12.%" OR Version LIKE "14.0%"'
                    }
                }      
                else
                {
                    $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "11.0%" OR Version LIKE "12.%"'
                }
            }
            elseif($OfficeVersion -contains "Office2010")
            {
                if($OfficeVersion -contains "Office2013")
                {
                    $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "11.0%" OR Version LIKE "14.%" OR Version LIKE "15.0%"'
                }
                else
                {
                    $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "11.0%" OR Version LIKE "14.%"'
                }
            }
            elseif($OfficeVersion -contains "Office2013")
            {
                $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "11.0%" OR Version LIKE "15.0%"'
            }
            else
            {
                $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "11.0%"'
            }
        }
        
        elseif($OfficeVersion -contains "Office2007")
        {
            if($OfficeVersion -contains "Office2010")
            {
                if($OfficeVersion -contains "Office2013")
                {
                    $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "12.0%" OR Version LIKE "14.%" OR Version LIKE "15.0%"'
                }
                else
                {
                    $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "12.0%" OR Version LIKE "14.%"'
                }
            }
            elseif($OfficeVersion -contains "Office2013")
            {
                $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "12.0%" OR Version LIKE "15.0%"'
            }
            else
            {
                $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "12.0%"'
            }
        }
        elseif($OfficeVersion -contains "Office2010")
        {
            if($OfficeVersion -contains "Office2013")
            {
                $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "14.%" OR Version LIKE "15.%"'
            }
            else
            {
                $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "14.%"'
            }
        }            
        elseif($OfficeVersion -eq "Office2013")
        {         
            $WqlQuery = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Microsoft Office%" AND Version LIKE "15.%"'
        }               


    $Wql2016Query = 'SELECT * FROM Win32_Product WHERE Caption LIKE "Office 16%"'        
    [string[]]$object = $WqlQuery,$Wql2016Query

    [string[]]$Expression = $object
       
    . .\Manage-OfficeWmiFilters.ps1

    New-GPWmiFilter -WmiFilterName $WmiFilterName -Expression $Expression
    Add-GPWmiLink -WmiFilterName $WmiFilterName -GpoName $GpoName

    $results = new-object PSObject[] 0;
    $Result = New-Object –TypeName PSObject
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "GpoName" -Value $GpoName
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "WmiFilterName" -Value $WmiFilterName    
    $result
}
    
Function SetGpoPolValues{
               
        $appStrings = @("C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\WINWORD.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\WINWORD.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\OUTLOOK.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\OUTLOOK.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\ONENOTE.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\ONENOTE.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\POWERPNT.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\POWERPNT.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\MSACCESS.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\MSACCESS.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\EXCEL.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\EXCEL.EXE")

        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\Certificates" -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\CRLs" -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\CTLs" -Type String -Value ""
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\Certificates" -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\CRLs" -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\CTLs" -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName DefaultLevel -Type DWord -Value 262144 | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName ExecutableTypes -Type MultiString -Value "ADE\0ADP\0BAS\0BAT\0CHM\0CMD\0COM\0CPL\0CRT\0EXE\0HLP\0HTA\0INF\0INS\0ISP\0LNK\0MDB\0MDE\0MSC\0MSI\0MSP\0MST\0OCX\0PCD\0PIF\0REG\0SCR\0SHS\0URL\0VB\0WSC" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName PolicyScope -Type DWord -Value 0 | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName TransparentEnabled -Type DWord -Value 1 | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName Description -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName LastModified -Type QWord -Value $date | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null

        foreach($app in $appStrings)
        {
           $guid = ([system.guid]::NewGuid())
           $guidString = "{$($guid.ToString())}"

           Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName LastModified -Type QWord -Value $date | Out-Null
           Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName Description -Type String -Value "" | Out-Null
           Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null
           Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName ItemData -Type ExpandString -Value $app | Out-Null   
        }      
} 