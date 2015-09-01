[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
    [string[]] $ComputerName = $env:COMPUTERNAME,
    [PSCredential] $Credentials
)


function Get-ModernOfficeApps {
<#
.SYNOPSIS
Gets a list of Modern Office Apps, versions, and the number of installs for each computer

.DESCRIPTION
This script will query the local or remote computers and detect if Modern Office Apps are installed. Since
the Modern apps on installed on a per user basis the script will also include the numbers of users
who have an instance of the application installed

.PARAMETER ComputerNames
The list of computer names that you would like limit your search to.

.PARAMETER credentials
The PSCredentials that are used to invoke commands on the remote computers. 
Will be prompted for if not provided.

.Example
./Get-ModernAppsRemotely.ps1
Gets the list of all the office modern apps that are installed on all computers in your domain.
You will be prompted for credentials.

.Example
./Get-ModernAppsRemotely.ps1 -ComputerNames ($myArray)
Gets the list of all the office modern apps that are installed on the specified computers.
You will be prompted for credentials.

.Outputs
Application Name, Version, Number of users who have the application installed, ComputerName

#>
[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
    [string[]] $ComputerName = $env:COMPUTERNAME,
    [PSCredential] $Credentials
)

begin {
    $defaultDisplaySet = 'Name','Version', 'NumberOfInstalls', 'ComputerName';

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet);
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet);
}

Process
{
    $HKLM = [UInt32] "0x80000002";
    $HKCR = [UInt32] "0x80000000";
    $HKEY_Users = 2147483651;

	$results = new-object PSObject[] 1;

	foreach($computer in $ComputerName)
	{

        #Actual Functionality
        if ($Credentials) {
           $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials;
        } else {
           $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer;
        }

        foreach ($userKey in $regProv.EnumKey($HKEY_Users,"").sNames) {

          $regpath = "Software\Classes\ActivatableClasses\Package";
          $packagePath = join-path $userKey $regpath;

           foreach ($packageKey in $regProv.EnumKey($HKEY_Users,$packagePath).sNames) {
               $packageName = $packageKey.Split('_')[0];
               $packageVersion = $packageKey.Split('_')[1];

               if (!$packageName.ToLower().StartsWith("Microsoft.Office".ToLower())) {
                  continue;
               }
        
                $exists = $false;
                foreach ($result in $results) {
                  if ($result.Name) {
                     if ($result.Name.ToUpper() -eq $packageName.ToUpper() -and $result.Version -eq $packageVersion -and $result.ComputerName -eq $computer) {
                         $exists = $true;
                         $result.NumInstalls += 1;
                     }
                  }
                }

                if (!$exists) {
                   $object = New-Object PSObject -Property @{ Name=$packageName; Version=$packageVersion; ComputerName=$computer; NumberOfInstalls=1;};
                   $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers;
                   $results += $object;
                }

           }

        }

	}

	return $results;
}
}

Get-ModernOfficeApps -ComputerName $ComputerName -Credentials $Credentials;