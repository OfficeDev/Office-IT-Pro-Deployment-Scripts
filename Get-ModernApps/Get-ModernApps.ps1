<#
.SYNOPSIS
Gets a list of Modern Office Apps, versions, and the number of installs for each computer

.DESCRIPTION
Long Description

.PARAMETER ComputerFilter
The filter value that is fed to Get-ADComputer.

.PARAMETER OUFilter
The name of the OU you would like to limit your search to.

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
./Get-ModernAppsRemotely.ps1 -ComputerNames ($myArrayOfNames)
Gets the list of all the office modern apps that are installed on the specified computers.
You will be prompted for credentials.

.Example
./Get-ModernAppsRemotely.ps1 -OUFilter "OUName"
Gets the list of all the office modern apps that are installed on the computers in the specified OU in your domain.
You will be prompted for credentials.

.Notes
Additional explaination. Long and indepth examples should also go here.

.Link
Get-ADComputer

.Link
Get-AppxPackage

.Link
Invoke-Command


#>
[CmdletBinding(DefaultParameterSetName="Filter")]
Param
(

    [Parameter(ParameterSetName="Filter")]
    [string] $ComputerFilter = "*",

    [Parameter(ParameterSetName="Filter")]
    [string] $OUFilter = $null,
	
	[Parameter(ParameterSetName="Name",ValueFromPipeline=$true,Mandatory=$true)]
    [string[]] $ComputerNames,

    [Parameter()]
    [PSCredential] $Credentials

)

Process
{
    #Actual Functionality
    if($Credentials -eq $null)
    {
        $Credentials = Get-Credential;
    }

	$prevErrorActionPreference = $ErrorActionPreference;
	$ErrorActionPreference = 'Continue';
	switch($PSCmdlet.ParameterSetName)
	{
		"Filter" 
		{
			if([string]::IsNullOrWhiteSpace($OUFilter) -eq $false)
			{
				$appList = (Get-ADComputer -Filter $ComputerFilter | Select Name) | ? {$_.DistinguishedName -like "*ou=$OUFilter,*"}
					ForEach-Object {Invoke-Command -ScriptBlock {Get-AppxPackage -AllUsers | select Name, version | ? Name -like "Microsoft.Office.*" } -ComputerName $_.Name -Credential $Credentials } ;
			}
			else
			{
				$appList = (Get-ADComputer -Filter $ComputerFilter | Select Name) |
					ForEach-Object {Invoke-Command -ScriptBlock {Get-AppxPackage -AllUsers | select Name, version | ? Name -like "Microsoft.Office.*" } -ComputerName $_.Name -Credential $Credentials } ;
			}
		}
		"Name"
		{
            $ComputerNames
			$appList = $ComputerNames | ForEach-Object {Invoke-Command -ScriptBlock {Get-AppxPackage -AllUsers | select Name, version | ? Name -like "Microsoft.Office.*" } -ComputerName $_ -Credential $Credentials } ;
		}
	}
	$ErrorActionPreference = $prevErrorActionPreference;

	$names = $appList | Select Name -Unique;
	$computers = $appList | Select PSComputerName -Unique;
	$results = new-object PSObject[] 1;
	foreach($name in $names)
	{
		$versions = $appList | ? Name -Like $name.Name | Select version -Unique;
		foreach($version in $versions)
		{

			foreach($computer in $computers)
			{
				$count = $appList | ? Name -Like $name.Name | ? version -Like $version.version | ? PSComputerName -Like $computer.PSComputerName | measure;
				if($count.Count -gt 0)
				{
					if($results -eq $null)
					{
						$results[0] = New-Object PSObject -Property @{ Name=$name.Name; Version=$version.version; PSComputerName=$computer.PSComputerName; NumInstalls=$count.Count;};
					}
					else
					{
						$results += New-Object PSObject -Property @{ Name=$name.Name; Version=$version.version; PSComputerName=$computer.PSComputerName; NumInstalls=$count.Count;};
					}
				}#end if $count
			}#end foreach computer
		}#end foreach version
	}#end foreach name
	return $results;
}#end process