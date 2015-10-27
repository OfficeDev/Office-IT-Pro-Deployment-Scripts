function New-GPWmiFilter {
<#
.SYNOPSIS
Create a new WMI filter for Group Policy with given name, WQL query and description.

.DESCRIPTION
The New-GPWmiFilter function create an AD object for WMI filter with specific name, WQL query expressions and description.
With -PassThru switch, it output the WMIFilter instance which can be assigned to GPO.WMIFilter property.

.PARAMETER Name
The name of new WMI filter.

.PARAMETER Expression
The expression(s) of WQL query in new WMI filter. Pass an array to this parameter if multiple WQL queries applied.

.PARAMETER Description
The description text of the WMI filter (optional). 

.PARAMETER PassThru
Output the new WMI filter instance with this switch.

.EXAMPLE
New-GPWmiFilter -Name 'Virtual Machines' -Expression 'SELECT * FROM Win32_ComputerSystem WHERE Model = "Virtual Machine"' -Description 'Only apply on virtual machines'

Create a WMI filter to apply GPO only on virtual machines

.EXAMPLE 
$filter = New-GPWmiFilter -Name 'Workstation 32-bit' -Expression 'SELECT * FROM WIN32_OperatingSystem WHERE ProductType=1', 'SELECT * FROM Win32_Processor WHERE AddressWidth = "32"' -PassThru
$gpo = New-GPO -Name "Test GPO"
$gpo.WmiFilter = $filter

Create a WMI filter for 32-bit work station and link it to a new GPO named "Test GPO".

.NOTES
Domain administrator priviledge is required for executing this cmdlet

#>
   [CmdletBinding()] 
    Param
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNull()]
        [string] $WmiFilterName = $GpoName,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNull()]
        [string] $GpoName,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNull()]
        [string[]] $Expression,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Description,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [switch] $PassThru
    )


    if ($Expression.Count -lt 1)
    {
        Write-Error "At least one Expression Method is required to create a WMI Filter."
        return
    }

    $Guid = [System.Guid]::NewGuid()
    $defaultNamingContext = (Get-ADRootDSE).DefaultNamingContext 
    $msWMIAuthor = Get-Author
    $msWMICreationDate = (Get-Date).ToUniversalTime().ToString("yyyyMMddhhmmss.ffffff-000")
    $WMIGUID = "{$Guid}"
    $WMIDistinguishedName = "CN=$WMIGUID,CN=SOM,CN=WMIPolicy,CN=System,$defaultNamingContext"
    $msWMIParm1 = "$Description "
    $msWMIParm2 = $Expression.Count.ToString() + ";"
    $Expression | ForEach-Object {
        $msWMIParm2 += "3;10;" + $_.Length + ";WQL;root\CIMv2;" + $_ + ";"
    }

    $Attr = @{
        "msWMI-Name" = $WmiFilterName;
        "msWMI-Parm1" = $msWMIParm1;
        "msWMI-Parm2" = $msWMIParm2;
        "msWMI-Author" = $msWMIAuthor;
        "msWMI-ID"= $WMIGUID;
        "instanceType" = 4;
        "showInAdvancedViewOnly" = "TRUE";
        "distinguishedname" = $WMIDistinguishedName;
        "msWMI-ChangeDate" = $msWMICreationDate; 
        "msWMI-CreationDate" = $msWMICreationDate
    }
    
    $WMIPath = ("CN=SOM,CN=WMIPolicy,CN=System,$defaultNamingContext")

    Enable-ADSystemOnlyChange

    $ADObject = New-ADObject -Name $WMIGUID -Type "msWMI-Som" -Path $WMIPath -OtherAttributes $Attr -PassThru

    if ($PassThru)
    {
        ConvertTo-WmiFilter $ADObject | Write-Output
    }
}

function Add-GPWmiLink {

    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $WmiFilterName = $null,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $GpoName = $null
    )

    $officeLockGpo = Get-GPO -Name $GpoName
    $wmiFilterLink = Get-GPWmiFilter -WmiFilterName $WmiFilterName
    $wmiLink = $officeLockGpo.WmiFilter = $wmiFilterLink

    return $wmiLink
}

function Get-GPWmiFilter {
<#
.SYNOPSIS
Get a WMI filter in current domain

.DESCRIPTION
The Get-GPWmiFilter function query WMI filter(s) in current domain with specific name or GUID.

.PARAMETER Guid
The guid of WMI filter you want to query out.

.PARAMETER Name
The name of WMI filter you want to query out.

.PARAMETER All
Query all WMI filters in current domain With this switch.

.EXAMPLE
Get-GPWmiFilter -Name 'Virtual Machines'

Get WMI filter(s) with the name 'Virtual Machines'

.EXAMPLE 
Get-GPWmiFilter -All

Get all WMI filters in current domain

#>
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="ByGUID")]
        [ValidateNotNull()]
        [Guid[]] $Guid,

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="ByName")]
        [ValidateNotNull()]
        [string[]] $WmiFilterName,
        
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="GetAll")]
        [ValidateNotNull()]
        [switch] $All
    )
    if ($Guid)
    {
        $ADObject = Get-WMIFilterInADObject -Guid $Guid
    }
    elseif ($WmiFilterName)
    {
        $ADObject = Get-WMIFilterInADObject -WmiFilterName $WmiFilterName
    }
    elseif ($All)
    {
        $ADObject = Get-WMIFilterInADObject -All
    }
    ConvertTo-WmiFilter $ADObject | Write-Output
}

function Remove-GPWmiFilter {
<#
.SYNOPSIS
Remove a WMI filter from current domain

.DESCRIPTION
The Remove-GPWmiFilter function remove WMI filter(s) in current domain with specific name or GUID.

.PARAMETER Guid
The guid of WMI filter you want to remove.

.PARAMETER Name
The name of WMI filter you want to remove.

.EXAMPLE
Remove-GPWmiFilter -Name 'Virtual Machines'

Remove the WMI filter with name 'Virtual Machines'

.NOTES
Domain administrator priviledge is required for executing this cmdlet

#>
   [CmdletBinding(DefaultParametersetName="ByGUID")] 
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="ByGUID")]
        [ValidateNotNull()]
        [Guid[]] $Guid,

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="ByName")]
        [ValidateNotNull()]
        [string[]] $WmiFilterName
    )
    if ($Guid)
    {
        $ADObject = Get-WMIFilterInADObject -Guid $Guid
    }
    elseif ($WmiFilterName)
    {
        $ADObject = Get-WMIFilterInADObject -WmiFilterName $WmiFilterName
    }
    $ADObject | ForEach-Object  {
        if ($_.DistinguishedName)
        {
            Remove-ADObject $_ -Confirm:$false
        }
    }
}

function Set-GPWmiFilter {
<#
.SYNOPSIS
Get a WMI filter in current domain and update the content of it

.DESCRIPTION
The Set-GPWmiFilter function query WMI filter(s) in current domain with specific name or GUID and then update the content of it.

.PARAMETER Guid
The guid of WMI filter you want to query out.

.PARAMETER Name
The name of WMI filter you want to query out.

.PARAMETER Expression
The expression(s) of WQL query in new WMI filter. Pass an array to this parameter if multiple WQL queries applied.

.PARAMETER Description
The description text of the WMI filter (optional). 

.PARAMETER PassThru
Output the updated WMI filter instance with this switch.

.EXAMPLE
Set-GPWmiFilter -Name 'Workstations' -Expression 'SELECT * FROM Win32_OperatingSystem WHERE ProductType = "1"'

Set WMI filter named with "Workstations" to specific WQL query

.NOTES
Domain administrator priviledge is required for executing this cmdlet.
Either -Expression or -Description should be assigned when executing.

#>
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="ByGUID")]
        [ValidateNotNull()]
        [Guid[]] $Guid,

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="ByName")]
        [ValidateNotNull()]
        [string[]] $WmiFilterName,
        
        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, Position=1)]
        [ValidateNotNull()]
        [string[]] $Expression,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, Position=2)]
        [string] $Description,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, Position=3)]
        [switch] $PassThru

    )
    if ($Guid)
    {
        $ADObject = Get-WMIFilterInADObject -Guid $Guid
    }
    elseif ($WmiFilterName)
    {
        $ADObject = Get-WMIFilterInADObject -WmiFilterName $WmiFilterName
    }
    $msWMIAuthor = Get-Author
    $msWMIChangeDate = (Get-Date).ToUniversalTime().ToString("yyyyMMddhhmmss.ffffff-000")
    $Attr = @{
        "msWMI-Author" = $msWMIAuthor;
        "msWMI-ChangeDate" = $msWMIChangeDate;
    }
    if ($Expression)
    {
        $msWMIParm2 = $Expression.Count.ToString() + ";"
        $Expression | ForEach-Object {
            $msWMIParm2 += "3;10;" + $_.Length + ";WQL;root\CIMv2;" + $_ + ";"
        }
        $Attr.Add("msWMI-Parm2", $msWMIParm2);
    }
    elseif ($Description)
    {
        $msWMIParm1 = $Description + " "
        $Attr.Add("msWMI-Parm2", $msWMIParm2);
    }
    else
    {
        Write-Warning "No content need to be set. Please set either Expression or Description."
        return
    }

    Enable-ADSystemOnlyChange

    $ADObject | ForEach-Object  {
        if ($_.DistinguishedName)
        {
            Set-ADObject -Identity $_ -Replace $Attr
            if ($PassThru)
            {
                ConvertTo-WmiFilter $ADObject | Write-Output
            }
        }
    }
}

function Rename-GPWmiFilter {
<#
.SYNOPSIS
Get a WMI filter in current domain and rename it

.DESCRIPTION
The Rename-GPWmiFilter function query WMI filter in current domain with specific name or GUID and then change it to a new name.

.PARAMETER Guid
The guid of WMI filter you want to query out.

.PARAMETER Name
The name of WMI filter you want to query out.

.PARAMETER TargetName
The new name of WMI filter.

.PARAMETER PassThru
Output the renamed WMI filter instance with this switch.

.EXAMPLE
Rename-GPWmiFilter -Name 'Workstations' -TargetName 'Client Machines'

Rename WMI filter "Workstations" to "Client Machines"

.NOTES
Domain administrator priviledge is required for executing this cmdlet.

#>
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="ByGUID")]
        [ValidateNotNull()]
        [Guid[]] $Guid,

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="ByName")]
        [ValidateNotNull()]
        [string[]] $WmiFilterName,
        
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=1)]
        [ValidateNotNull()]
        [string] $TargetName,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, Position=3)]
        [switch] $PassThru
    )
    if ($Guid)
    {
        $ADObject = Get-WMIFilterInADObject -Guid $Guid
    }
    elseif ($Name)
    {
        $ADObject = Get-WMIFilterInADObject -WmiFilterName $WmiFilterName
    }

    if (!$WmiFilterName)
    {
        $WmiFilterName = $ADObject."msWMI-Name"
    }
    if ($TargetName -eq $WmiFilterName)
    {
        return
    }

    $msWMIAuthor = Get-Author
    $msWMIChangeDate = (Get-Date).ToUniversalTime().ToString("yyyyMMddhhmmss.ffffff-000")
    $Attr = @{
        "msWMI-Author" = $msWMIAuthor;
        "msWMI-ChangeDate" = $msWMIChangeDate; 
        "msWMI-Name" = $TargetName;
    }

    Enable-ADSystemOnlyChange

    $ADObject | ForEach-Object  {
        if ($_.DistinguishedName)
        {
            Set-ADObject -Identity $_ -Replace $Attr
            if ($PassThru)
            {
                ConvertTo-WmiFilter $ADObject | Write-Output
            }
        }
    }
}

function ConvertTo-WmiFilter([Microsoft.ActiveDirectory.Management.ADObject[]] $ADObject){
    $gpDomain = New-Object -Type Microsoft.GroupPolicy.GPDomain
    $ADObject | ForEach-Object {
        $path = 'MSFT_SomFilter.Domain="' + $gpDomain.DomainName + '",ID="' + $_.Name + '"'
        try 
        {
            $filter = $gpDomain.GetWmiFilter($path)
        }
        catch { }
        if ($filter)
        {
            [Guid]$Guid = $_.Name.Substring(1, $_.Name.Length - 2)
            $filter | Add-Member -MemberType NoteProperty -Name Guid -Value $Guid -PassThru | Add-Member -MemberType NoteProperty -Name Content -Value $_."msWMI-Parm2" -PassThru | Write-Output
        }
    }
}

function ConvertTo-ADObject([Microsoft.GroupPolicy.WmiFilter[]] $WmiFilter){
    $wmiFilterAttr = "msWMI-Name", "msWMI-Parm1", "msWMI-Parm2", "msWMI-Author", "msWMI-ID"
    $WmiFilter | ForEach-Object {
        $match = $_.Path | Select-String -Pattern 'ID=\"\{(?<id>[\-|a-f|0-9]+)\}\"' | Select-Object -Expand Matches | ForEach-Object { $_.Groups[1] }
        [Guid]$Guid = $match.Value
        $ldapFilter = "(&(objectClass=msWMI-Som)(Name={$Guid}))"
        Get-ADObject -LDAPFilter $ldapFilter -Properties $wmiFilterAttr | Write-Output
    }
}

function Enable-ADSystemOnlyChange([switch] $disable){
    $valueData = 1
    if ($disable)
    {
        $valueData = 0
    }
    $key = Get-Item HKLM:\System\CurrentControlSet\Services\NTDS\Parameters -ErrorAction SilentlyContinue
    if (!$key) {
        New-Item HKLM:\System\CurrentControlSet\Services\NTDS\Parameters -ItemType RegistryKey | Out-Null
    }
    $kval = Get-ItemProperty HKLM:\System\CurrentControlSet\Services\NTDS\Parameters -Name "Allow System Only Change" -ErrorAction SilentlyContinue
    if (!$kval) {
        New-ItemProperty HKLM:\System\CurrentControlSet\Services\NTDS\Parameters -Name "Allow System Only Change" -Value $valueData -PropertyType DWORD | Out-Null
    } else {
        Set-ItemProperty HKLM:\System\CurrentControlSet\Services\NTDS\Parameters -Name "Allow System Only Change" -Value $valueData | Out-Null
    }
}

function Get-WMIFilterInADObject {
    Param(
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="ByGUID")]
        [ValidateNotNull()]
        [Guid[]] $Guid,

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="ByName")]
        [ValidateNotNull()]
        [string[]] $WmiFilterName,
        
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0, ParameterSetName="GetAll")]
        [ValidateNotNull()]
        [switch] $All

    )
    $wmiFilterAttr = "msWMI-Name", "msWMI-Parm1", "msWMI-Parm2", "msWMI-Author", "msWMI-ID"
    if ($Guid)
    {
        $Guid | ForEach-Object {
            $ldapFilter = "(&(objectClass=msWMI-Som)(Name={$_}))"
            Get-ADObject -LDAPFilter $ldapFilter -Properties $wmiFilterAttr | Write-Output
        }
    }
    elseif ($WmiFilterName)
    {
        $WmiFilterName | ForEach-Object {
            $ldapFilter = "(&(msWMI-Name=$_))"
            Get-ADObject -LDAPFilter $ldapFilter -Properties $wmiFilterAttr | Write-Output
        }
    }
    elseif ($All)
    {
        $ldapFilter = "(objectClass=msWMI-Som)"
        Get-ADObject -LDAPFilter $ldapFilter -Properties $wmiFilterAttr | Write-Output
    }
}

function Get-Author{
    $author = (Get-ADUser $env:USERNAME).UserPrincipalName
    if (!$author)
    {
        $author = (Get-ADUser $env:USERNAME).Name
    }
    if (!$author)
    {
        $author = $env:USERNAME
    }
    return $author
}