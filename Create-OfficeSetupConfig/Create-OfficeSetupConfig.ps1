<#
.SYNOPSIS
Short Description

.DESCRIPTION
Long Description

.PARAMETER
Explaination of myParam1

.PARAMETER
Explaination of myParam2

.Example
./Skeleton.ps1 -myParam1 "Value1" -myParam2 "Value2"
Usage example one

.Example
./Skeleton.ps1 -Param1 "Value1"
Usage example two

.Notes
Additional explaination. Long and indepth examples should also go here.

.Link
http://relevantlink.com

.Link
relevent-command

#>

[CmdletBinding()]
Param(

    [Parameter()]
    [string] $ProductId,

#    [Parameter()]
#    [Hashtable] $ARPOptions,

#    [Parameter()]
#    [Hashtable] $CommandOptions,

    [Parameter()]
    [string] $CompanyName,

#    [Parameter()]
#    [Hashtable] $DisplayOptions,

#    [Parameter()]
#    [string] $DistributionPointPath,

#    [Parameter()]
#    [string] $InstallLocation,

#    [Parameter()]
#    [Hashtable] $LISOptions,

#    [Parameter()]
#    [Hashtable] $LoggingOptions,

    [Parameter()]
    [Hashtable[]] $OptionStateList,

    [Parameter()]
    [string] $PIDKEY,

    [Parameter()]
    [

)