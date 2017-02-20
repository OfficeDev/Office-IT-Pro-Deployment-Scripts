

# Functions stolen from interwebs and adjusted for Office Automation purposes
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Get-SCCMCommands {
    # List all SCCM-commands
    [CmdletBinding()]
    PARAM ()
    PROCESS {
        return Get-Command -Name *-SCCM* -CommandType Function  | Sort-Object Name | Format-Table Name, Module
    }
}
 
Function Connect-SCCMServer {
    # Connect to one SCCM server
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$false,HelpMessage="SCCM Server Name or FQDN",ValueFromPipeline=$true)][Alias("ServerName","FQDN","ComputerName")][String] $HostName = (Get-Content env:computername),
        [Parameter(Mandatory=$false,HelpMessage="Optional SCCM Site Code",ValueFromPipelineByPropertyName=$true )][String] $siteCode = $null,
        [Parameter(Mandatory=$false,HelpMessage="Credentials to use" )][System.Management.Automation.PSCredential] $credential = $null
    )
 
    PROCESS {
        # Get the pointer to the provider for the site code
        if ($siteCode -eq $null -or $siteCode -eq "") {
            Write-Verbose "Getting provider location for default site on server $HostName"
            if ($credential -eq $null) {
                $sccmProviderLocation = Get-WmiObject -query "select * from SMS_ProviderLocation where ProviderForLocalSite = true" -Namespace "root\sms" -computername $HostName -errorAction Stop
            } else {
                $sccmProviderLocation = Get-WmiObject -query "select * from SMS_ProviderLocation where ProviderForLocalSite = true" -Namespace "root\sms" -computername $HostName -credential $credential -errorAction Stop
            }
        } else {
            Write-Verbose "Getting provider location for site $siteCode on server $HostName"
            if ($credential -eq $null) {
                $sccmProviderLocation = Get-WmiObject -query "SELECT * FROM SMS_ProviderLocation where SiteCode = '$siteCode'" -Namespace "root\sms" -computername $HostName -errorAction Stop
            } else {
                $sccmProviderLocation = Get-WmiObject -query "SELECT * FROM SMS_ProviderLocation where SiteCode = '$siteCode'" -Namespace "root\sms" -computername $HostName -credential $credential -errorAction Stop
            }
        }
 
        # Split up the namespace path
        $parts = $sccmProviderLocation.NamespacePath -split "\\", 4
        Write-Verbose "Provider is located on $($sccmProviderLocation.Machine) in namespace $($parts[3])"
 
        # Create a new object with information
        $retObj = New-Object -TypeName System.Object
        $retObj | add-Member -memberType NoteProperty -name Machine -Value $HostName
        $retObj | add-Member -memberType NoteProperty -name Namespace -Value $parts[3]
        $retObj | add-Member -memberType NoteProperty -name SccmProvider -Value $sccmProviderLocation
 
        return $retObj
    }
}
 
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Get-SCCMObject {
    #  Generic query tool
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipelineByPropertyName=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true, HelpMessage="SCCM Class to query",ValueFromPipeline=$true)][Alias("Table","View")][String] $class,
        [Parameter(Mandatory=$false,HelpMessage="Optional Filter on query")][String] $Filter = $null
    )
 
    PROCESS {
        if ($Filter -eq $null -or $Filter -eq "")
        {
            Write-Verbose "WMI Query: SELECT * FROM $class"
            $retObj = get-wmiobject -class $class -computername $SccmServer.Machine -namespace $SccmServer.Namespace
        }
        else
        {
            Write-Verbose "WMI Query: SELECT * FROM $class WHERE $Filter"
            $retObj = get-wmiobject -query "SELECT * FROM $class WHERE $Filter" -computername $SccmServer.Machine -namespace $SccmServer.Namespace
        }
 
        return $retObj
    }
}
 
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Get-SCCMPackage {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Optional Filter on query")][String] $Filter = $null
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_Package" -Filter $Filter
    }
}
 
Function Get-SCCMCollection {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Optional Filter on query")][String] $Filter = $null
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_Collection" -Filter $Filter
    }
}
 
Function Get-SCCMAdvertisement {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Optional Filter on query")][String] $Filter = $null
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_Advertisement" -Filter $Filter
    }
}
 
Function Get-SCCMDriver {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Optional Filter on query")][String] $Filter = $null
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_Driver" -Filter $Filter
    }
}
 
Function Get-SCCMDriverPackage {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Optional Filter on query")][String] $Filter = $null
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_DriverPackage" -Filter $Filter
    }
}
 
Function Get-SCCMTaskSequence {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Optional Filter on query")][String] $Filter = $null
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_TaskSequence" -Filter $Filter
    }
}
 
Function Get-SCCMSite {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Optional Filter on query")][String] $Filter = $null
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_Site" -Filter $Filter
    }
}
 
Function Get-SCCMImagePackage {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Optional Filter on query")][String] $Filter = $null
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_ImagePackage" -Filter $Filter
    }
}
 
Function Get-SCCMOperatingSystemInstallPackage {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Optional Filter on query")][String] $Filter = $null
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_OperatingSystemInstallPackage" -Filter $Filter
    }
}
 
Function Get-SCCMBootImagePackage {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Optional Filter on query")][String] $Filter = $null
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_BootImagePackage" -Filter $Filter
    }
}
 
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 
Function Get-SCCMComputer {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Filter on SCCM Resource ID",ValueFromPipelineByPropertyName=$true)][int32] $ResourceID = $false,
        [Parameter(Mandatory=$false, HelpMessage="Filter on Netbiosname on computer",ValueFromPipeline=$true)][String] $NetbiosName = "%",
        [Parameter(Mandatory=$false, HelpMessage="Filter on Domain name",ValueFromPipelineByPropertyName=$true)][Alias("Domain", "Workgroup")][String] $ResourceDomainOrWorkgroup = "%",
        [Parameter(Mandatory=$false, HelpMessage="Filter on SmbiosGuid (UUID)")][String] $SmBiosGuid = "%"
    )
 
    PROCESS {
        if ($ResourceID -eq $false -and $NetbiosName -eq "%" -and $ResourceDomainOrWorkgroup -eq "%" -and $SmBiosGuid -eq "%") {
            throw "Need at least one filter..."
        }
 
        if ($ResourceID -eq $false) {
            return Get-SCCMObject -sccmServer $SccmServer -class "SMS_R_System" -Filter "NetbiosName LIKE '$NetbiosName' AND ResourceDomainOrWorkgroup LIKE '$ResourceDomainOrWorkgroup' AND SmBiosGuid LIKE '$SmBiosGuid'"
        } else {
            return Get-SCCMObject -sccmServer $SccmServer -class "SMS_R_System" -Filter "ResourceID = $ResourceID"
        }
    }
}
 
Function Get-SCCMUser {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Filter on SCCM Resource ID",ValueFromPipelineByPropertyName=$true)][int32] $ResourceID = $false,
        [Parameter(Mandatory=$false, HelpMessage="Filter on unique username in form DOMAIN\UserName",ValueFromPipelineByPropertyName=$true)][String] $UniqueUserName = "%",
        [Parameter(Mandatory=$false, HelpMessage="Filter on Domain name",ValueFromPipelineByPropertyName=$true)][Alias("Domain")][String] $WindowsNTDomain = "%",
        [Parameter(Mandatory=$false, HelpMessage="Filter on UserName",ValueFromPipeline=$true)][String] $UserName = "%"
    )
 
    PROCESS {
        if ($ResourceID -eq $false -and $UniqueUserName -eq "%" -and $WindowsNTDomain -eq "%" -and $UserName -eq "%") {
            throw "Need at least one filter..."
        }
 
        if ($ResourceID -eq $false) {
            return Get-SCCMObject -sccmServer $SccmServer -class "SMS_R_User" -Filter "UniqueUserName LIKE '$UniqueUserName' AND WindowsNTDomain LIKE '$WindowsNTDomain' AND UserName LIKE '$UserName'"
        } else {
            return Get-SCCMObject -sccmServer $SccmServer -class "SMS_R_User" -Filter "ResourceID = $ResourceID"
        }
    }
}
 
Function Get-SCCMCollectionMembers {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="CollectionID", ValueFromPipeline=$true)][String] $CollectionID
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_CollectionMember_a" -Filter "CollectionID = '$CollectionID'"
    }
}
 
Function Get-SCCMSubCollections {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true, HelpMessage="CollectionID",ValueFromPipeline=$true)][Alias("parentCollectionID")][String] $CollectionID
    )
 
    PROCESS {
        return Get-SCCMObject -sccmServer $SccmServer -class "SMS_CollectToSubCollect" -Filter "parentCollectionID = '$CollectionID'"
    }
}
 
Function Get-SCCMParentCollection {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true, HelpMessage="CollectionID",ValueFromPipeline=$true)][Alias("subCollectionID")][String] $CollectionID
    )
 
    PROCESS {
        $parentCollection = Get-SCCMObject -sccmServer $SccmServer -class "SMS_CollectToSubCollect" -Filter "subCollectionID = '$CollectionID'"
 
        return Get-SCCMCollection -sccmServer $SccmServer -Filter "CollectionID = '$($parentCollection.parentCollectionID)'"
    }
}
 
Function Get-SCCMSiteDefinition {
    # Get all definitions for one SCCM site
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer
    )
 
    PROCESS {
        Write-Verbose "Refresh the site $($SccmServer.SccmProvider.SiteCode) control file"
        Invoke-WmiMethod -path SMS_SiteControlFile -name RefreshSCF -argumentList $($SccmServer.SccmProvider.SiteCode) -computername $SccmServer.Machine -namespace $SccmServer.Namespace
 
        Write-Verbose "Get the site definition object for this site"
        return get-wmiobject -query "SELECT * FROM SMS_SCI_SiteDefinition WHERE SiteCode = '$($SccmServer.SccmProvider.SiteCode)' AND FileType = 2" -computername $SccmServer.Machine -namespace $SccmServer.Namespace
    }
}
 
Function Get-SCCMSiteDefinitionProps {
    # Get definitionproperties for one SCCM site
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer
    )
 
    PROCESS {
        return Get-SCCMSiteDefinition -sccmServer $SccmServer | ForEach-Object { $_.Props }
    }
}
 
Function Get-SCCMIsR2 {
    # Return $true if the SCCM server is R2 capable
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer
    )
 
    PROCESS {
        $result = Get-SCCMSiteDefinitionProps -sccmServer $SccmServer | ? {$_.PropertyName -eq "IsR2CapableRTM"}
        if (-not $result) {
            return $false
        } elseif ($result.Value = 31) {
            return $true
        } else {
            return $false
        }
    }
}
 
Function Get-SCCMCollectionRules {
    # Get a set of all collectionrules
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="CollectionID", ValueFromPipeline=$true)][String] $CollectionID
    )
 
    PROCESS {
        Write-Verbose "Collecting rules for $CollectionID"
        $col = [wmi]"$($SccmServer.SccmProvider.NamespacePath):SMS_Collection.CollectionID='$($CollectionID)'"
 
        return $col.CollectionRules
    }
}
 
Function Get-SCCMInboxes {
    # Give a count of files in the SCCM-inboxes
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server",ValueFromPipeline=$true)][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$false, HelpMessage="Minimum number of files in directory")][int32] $minCount = 1
    )
 
    PROCESS {
        Write-Verbose "Reading \\$($SccmServer.Machine)\SMS_$($SccmServer.SccmProvider.SiteCode)\inboxes"
        return Get-ChildItem \\$($SccmServer.Machine)\SMS_$($SccmServer.SccmProvider.SiteCode)\inboxes -Recurse | Group-Object Directory | Where { $_.Count -gt $minCount } | Format-Table Count, Name -AutoSize
    }
}
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 
Function New-SCCMCollection {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true, HelpMessage="Collection Name", ValueFromPipeline=$true)][String] $name,
        [Parameter(Mandatory=$false, HelpMessage="Collection comment")][String] $comment = "",
        [Parameter(Mandatory=$false, HelpMessage="Refresh Rate in Minutes")] [ValidateRange(0, 59)] [int] $refreshMinutes = 0,
        [Parameter(Mandatory=$false, HelpMessage="Refresh Rate in Hours")] [ValidateRange(0, 23)] [int] $refreshHours = 0,
        [Parameter(Mandatory=$false, HelpMessage="Refresh Rate in Days")] [ValidateRange(0, 31)] [int] $refreshDays = 0,
        [Parameter(Mandatory=$false, HelpMessage="Parent CollectionID")][String] $parentCollectionID = "COLLROOT"
    )
 
    PROCESS {
        # Build the parameters for creating the collection
        $arguments = @{Name = $name; Comment = $comment; OwnedByThisSite = $true}
        $newColl = Set-WmiInstance -class "SMS_Collection" -arguments $arguments -computername $SccmServer.Machine -namespace $SccmServer.Namespace
 
        # Hack - for some reason without this we don't get the CollectionID value
        $hack = $newColl.PSBase | select * | out-null
 
        # It's really hard to set the refresh schedule via Set-WmiInstance, so we'll set it later if necessary
        if ($refreshMinutes -gt 0 -or $refreshHours -gt 0 -or $refreshDays -gt 0)
        {
            Write-Verbose "Create the recur interval object"
            $intervalClass = [WMICLASS]"\\$($SccmServer.Machine)\$($SccmServer.Namespace):SMS_ST_RecurInterval"
            $interval = $intervalClass.CreateInstance()
            if ($refreshMinutes -gt 0) {
                $interval.MinuteSpan = $refreshMinutes
            }
            if ($refreshHours -gt 0) {
                $interval.HourSpan = $refreshHours
            }
            if ($refreshDays -gt 0) {
                $interval.DaySpan = $refreshDays
            }
 
            Write-Verbose "Set the refresh schedule"
            $newColl.RefreshSchedule = $interval
            $newColl.RefreshType=2
            $path = $newColl.Put()
        }   
 
        Write-Verbose "Setting the new $($newColl.CollectionID) parent to $parentCollectionID"
        $subArguments  = @{SubCollectionID = $newColl.CollectionID}
        $subArguments += @{ParentCollectionID = $parentCollectionID}
 
        # Add the link
        $newRelation = Set-WmiInstance -Class "SMS_CollectToSubCollect" -arguments $subArguments -computername $SccmServer.Machine -namespace $SccmServer.Namespace
 
        Write-Verbose "Return the new collection with ID $($newColl.CollectionID)"
        return $newColl
    }
}
 
Function Add-SCCMCollectionRule {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true,  HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true,  HelpMessage="CollectionID", ValueFromPipelineByPropertyName=$true)] $collectionID,
        [Parameter(Mandatory=$false, HelpMessage="Computer name to add (direct)", ValueFromPipeline=$true)] [String] $name,
        [Parameter(Mandatory=$false, HelpMessage="WQL Query Expression", ValueFromPipeline=$true)] [String] $queryExpression = $null,
        [Parameter(Mandatory=$false, HelpMessage="Limit to collection (Query)", ValueFromPipeline=$false)] [String] $limitToCollectionId = $null,
        [Parameter(Mandatory=$true,  HelpMessage="Rule Name", ValueFromPipeline=$true)] [String] $queryRuleName
    )
 
    PROCESS {
        # Get the specified collection (to make sure we have the lazy properties)
        $coll = [wmi]"$($SccmServer.SccmProvider.NamespacePath):SMS_Collection.CollectionID='$collectionID'"
 
        # Build the new rule
        if ($queryExpression.Length -gt 0) {
            # Create a query rule
            $ruleClass = [WMICLASS]"$($SccmServer.SccmProvider.NamespacePath):SMS_CollectionRuleQuery"
            $newRule = $ruleClass.CreateInstance()
            $newRule.RuleName = $queryRuleName
            $newRule.QueryExpression = $queryExpression
            if ($limitToCollectionId -ne $null) {
                $newRule.LimitToCollectionID = $limitToCollectionId
            }
 
            $null = $coll.AddMembershipRule($newRule)
        } else {
            $ruleClass = [WMICLASS]"$($SccmServer.SccmProvider.NamespacePath):SMS_CollectionRuleDirect"
 
            # Find each computer
            $computer = Get-SCCMComputer -sccmServer $SccmServer -NetbiosName $name
            # See if the computer is already a member
            $found = $false
            if ($coll.CollectionRules -ne $null) {
                foreach ($member in $coll.CollectionRules) {
                    if ($member.ResourceID -eq $computer.ResourceID) {
                        $found = $true
                    }
                }
            }
            if (-not $found) {
                Write-Verbose "Adding new rule for computer $name"
                $newRule = $ruleClass.CreateInstance()
                $newRule.RuleName = $name
                $newRule.ResourceClassName = "SMS_R_System"
                $newRule.ResourceID = $computer.ResourceID
 
                $null = $coll.AddMembershipRule($newRule)
            } else {
                Write-Verbose "Computer $name is already in the collection"
            }
        }
    }
}
 
Function Add-SCCMDirUserCollectionRule {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(ValueFromPipelineByPropertyName=$true)][String] $CollectionID,
        [Parameter(ValueFromPipeline=$true)][String] $UserName
    )
 
    PROCESS {
        $coll = [wmi]"\\$($SccmServer.Machine)\$($SccmServer.Namespace):SMS_Collection.CollectionID='$CollectionID'"
        $ruleClass = [WMICLASS]"\\$($SccmServer.Machine)\$($SccmServer.Namespace):SMS_CollectionRuleDirect"
 
        $RuleClass
        $UserRule=Get-SCCMUser "userName='$UserName'"
        $NewRuleName=$UserRule.name
        $NewRuleResourceID = $UserRule.ResourceID
        $newRule = $ruleClass.CreateInstance()
 
        $newRule.RuleName = $NewRuleName
        $newRule.ResourceClassName = "SMS_R_User"
        $newRule.ResourceID = $NewRuleResourceID
 
        $null = $coll.AddMembershipRule($newRule)
        $coll.requestrefresh()
        Clear-Variable -name oldrule -errorAction SilentlyContinue
        Clear-Variable -name Coll -errorAction SilentlyContinue
    }
}
 
Function New-SCCMPackage {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true, HelpMessage="Package Name", ValueFromPipeline=$true)][String] $Name,
 
        [Parameter(Mandatory=$false, HelpMessage="Package Version")][String] $Version = "",
        [Parameter(Mandatory=$false, HelpMessage="Package Manufacturer")][String] $Manufacturer = "",
        [Parameter(Mandatory=$false, HelpMessage="Package Language")][String] $Language = "",
        [Parameter(Mandatory=$false, HelpMessage="Package Description")][String] $Description = "",
        [Parameter(Mandatory=$false, HelpMessage="Package Data Source Path")][String] $PkgSourcePath = "",
        [Parameter(Mandatory=$false, HelpMessage="Package Sharename")][String] $PkgShareName = ""
    )
 
    PROCESS {
        $packageClass = [WMICLASS]"\\$($SccmServer.Machine)\$($SccmServer.Namespace):SMS_Package"
        $newPackage = $packageClass.createInstance() 
 
        $newPackage.Name = $Name
        if ($Version -ne "")        { $newPackage.Version = $Version }
        if ($Manufacturer -ne "")   { $newPackage.Manufacturer = $Manufacturer }
        if ($Language -ne "")       { $newPackage.Language = $Language }
        if ($Description -ne "")    { $newPackage.Description = $Description }
 
        if ($PkgSourcePath -ne "") {
            $newPackage.PkgSourceFlag = 2  # Direct (3 = Compressed)
            $newPackage.PkgSourcePath = $PkgSourcePath
            if ($PkgShareName -ne "") {
                $newPackage.ShareName = $PkgShareName
                $newPackage.ShareType = 2
            }
        } else {
            $newPackage.PkgSourceFlag = 1  # No source
            $newPackage.PkgSourcePath = $null
        }
        $newPackage.Put()
 
        $newPackage.Get()
        Write-Verbose "Return the new package with ID $($newPackage.PackageID)"
        return $newPackage
    }
}
 
Function New-SCCMAdvertisement {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true)] $AdvertisementName,
        [Parameter(Mandatory=$true)] $collectionID,
        [Parameter(Mandatory=$true)] $PackageID,
        [Parameter(Mandatory=$true)] $ProgramName,
        [Switch] $Download,
        [Parameter(Mandatory=$false, HelpMessage="YYYYMMDDhhmm")] $StartTime,
        [Parameter(Mandatory=$false, HelpMessage="YYYYMMDDhhmm")] $EndTime,
        [Parameter(Mandatory=$false, HelpMessage="YYYYMMDDhhmm")] $MandatoryTime
    )
    PROCESS {
        $strServer = $SccmServer.machine
        $strNamespace= $SccmServer.namespace
        $AdvClass = [WmiClass]("\\$strServer\" + "$strNameSpace" + ":SMS_Advertisement")
        if ($Download) {
            $RemoteClientFlags = "3152"
        } else {
            $RemoteClientFlags = "3208"
        }
        if ($StartTime -ne $null) {
            $PresentTime = $StartTime + "00.000000+***"
        } else {
            $PresentTime = "20200110000000.000000+***"
        }
        if ($EndTime -ne $null) {
            $ExpirationTime = $Endtime + "00.000000+***"
            $ExpirationTimeEnabled = $true
        } else {
            $ExpirationTime = "20200113000000.000000+***"
            $ExpirationTimeEnabled = $false
        }
        if ($MandatoryTime -ne $null) {
            $Deadline = $MandatoryTime + "00.000000+***"
        } else {
            $Deadline = $null
        }
 
        # Get the all the Advertisement Properties
        $newAdvertisement = $AdvClass.CreateInstance()
        $newAdvertisement.AdvertisementName = $AdvertisementName
        $newAdvertisement.CollectionID = $collectionID
        $newAdvertisement.PackageID = $PackageID
        $newAdvertisement.ProgramName = $ProgramName
        $newAdvertisement.RemoteClientFlags = $RemoteClientFlags
        $newAdvertisement.PresentTime = $PresentTime
        $newAdvertisement.ExpirationTime = $ExpirationTime
        $newAdvertisement.ExpirationTimeEnabled = $ExpirationTimeEnabled
        $newAdvertisement.Priority = "2"
        $newAdvertisement.IncludeSubCollection = $false
 
        # Create Advertisement
        $retval = $newAdvertisement.psbase.Put()
        if ($Deadline -ne $null) {
            # Create Mandatory Schedule
            $wmiClassSchedule = [WmiClass]("\\$strServer\" + "$strNameSpace" + ":SMS_ST_NonRecurring")
            $AssignedSchedule = $wmiClassSchedule.psbase.createinstance()
            $AssignedSchedule.starttime = $Deadline
            if ($Download) {
                $newAdvertisement.RemoteClientFlags = "9296"
            } else {
                $newAdvertisement.RemoteClientFlags = "9352"
            }
            $newAdvertisement.AssignedSchedule = $AssignedSchedule
            $newAdvertisement.AssignedScheduleEnabled = $true
            $newAdvertisement.psbase.put()
            $NewAdvertisementProperties = $newAdvertisement.AssignedSchedule
            foreach ($Adv in $NewAdvertisementProperties) {
                write-verbose "Created Advertisement. Name = $($newAdvertisement.AdvertisementName)"
                write-verbose "Created Advertisement. ID = $newAdvertisement"
                Write-Verbose "Mandatory Deadline created: $($Adv.StartTime)"
            }
        } else {
            write-verbose "Created Advertisement. Name = $($newAdvertisement.AdvertisementName)"
            write-verbose "Created Advertisement. ID = $newAdvertisement"
            write-verbose "No Mandatory-Deadline defined"
        }
    }
}
 
Function New-SCCMProgram {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true, HelpMessage="Program Name")][String] $PrgName = "",
        [Parameter(Mandatory=$true, HelpMessage="Program PackageID")]$PrgPackageID,
        [Parameter(Mandatory=$false, HelpMessage="Program Comment")][String] $PrgComment = "",
        [Parameter(Mandatory=$false, HelpMessage="Program CommandLine")][String] $PrgCommandLine = "",
        [Parameter(Mandatory=$false, HelpMessage="Program MaxRunTime")]$PrgMaxRunTime,
        [Parameter(Mandatory=$false, HelpMessage="Program Diskspace Requirement")]$PrgSpaceReq,
        [Parameter(Mandatory=$false, HelpMessage="Program Working Directory")][String] $PrgWorkDir = "",
        [Parameter(Mandatory=$false, HelpMessage="Program Flags")] $PrgFlags
    )
    PROCESS {
        $programClass = [WMICLASS]"\\$($SccmServer.Machine)\$($SccmServer.Namespace):SMS_Program"
        $newProgram = $programClass.createInstance()
        $newProgram.ProgramName = $PrgName
        $newProgram.PackageID = $PrgPackageID
        if ($PrgComment -ne "") { $newProgram.Comment = $PrgComment }
        if ($PrgCommandLine -ne "") { $newProgram.CommandLine = $PrgCommandLine }
        if ($PrgMaxRunTime -ne $null) { $newProgram.Duration = $PrgMaxRunTime} else { $newProgram.Duration = "0" }
        if ($PrgSpaceReq -ne $null) { $newProgram.DiskSpaceReq = $PrgSpaceReq }
        if ($PrgWorkDir -ne "") { $newProgram.WorkingDirectory = $PrgWorkDir }
        if ($PrgFlags -ne $null) { $newProgram.ProgramFlags = $PrgFlags} else { $newProgram.ProgramFlags = "2299568128" }        
        $newProgram.Put()
        $newProgram.Get()
        Write-Verbose "Return the new program for Package $($newProgram.PackageID)"
        return $newProgram
    }
}
 
Function Add-SCCMDistributionPoint {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true, HelpMessage="PackageID")][String] $DPPackageID,
        [Parameter(Mandatory=$false, HelpMessage="DistributionPoint Servername")][String]$DPName = "",
        [Parameter(Mandatory=$false, HelpMessage="All DistributionPoints of SiteCode")][String] $DPsSiteCode = "",
        [Parameter(Mandatory=$false, HelpMessage="Distribution Point Group")][String] $DPGroupName = "",
        [Switch] $AllDPs
    )
    PROCESS {
        if ($DPName -ne "") {
            $Resource = Get-SCCMObject -SccmServer $SccmServer -class SMS_SystemResourceList -Filter "RoleName = 'SMS Distribution Point' and Servername = '$DPName'"
            $DPClass = [WMICLASS]"\\$($SccmServer.Machine)\$($SccmServer.Namespace):SMS_DistributionPoint"
            $newDistributionPoint = $DPClass.createInstance()
            $newDistributionPoint.PackageID = $DPPackageID
            $newDistributionPoint.ServerNALPath = $Resource.NALPath
            $newDistributionPoint.SiteCode = $Resource.SiteCode
            $newDistributionPoint.Put()
            $newDistributionPoint.Get()
            Write-Verbose "Assigned Package: $($newDistributionPoint.PackageID)"
        }
        if ($DPsSiteCode -ne "") {
            $ListOfResources = Get-SCCMObject -SccmServer $SccmServer -class SMS_SystemResourceList -Filter "RoleName = 'SMS Distribution Point' and SiteCode = '$DPsSiteCode'"
            $DPClass = [WMICLASS]"\\$($SccmServer.Machine)\$($SccmServer.Namespace):SMS_DistributionPoint"
            $newDistributionPoint = $DPClass.createInstance()
            $newDistributionPoint.PackageID = $DPPackageID
            foreach ($resource in $ListOfResources) {
                $newDistributionPoint.ServerNALPath = $Resource.NALPath
                $newDistributionPoint.SiteCode = $Resource.SiteCode
                $newDistributionPoint.Put()
                $newDistributionPoint.Get()
                Write-Verbose "Assigned Package: $($newDistributionPoint.PackageID)"
            }
        }
        if ($DPGroupName -ne "") {
            $DPGroup = Get-SCCMObject -sccmserver $SccmServer -class SMS_DistributionPointGroup -Filter "sGroupName = '$DPGroupName'"
            $DPGroupNALPaths = $DPGroup.arrNALPath
            $DPClass = [WMICLASS]"\\$($SccmServer.Machine)\$($SccmServer.Namespace):SMS_DistributionPoint"
            $newDistributionPoint = $DPClass.createInstance()
            $newDistributionPoint.PackageID = $DPPackageID
            foreach ($DPGroupNALPath in $DPGroupNALPaths) {
                $DPResource = Get-SCCMObject -SccmServer $SccmServer -class SMS_SystemResourceList -Filter "RoleName = 'SMS Distribution Point'" | Where-Object {$_.NALPath -eq $DPGroupNALPath}
                if ($DPResource -ne $null) {
                    Write-Verbose "$DPResource"
                    $newDistributionPoint.ServerNALPath = $DPResource.NALPath
                    Write-Verbose "ServerNALPath = $($newDistributionPoint.ServerNALPath)"
                    $newDistributionPoint.SiteCode = $DPResource.SiteCode
                    Write-Verbose "SiteCode = $($newDistributionPoint.SiteCode)"
                    $newDistributionPoint.Put()
                    $newDistributionPoint.Get()
                    Write-Host "Assigned Package: $($newDistributionPoint.PackageID) to $($DPResource.ServerName)"
                } else {
                    Write-Host "DP not found = $DPGroupNALPath"
                }
            }
        }
        if ($AllDPs) {
            $ListOfResources = Get-SCCMObject -SccmServer $SccmServer -class SMS_SystemResourceList -Filter "RoleName = 'SMS Distribution Point'"
            $DPClass = [WMICLASS]"\\$($SccmServer.Machine)\$($SccmServer.Namespace):SMS_DistributionPoint"
            $newDistributionPoint = $DPClass.createInstance()
            $newDistributionPoint.PackageID = $DPPackageID
            foreach ($resource in $ListOfResources) {
                $newDistributionPoint.ServerNALPath = $Resource.NALPath
                $newDistributionPoint.SiteCode = $Resource.SiteCode
                $newDistributionPoint.Put()
                $newDistributionPoint.Get()
                Write-Verbose "Assigned Package: $($newDistributionPoint.PackageID) $($newDistributionPoint.ServerNALPath)"
            }
        }
    }
}
 
Function Update-SCCMDriverPkgSourcePath {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true, HelpMessage="Current Path", ValueFromPipeline=$true)][String] $currentPath,
        [Parameter(Mandatory=$true, HelpMessage="New Path", ValueFromPipeline=$true)][String] $newPath
    )
 
    PROCESS {
        Get-SCCMDriverPackage -sccmserver $SccmServer | Where-Object {$_.PkgSourcePath -ilike "*$($currentPath)*" } | Foreach-Object {
            $newSourcePath = ($_.PkgSourcePath -ireplace [regex]::Escape($currentPath), $newPath)
            Write-Verbose "Changing from '$($_.PkgSourcePath)' to '$($newSourcePath)' on $($_.PackageID)"
            $_.PkgSourcePath = $newSourcePath
            $_.Put() | Out-Null
        }
    }
}
 
Function Update-SCCMPackageSourcePath {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true, HelpMessage="Current Path", ValueFromPipeline=$true)][String] $currentPath,
        [Parameter(Mandatory=$true, HelpMessage="New Path", ValueFromPipeline=$true)][String] $newPath
    )
 
    PROCESS {
        Get-SCCMPackage -sccmserver $SccmServer | Where-Object {$_.PkgSourcePath -ilike "*$($currentPath)*" } | Foreach-Object {
            $newSourcePath = ($_.PkgSourcePath -ireplace [regex]::Escape($currentPath), $newPath)
            Write-Verbose "Changing from '$($_.PkgSourcePath)' to '$($newSourcePath)' on $($_.PackageID)"
            $_.PkgSourcePath = $newSourcePath
            $_.Put() | Out-Null
        }
    }
}
 
Function Update-SCCMDriverSourcePathnotepad {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="SCCM Server")][Alias("Server","SmsServer")][System.Object] $SccmServer,
        [Parameter(Mandatory=$true, HelpMessage="Current Path", ValueFromPipeline=$true)][String] $currentPath,
        [Parameter(Mandatory=$true, HelpMessage="New Path", ValueFromPipeline=$true)][String] $newPath
    )
 
    PROCESS {
        Get-SCCMDriver -sccmserver $SccmServer | Where-Object {$_.ContentSourcePath -ilike "*$($currentPath)*" } | Foreach-Object {
            $newSourcePath = ($_.ContentSourcePath -ireplace [regex]::Escape($currentPath), $newPath)
            Write-Verbose "Changing from '$($_.ContentSourcePath)' to '$($newSourcePath)' on $($_.PackageID)"
            $_.ContentSourcePath = $newSourcePath
            $_.Put() | Out-Null
        }
    }
}
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
# EOF