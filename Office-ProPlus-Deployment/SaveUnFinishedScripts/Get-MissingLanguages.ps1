[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [string[]] $UpdateSource,
    [string] $SQLServer,
    [string] $SiteCode
)

function Get-MissingLanguages{
param(
    [string] $UpdateSource
)
    
Begin{
    $defaultDisplaySet = 'UpdateSource,UpdateSourceLanguages,MissingLanguages'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}

Process{
    $results = New-Object PSObject[] 0;

    $UpdateSource = "C:\Packages"     
    $sourceFolders = Get-ChildItem $UpdateSource -Recurse | Where-Object {$_.Name -contains "office"} | foreach {$_.FullName} 

    $envLanguages = @("en-us","es-es","ru-ru","fr-fr")
    $updateSourcesLangs = Get-OfficeLanguages -SQLServer $SQLServer
    $sourceLangs = foreach($source in $sourceFolders){
        $streams = Get-ChildItem -Path $source -Recurse | Where-Object {($_.Name -like "stream*") -and ($_.Name -notlike "*none*")} | foreach {$_.Name}

        $updateSourceLangs = foreach($stream in $streams){
            $stream = $stream.Split(".")[2]
            $stream
        }
        $missingLangs = Compare-Object -ReferenceObject ($envLanguages | Sort-Object) -DifferenceObject ($updateSourceLangs | Sort-Object) | Where-Object {$_.SideIndicator -eq "<="} | foreach {$_.InputObject}

        $object = New-Object PSObject -Property @{UpdateSource = $source; UpdateSourceLanguages = $updateSourceLangs; MissingLanguages = $missingLangs}
        $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
        $results += $object
    }

    $results = Get-Unique -InputObject $results 
    $results | select UpdateSource,UpdateSourceLanguages,MissingLanguages -Unique | ft -AutoSize -Wrap
}
}

function Get-OfficeLanguages{
param(
[Parameter(Mandatory=$true)]
[string] $SQLServer = "SCCMLAB-SCCM",
[string] $SiteCode,
[array] $SqlQuery = @'
SELECT TOP 1000 [MachineID]
      ,[DisplayName00]
      ,[ProdID00]  
      ,[TimeKey]         
      ,[InstallDate00]
      ,[Publisher00]
      ,[Version00]
  FROM [CM_S01].[dbo].[Add_Remove_Programs_64_DATA]
'@
)

begin{
    $defaultDisplaySet = 'OfficeLanguages'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}

process{

    $results = New-Object PSObject[] 0;

    if(!$SiteCode){
        $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
    }
    
    $Database = "CM_$SiteCode"

    ## - Connect to SQL Server using non-SMO class 'System.Data';
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection;
    $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $Database; Integrated Security = True";

    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand;
    $SqlCmd.CommandText = $SqlQuery;
    $SqlCmd.Connection = $SqlConnection;

    ## - Extract and build the SQL data object '$DataSetTable';
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter;
    $SqlAdapter.SelectCommand = $SqlCmd;
    $DataSet = New-Object System.Data.DataSet;
    $SqlAdapter.Fill($DataSet) | Out-Null;
    $DataSetTable = $DataSet.Tables["Table"];

    foreach($data in $DataSetTable){    
        if($DataSetTable.DisplayName00 -like "*Microsoft Office 365*"){
            $displayName = $DataSetTable.DisplayName00 | Where-Object {$_ -like "*Microsoft Office 365*"}
            foreach($dn in $displayName){
                $dn = $dn.split("")[5]
                $object = New-Object PSObject -Property @{OfficeLanguages = $dn}
                $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                $results += $object
            }    
        }
    }
    
    $results = Get-Unique -InputObject $results 
    $results | select OfficeLanguages -Unique | foreach {$_.OfficeLanguages}
}

}

Get-MissingLanguages -UpdateSource $UpdateSource