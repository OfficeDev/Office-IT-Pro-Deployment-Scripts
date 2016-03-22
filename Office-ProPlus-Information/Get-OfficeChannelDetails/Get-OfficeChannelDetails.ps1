function Get-ChannelXml() {
    [CmdletBinding()]	
    Param
	(
	    [Parameter()]
	    [string]$FolderPath = $null,

	    [Parameter()]
	    [bool]$OverWrite = $false
	)

   process {
       $cabPath = "$PSScriptRoot\ofl.cab"

       $webclient = New-Object System.Net.WebClient
       $XMLFilePath = "$env:TEMP/ofl.cab"
       $XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"
       $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)

       $tmpName = "o365client_64bit.xml"
       expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\o365client_64bit.xml"
       
       [xml]$channelXml = Get-Content $tmpName

       return $channelXml
   }

}

function Get-ChannelUrl() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [Channel]$Channel
   )

   Process {
      $channelXml = Get-ChannelXml

      $currentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
      return $currentChannel
   }

}

function Get-BranchLatestVersion() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string]$ChannelUrl,

      [Parameter(Mandatory=$true)]
      [string]$Channel
   )

   process {

    $webclient = New-Object System.Net.WebClient
    $CABFilePath = "$env:TEMP/v64.cab"
    if (Test-Path -Path $CABFilePath) {
      Remove-Item -Path $CABFilePath -Force
    }

    $XMLDownloadURL = "$ChannelUrl/Office/Data/v64.cab"
    $webclient.DownloadFile($XMLDownloadURL,$CABFilePath)

    $tmpName = "VersionDescriptor.xml"
    expand $CABFilePath $env:TEMP -f:$tmpName | Out-Null
    $tmpName = $env:TEMP + "\VersionDescriptor.xml"
    [xml]$versionXml = Get-Content $tmpName

    return $versionXml.Version.Available.Build
   }
}

$ChannelXml = Get-ChannelXml
$results = new-object PSObject[] 0;
$Channels = @("Deferred", "Current", "FirstReleaseDeferred", "FirstReleaseCurrent")

foreach ($Channel in $Channels) {
    $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
    $latestVersion = Get-BranchLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel

    $Result = New-Object –TypeName PSObject 
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Channel" -Value $Channel
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "LatestVersion" -Value $latestVersion
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "URL" -Value $selectChannel.URL
    $Results += $Result
}

$Results