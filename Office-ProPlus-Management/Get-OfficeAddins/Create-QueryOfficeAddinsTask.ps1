[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
    [Parameter()]
    [bool] $WaitForUpdateToFinish = $true,

    [Parameter()]
    [bool] $EnableUpdateAnywhere = $true,

    [Parameter()]
    [bool] $ForceAppShutdown = $true,

    [Parameter()]
    [bool] $UpdatePromptUser = $false,

    [Parameter()]
    [bool] $DisplayLevel = $false,

    [Parameter()]
    [string] $UpdateToVersion = $NULL,

    [Parameter()]
    [bool] $UseRandomStartTime = $true,

    [Parameter()]
    [string] $RandomTimeStart = "08:00",

    [Parameter()]
    [string] $RandomTimeEnd = "17:00",

    [Parameter()]
    [string] $StartTime = "12:00",

    [Parameter()]
    [string] $LogPath = $NULL,

    [Parameter()]
    [string] $LogName = $NULL,
        
    [Parameter()]
    [bool] $ValidateUpdateSourceFiles = $true,

    [Parameter()]
    [bool] $UseScriptLocationAsUpdateSource = $false
    
)

Function Create-QueryOfficeAddinsTask {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [bool] $WaitForUpdateToFinish = $true,

        [Parameter()]
        [bool] $EnableUpdateAnywhere = $true,

        [Parameter()]
        [bool] $UpdatePromptUser = $false,

        [Parameter()]
        [bool] $DisplayLevel = $false,

        [Parameter()]
        [bool] $UseRandomStartTime = $true,

        [Parameter()]
        [string] $RandomTimeStart = "08:00",

        [Parameter()]
        [string] $RandomTimeEnd = "17:00",

        [Parameter()]
        [string] $StartTime = "12:00"    
    )

    Begin {
	    $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }

    Process {
       $TaskName = "Query Office Add-ins"
       $scriptRoot = GetScriptRoot
 
       if ($UseRandomStartTime) {
           $taskStartTime = Generate-StartTime
       } else {
           $taskStartTime = $StartTime
       }

       $outputPath = "$env:temp\QueryOfficeAddinsTask.xml"

       if (Test-Path -Path "$scriptRoot\Get-OfficeAddins.ps1") {
           Copy-Item -Path "$scriptRoot\Get-OfficeAddins.ps1" -Destination "$env:Windir\Temp\Get-OfficeAddins.ps1" -Force
       }

       $exePath = "PowerShell -Command $env:windir\Temp\Get-OfficeAddins.ps1.ps1"

       if ($UpdateToVersion) {
          $exePath += " -UpdateToVersion " + $UpdateToVersion
       }

       $runAsUser = "NT AUTHORITY\SYSTEM"
       if (($UpdatePromptUser) -or ($DisplayLevel) -or (!($ForceAppShutdown))) {
          $runAsUser = "BUILTIN\Users"
       }

       schtasks /create /tn $TaskName /tr `"$exePath`" /sc WEEKLY /st $taskStartTime /f /D TUE /RU $runAsUser /RL Highest | Out-null
       schtasks /query /tn $TaskName /xml > $outputPath  | Out-null

       [xml]$xml = Get-Content -Path $outputPath

       $xml.Task.RegistrationInfo.Author = ""

       $xml.Task.Settings.DisallowStartIfOnBatteries = "false"
       $xml.Task.Settings.StopIfGoingOnBatteries = "false"

       if (!($xml.Task.Settings.Priority)) {
          $priority = $xml.CreateElement("Priority", "http://schemas.microsoft.com/windows/2004/02/mit/task")
          $xml.Task.Settings.AppendChild($priority) | Out-Null
       }
       $xml.Task.Settings.Priority = "4"

       if (!($xml.Task.Settings.WakeToRun)) {
          $wakeToRun = $xml.CreateElement("WakeToRun", "http://schemas.microsoft.com/windows/2004/02/mit/task")
          $xml.Task.Settings.AppendChild($wakeToRun) | Out-Null
       }
       $xml.Task.Settings.WakeToRun = "true"

       if (!($xml.Task.Settings.StartWhenAvailable)) {
          $startWhenAvailable = $xml.CreateElement("StartWhenAvailable", "http://schemas.microsoft.com/windows/2004/02/mit/task")
          $xml.Task.Settings.AppendChild($startWhenAvailable) | Out-Null
       }
       $xml.Task.Settings.StartWhenAvailable = "true"

       $xml.Save($outputPath)
  
       schtasks /delete /tn $TaskName /F | Out-null
       schtasks /create /tn $TaskName /xml $outputPath /F
    }

    End {
       $setLocation = Set-Location $startLocation
    }

}

Function GetScriptRoot() {
 process {
     [string]$scriptPath = "."

     if ($PSScriptRoot) {
       $scriptPath = $PSScriptRoot
     } else {
       $scriptPath = (Get-Item -Path ".\").FullName
     }

     return $scriptPath
 }
}

Function Generate-StartTime() {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [string] $RandomTimeStart = "08:00",

        [Parameter()]
        [string] $RandomTimeEnd = "17:00"
    )

    Process {

        [DateTime]$theMin = [DateTime]::Now.Date
        [DateTime]$theMax = [DateTime]::Now.Date.AddDays(1).AddSeconds(-1)

        $strTheMin = $theMin.Date.ToShortDateString() + " " + $RandomTimeStart
        $strTheMax = $theMax.Date.ToShortDateString() + " " + $RandomTimeEnd
    
        $theMin = [DateTime]$strTheMin
        $theMax = [DateTime]$strTheMax

        $theRandomGen = New-Object random
        $theRandomTicks = [Convert]::ToInt64( ($theMax.ticks * 1.0 - $theMin.Ticks * 1.0 ) * $theRandomGen.NextDouble() + $theMin.Ticks * 1.0 )
        [DateTime]$startDate = new-object DateTime($theRandomTicks)

        $hour = $startDate.Hour.ToString()
        $minute = $startDate.Minute.ToString()

        if ($hour.Length -eq 1) { $hour = "0" + $hour }
        if ($minute.Length -eq 1) { $minute = "0" + $minute }

        $startTime = $hour + ":" + $minute
        return $startTime
    }
}

Function Convert-Bool() {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$true)]
        [bool] $value
    )

    $newValue = "$" + $value.ToString()
    return $newValue 
}

Create-Office365AnywhereTask `
-WaitForUpdateToFinish $WaitForUpdateToFinish `
-EnableUpdateAnywhere $EnableUpdateAnywhere `
-ForceAppShutdown $ForceAppShutdown `
-UpdatePromptUser $UpdatePromptUser `
-DisplayLevel $DisplayLevel `
-UpdateToVersion $UpdateToVersion `
-UseRandomStartTime $UseRandomStartTime `
-RandomTimeStart $RandomTimeStart `
-RandomTimeEnd $RandomTimeEnd `
-StartTime $StartTime `
-LogPath $LogPath `
-LogName $LogName `
-ValidateUpdateSourceFiles $ValidateUpdateSourceFiles

$scriptRoot = GetScriptRoot

if (Test-Path -Path "$scriptRoot\Update-Office365Anywhere.ps1") {

& $scriptRoot\Update-Office365Anywhere.ps1 -WaitForUpdateToFinish $WaitForUpdateToFinish `
    -EnableUpdateAnywhere $EnableUpdateAnywhere `
    -ForceAppShutdown $ForceAppShutdown `
    -UpdatePromptUser $UpdatePromptUser `
    -DisplayLevel $DisplayLevel `
    -UpdateToVersion $UpdateToVersion `
    -LogPath $LogPath `
    -LogName $LogName `
    -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles `
    -UseScriptLocationAsUpdateSource $UseScriptLocationAsUpdateSource
}