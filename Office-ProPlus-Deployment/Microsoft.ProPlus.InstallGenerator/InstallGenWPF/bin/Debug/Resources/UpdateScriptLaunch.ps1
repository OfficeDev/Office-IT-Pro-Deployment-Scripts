param(
        [Parameter()]
        [string] $machineToRun = $Null,

        [Parameter()]
        [bool] $DisplayLevel = $false,

        [Parameter()]
        [string] $UpdateToVersion = $NULL,

        [Parameter()]
        [string] $Channel = $NULL
        
    )

    Process{
    try{
    if ($PSScriptRoot) {   $scriptPath = $PSScriptRoot } else {   $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition }


        $textToWrite = Invoke-Command -ComputerName $machineToRun -FilePath $scriptPath\Update-Office.ps1 -ArgumentList $UpdateToVersion, $Channel, $DisplayLevel

    }
    catch    {
    Write-Host $_.Exception.Message
           throw;
    }
    $filename = $machineToRun+"PowershellAttempt.txt"
    $textToWrite | Out-File $env:temp\$filename

    #Stop-Process -Id $PID
    }
