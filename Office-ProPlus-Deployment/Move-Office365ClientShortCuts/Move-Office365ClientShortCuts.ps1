function Move-Office365ClientShortCuts {
    [CmdletBinding()]
    Param(
       [Parameter(ValueFromPipelineByPropertyName=$true, Position=0)]
       [string]$FolderName = "Microsoft Office 2016",
       
       [Parameter(ValueFromPipelineByPropertyName=$true, Position=1)]
       [bool]$MoveToolsFolder = $false,
       
       [Parameter()]
       [string]$LogFilePath                                                                        
    )

    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    $sh = New-Object -COM WScript.Shell
    $programsPath = $sh.SpecialFolders.Item("AllUsersStartMenu")

    #Create new subfolder                                                                       
    if(!(Test-Path -Path "$programsPath\Programs\$FolderName")){
        New-Item -ItemType directory -Path "$programsPath\Programs\$FolderName"  -ErrorAction Stop | Out-Null
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Directory created: $programsPath\Programs\$FolderName" -LogFilePath $LogFilePath
    }    

    if ($MoveToolsFolder) {
        $toolsPath = "$programsPath\Programs\Microsoft Office 2016 Tools"
        if(Test-Path -Path $toolsPath){
            Move-Item -Path $toolsPath -Destination "$programsPath\Programs\$FolderName\Microsoft Office 2016 Tools"  -ErrorAction Stop | Out-Null
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Moved: $toolsPath to Destination: $programsPath\Programs\$FolderName\Microsoft Office 2016 Tools" -LogFilePath $LogFilePath
        }    
    }
    
    $items = Get-ChildItem -Path "$programsPath\Programs"

    $OfficeInstallPath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun" -Name "InstallPath").InstallPath
    
    $itemsToMove = $false
    foreach ($item in $items) {
       if ($item.Name -like "*.lnk") {

           $itemName = $item.Name
           
           $targetPath = $sh.CreateShortcut($item.fullname).TargetPath

           if ($targetPath -like "$OfficeInstallPath\root\*") {
              $itemsToMove = $true
              $movePath = "$programsPath\Programs\$FolderName\$itemName"

              Move-Item -Path $item.FullName -Destination $movePath -Force -ErrorAction Stop

              Write-Host "$itemName Moved"
              WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "$itemName Moved" -LogFilePath $LogFilePath
           }
       }
    }    

    if (!($itemsToMove)) {
       Write-Host "There are no Office 365 ProPlus client ShortCuts to Move"
       WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "There are no Office 365 ProPlus client ShortCuts to Move" -LogFilePath $LogFilePath
    }
}

function Get-CurrentLineNumber {
    $MyInvocation.ScriptLineNumber
}

function Get-CurrentFileName{
    $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
}

Function WriteToLogFile() {
    param( 
        [Parameter(Mandatory=$true)]
        [string]$LNumber,

        [Parameter(Mandatory=$true)]
        [string]$FName,

        [Parameter(Mandatory=$true)]
        [string]$ActionError,

        [Parameter()]
        [string]$LogFilePath
    )

    try{
        $headerString = "Time".PadRight(30, ' ') + "Line Number".PadRight(15,' ') + "FileName".PadRight(60,' ') + "Action"
        $stringToWrite = $(Get-Date -Format G).PadRight(30, ' ') + $($LNumber).PadRight(15, ' ') + $($FName).PadRight(60,' ') + $ActionError

        if(!$LogFilePath){
            $LogFilePath = "$env:windir\Temp\" + (Get-Date -Format u).Substring(0,10)+"_OfficeDeploymentLog.txt"
        }
        if(Test-Path $LogFilePath){
             Add-Content $LogFilePath $stringToWrite
        }
        else{#if not exists, create new
             Add-Content $LogFilePath $headerString
             Add-Content $LogFilePath $stringToWrite
        }
    } catch [Exception]{
        Write-Host $_
    }
}