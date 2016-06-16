[CmdletBinding(SupportsShouldProcess=$true)]
Param(
    [Parameter()]
    [string]$DeploymentType = "All",

    [Parameter()]
    [string]$Visibility,

    [Parameter()]
    [string]$TenantId
)

Add-Type  -ErrorAction SilentlyContinue -TypeDefinition @"
   public enum DeploymentType
   {
      All,
      DefaultToBusiness,
      EnableAddAccounts
   }
"@

Add-Type  -ErrorAction SilentlyContinue -TypeDefinition @"
   public enum Visibility
   {
        Silent
   }
"@

function Install-OneDriveForBusiness{
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true)]
    [DeploymentType]$DeploymentType,

    [Parameter()]
    [Visibility]$Visibility,

    [Parameter()]
    [string]$TenantId
)

Begin{
    $currentExecutionPolicy = Get-ExecutionPolicy
	Set-ExecutionPolicy Unrestricted -Scope Process -Force  
    $startLocation = Get-Location
}

Process{
    #Add the registry keys
    if($DeploymentType -eq "All"){
        $dtbRegCheck = Test-Path ".\DefaultToBusinessFRE.reg"
        $eaaRegCheck = Test-Path ".\EnableAddAccounts.reg"

        if(!$dtbRegCheck){
            Write-Host ""
            Write-Warning "Unable to find the DefaultToBusinessFRE registry key."
        }
        else{
            regedit.exe /s .\DefaultToBusinessFRE.reg
        }

        if(!$eaaRegCheck){
            Write-Host ""
            Write-Warning "Unable to find the EnableAddAccounts registry key."
        }
        else{
            regedit.exe /s .\EnableAddAccounts.reg
        }
    }

    if($DeploymentType -eq "DefaultToBusiness"){
        if(!(Test-Path .\DefaultToBusinessFRE.reg)){
            Write-Warning "Unable to find the DefaultToBusinessFRE registry key."
        }
        else{
            regedit.exe /s .\DefaultToBusinessFRE.reg
        }
    }

    if($DeploymentType -eq "EnableAddAccounts"){
        if(!(Test-Path .\EnableAddAccounts.reg)){
            Write-Warning "Unable to find the EnableAddAccounts registry key."
        }
        else{
            regedit.exe /s .\EnableAddAccounts.reg
        }
    }

    #Install OneDriveSetup.exe
    $OneDriveExePath = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
    if($Visibility -eq "Silent"){
        & .\OneDriveSetup.exe /silent
        if($TenantId){              
            & $OneDriveExePath "/configure_business:$TenantId"
        }                                 
    }
    else{
        if($TenantId){
            if(!(Test-Path $OneDriveExePath)){
                & .\OneDriveSetup.exe /silent
            }
            & $OneDriveExePath "/configure_business:$TenantId" 
        }
        else{
            & .\OneDriveSetup.exe
            
            Do{
                Get-Process -Name OneDriveSetup -ErrorAction Ignore | Out-Null
                Start-Sleep -Seconds 5
                $processCheck = Get-Process -Name OneDriveSetup -ErrorAction Ignore
            }
            Until($processCheck -eq $null)

            $OneDriveExePath = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
            if(Test-Path $OneDriveExePath){
                Write-Host ""
                Write-Host "OneDrive.exe has been successfully installed."
            }  
        }                  
    }            
}
   
}

if($DeploymentType -eq "All"){
    if($Visibility-eq "Silent"){
        if($TenantId){       
            Install-OneDriveForBusiness -DeploymentType All -Visibility Silent -TenantId $TenantId
        }
        else{
            Install-OneDriveForBusiness -DeploymentType All -Visibility Silent
        }
    }
    else{
        if($TenantId){
            Install-OneDriveForBusiness -DeploymentType All -TenantId $TenantId
        }
        else{
            Install-OneDriveForBusiness -DeploymentType All
        }
    }
}

if($DeploymentType -eq "DefaultToBusiness"){
    if($Visibility-eq "Silent"){       
        if($TenantId){       
            Install-OneDriveForBusiness -DeploymentType DefaultToBusiness -Visibility Silent -TenantId $TenantId
        }
        else{
            Install-OneDriveForBusiness -DeploymentType DefaultToBusiness -Visibility Silent
        }
    }
    else{
        if($TenantId){
            Install-OneDriveForBusiness -DeploymentType DefaultToBusiness -TenantId $TenantId
        }
        else{
            Install-OneDriveForBusiness -DeploymentType DefaultToBusiness
        }
    }
}

if($DeploymentType -eq "EnableAddAccounts"){
    if($Visibility-eq "Silent"){       
        if($TenantId){       
            Install-OneDriveForBusiness -DeploymentType EnableAddAccounts -Visibility Silent -TenantId $TenantId
        }
        else{
            Install-OneDriveForBusiness -DeploymentType EnableAddAccounts -Visibility Silent
        }
    }
    else{
        if($TenantId){
            Install-OneDriveForBusiness -DeploymentType EnableAddAccounts -TenantId $TenantId
        }
        else{
            Install-OneDriveForBusiness -DeploymentType EnableAddAccounts
        }
    }
}    