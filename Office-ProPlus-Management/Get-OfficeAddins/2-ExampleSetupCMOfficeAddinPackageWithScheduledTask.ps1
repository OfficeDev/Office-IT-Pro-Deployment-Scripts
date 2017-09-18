#  Office Add-in reporting script example
#
#  This script demonstrates how to create a System Center Configuration Manager package containing
#  a script that will store Office add-in information into a WMI class. This information can then be loaded
#  into a SQL database as a new Hardware Inventory class.

Process {
    # Get the path to the script
    $scriptPath = "."
    
    if ($PSScriptRoot) {
      $scriptPath = $PSScriptRoot
    } else {
      $scriptPath = (Get-Item -Path ".\").FullName
    }
    
    # Importing the required functions
    . $scriptPath\Setup-CMOfficeAddinPackage.ps1
    
    # Create the package
    $PackageName = "Update Office add-in repository"
    $ScriptFilesPath = $scriptPath
    
    Create-CMOfficeAddinPackage -PackageName $PackageName -ScriptFilesPath $ScriptFilesPath -MoveScriptFiles $true
    
    # Create the program
    $ProgramName = "Update with Scheduled Task"
    
    Create-CMOfficeAddinTaskProgram -PackageName $PackageName -ProgramName $ProgramName -UseRandomStartTime $true -RandomTimeStart "06:00" -RandomTimeEnd "18:00"
    
    # Distribute the package to a distribution point
    $DistributionPoint = "CM01.Contoso.com"
    
    Distribute-CMOfficeAddinPackage -PackageName $PackageName -DistributionPoint $DistributionPoint -WaitForDistributionToFinish $true
    
    # Deploy the program to a device collection
    $Collection = "All Desktop and Server Clients"
    
    Deploy-CMOfficeAddinProgram -PackageName $PackageName -ProgramName $ProgramName -Collection $Collection -DeploymentPurpose Available
}