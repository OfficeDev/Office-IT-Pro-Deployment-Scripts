

# Find a list of available add-ins
function Find-ComAddins {

    Write-Host "`nLooking for COM add-ins..."`n

    $addinOfficePath = "HKCU:\Software\Microsoft\Office"

    $result = Get-ChildItem -Path $addinOfficePath -Recurse | Where-Object { $_.PsChildName -match 'Addins' } | Get-ChildItem | Get-ItemProperty | select PSChildName,FriendlyName

    $result
}

# Look for Excel add-in files (xla,xlam, etc)
function Find-AddinExt {

    Write-Host "`nLooking for add-in files by extension..."`n

        
    $dir86 = Get-ChildItem ${env:ProgramFiles(x86)} -Recurse
    $dir = Get-ChildItem $env:ProgramFiles -Recurse
    $list = $dir | where { $_.Extension -eq ".xla" -or $_.Extension -eq ".xlam" -or $_.Extension -eq ".xll" -or $_.Extension -eq ".ppa" -or $_.Extension -eq ".ppam" -or $_.Extension -eq ".pa" -or $_.Extension -eq ".accda"  -or $_.Extension -eq ".mda" -or $_.Extension -eq ".wll" }
    $list86 = $dir86 | where { $_.Extension -eq ".xla" -or $_.Extension -eq ".xlam" -or $_.Extension -eq ".xll" -or $_.Extension -eq ".ppa" -or $_.Extension -eq ".ppam" -or $_.Extension -eq ".pa" -or $_.Extension -eq ".accda"  -or $_.Extension -eq ".mda" -or $_.Extension -eq ".wll" }
    $files = $list += $list86 
    $files = $files | Select Name, FullName
    $files
}
    
$ComAddins = Find-ComAddins

$ExtAddins = Find-AddinExt

$Addins = $ComAddins, $ExtAddins
$Addins[0] | Format-Table -a
$Addins[1] | Format-Table -a