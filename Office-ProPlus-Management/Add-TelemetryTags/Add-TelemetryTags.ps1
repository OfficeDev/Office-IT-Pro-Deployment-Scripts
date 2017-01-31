
Function AddTags{
    $errObjectNotCreated = 429
    $strTag1,$strTag2,$strTag3,$strTag4


    $strTag4 = CurrentUserDomain
    $objUser = GetActiveDirectoryInformation
    $strTag1 = $objUser.Properties.department
    $strTag3 = $objUser.Properties.physicaldeliveryofficename
    $strTag2 = $objUser.Properties.title
    #Write-Host $strTag1
    WriteToRegistry -strValueName "Tag1" -strValueData $strTag1
    WriteToRegistry -strValueName "Tag2" -strValueData $strTag2
    WriteToRegistry -strValueName "Tag3" -strValueData $strTag3
    WriteToRegistry -strValueName "Tag4" -strValueData $strTag4
    
}

Function GetActiveDirectoryInformation{

    
    $objUser = Get-LocalLogonInformation
    return $objUser

}

function Get-LocalLogonInformation
{

    $strName = $env:USERNAME
    $strFilter1 = "(&(objectCategory=User)(samAccountName=$strName))"

    $objFinder1 = New-Object System.DirectoryServices.DirectorySearcher
    $objFinder1.Filter = $strFilter1

    $objPath1 = $objFinder1.FindOne()
    $findCompINfo = $objPath1.Path




    $dn = New-Object System.DirectoryServices.DirectoryEntry($findCompINfo)
    $strFilter = "((objectCategory=person))"

    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($dn)
    $objSearcher.Filter = $strFilter

    

    $objSearcher.PropertiesToLoad.Add("physicalDeliveryOfficeName")
    $objSearcher.PropertiesToLoad.Add("title")
    $objSearcher.PropertiesToLoad.Add("department")
    $check = $objSearcher.FindOne()
    
    
    return $check
}



Function CurrentUserDomain{
    $objWScriptNetwork = New-Object -ComObject Wscript.Network
    $tempVar = $objWScriptNetwork.UserDomain
    return $tempVar
}


Function WriteToRegistry{
PARAM(
    [string]$strValueName,
    [string]$strValueData
)
    $regPath = "HKCU:\Software\Microsoft\Office\15.0\osm\"
    New-ItemProperty -Path $regPath -Name $strValueName -Value $strValueData -PropertyType String -Force | Out-Null

}