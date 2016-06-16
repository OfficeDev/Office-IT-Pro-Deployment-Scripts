#This package provider is based off of the software found here: https://github.com/OneGet/MyAlbum-Sample-Provider
Import-Module BitsTransfer
$Providername = "OfficeProvider"



Function Initialize-Provider     { Write-Debug "In $($Providername) - Initialize-Provider" }
Function Get-PackageProviderName { Return $Providername }

function checkForLanguage() {

    $langId =  Get-Culture | Select-Object -ExpandProperty Name

    $availableLangs = @("en-us",
    "ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee",
    "fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
    "ja-jp","kk-kh","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br",
    "pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
    "tr-tr","uk-ua");

    if ($availableLangs -contains ($langId.Trim().ToLower())) {
       return $langId
    } else {
       $langStart = $langId.Split('-')[0]
       $checkLang = $NULL

       foreach ($availabeLang in $availableLangs) {
          if ($availabeLang.ToLower().StartsWith($langStart.ToLower())) {
             $checkLang = $availabeLang
             break;
          }
       }

       return $checkLang
    }
}

function Get-DynamicOptions
{
    param
    (
       [Microsoft.PackageManagement.MetaProvider.PowerShell.OptionCategory] 
       $category
    )

    Write-Debug ($LocalizedData.ProviderDebugMessage -f ('Get-DynamicOptions'))   
    switch($category)
    {
            
            Install {
                $allowedBitness = @("64","32")
    
                $allowedChannel = @("Current","Deferred","FirstReleaseCurrent","FirstReleaseDeferred");


                $allowedLanguages = @("en-us",
                "ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee",
                "fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                "ja-jp","kk-kh","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br",
                "pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                "tr-tr","uk-ua")


                #Wait-Debugger
                 Write-Output -InputObject  (New-DynamicOption -Category $category -Name "Bitness" -ExpectedType String -IsRequired $true -permittedValues $allowedBitness) 
                 Write-Output -InputObject  (New-DynamicOption -Category $category -Name "Channel" -ExpectedType String -IsRequired $true -permittedValues $allowedChannel)  
                 Write-Output -InputObject  (New-DynamicOption -Category $category -Name "Languages" -ExpectedType StringArray -IsRequired $false -permittedValues $allowedLanguages)   
                }
    }

    
   
}





Function Install-Package {

    [CmdletBinding()]

    Param(
       [ValidateNotNullOrEmpty()]
       [Parameter(Mandatory=$true)]
       [string]
       $fastPackageReference
        
    )

    $Bitness = $request.Options["Bitness"]
    $Channel = $request.Options["Channel"]
    $Name = $request.Options["Name"]
    $Languages = $request.Options["Languages"]
 

    if($Name.ToLower() -ne "office")
    {
        break 
    }

    
    switch($Branch)
    {
        "Current" {$Channel = "Current"}
        "Deferred" {$Channel = "Business"}
        "FirstReleaseCurrent"{$Channel = "FirstReleaseCurrent"}
        "FirstReleaseDeferred"{$Channel = "FirstReleaseBusiness"}

    }


   $swidObject = @{
                FastPackageReference ="Office Installer";
                Name = "Office Installer";
                Version = New-Object System.Version ("0.1");
                versionScheme  = "MultiPartNumeric";
                summary = "Includes the ODT setup exectuable and configuration xml";
                Source = "Microsoft";      
                fromTrustedSource = $true;         
            }

    $sid = New-SoftwareIdentity @swidObject              
    Write-Output -InputObject $sid     
    
        

    $lang = checkForLanguage
    $destination = "$env:TEMP\OfficeInstaller\"
    $setup = $destination+"setup.exe"
    $config = $destination+"configuration.xml"
    
    New-Item $destination -type directory -force 
    
    Write-Output "Downloading configuration file"
    Start-BitsTransfer "https://officeinstallpackagewest.blob.core.windows.net/officeinstallfiles/configuration.xml" $config

    Write-Output "Downloading setup executable" 
    Start-BitsTransfer "https://officeinstallpackagewest.blob.core.windows.net/officeinstallfiles/setup.exe" $setup

    [xml]$xmlConfig = Get-Content $destination"configuration.xml"

    $langNodes = $xmlConfig.SelectNodes("/Configuration/Add/Product/Language")
    $productNode = $xmlConfig.SelectSingleNode("/Configuration/Add/Product") 
    $add = $xmlConfig.SelectNodes("/Configuration/Add")

    if($Languages.Count -eq 0)
    {
         $langNodes[0].setAttribute("ID",$lang.ToLower())
    }
    else
    {

      $langNodes[0].setAttribute("ID",$Languages)

      if($Languages.Count -gt 1)
      {
        for($index=1; $index -lt $Languages.Count;$index++)
        {
        
            $language = $xmlConfig.CreateElement("Language")
            $language.SetAttribute("ID",$Languages[$index])
            $productNode.AppendChild($language)
            
            
        }
      }
    }

    $add[0].setAttribute("Branch",$Channel)
    $add[0].setAttribute("OfficeClientEdition",$Bitness)

    $xmlConfig.Save($config) 
  
    

    Start-Process -FilePath $setup -ArgumentList "/configure $config"
    
    #Wait-Debugger
    
}

function Find-Package { 
    param(
        [string] $name
    )
        

            $swidObject = @{
                FastPackageReference ="Office Installer";
                Name = "Office Installer";
                Version = New-Object System.Version ("0.1");
                versionScheme  = "MultiPartNumeric";
                summary = "Includes the ODT setup exectuable and configuration xml";
                Source = "Microsoft"; 
                fromTrustedSource = $true;                
            }

            $sid = New-SoftwareIdentity @swidObject              
            Write-Output -InputObject $sid               
}