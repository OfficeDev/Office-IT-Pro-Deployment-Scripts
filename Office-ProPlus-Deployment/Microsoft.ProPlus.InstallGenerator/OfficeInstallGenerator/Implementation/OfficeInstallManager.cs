using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGenerator.Model;
using Microsoft.Win32;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeInstallManager : IManageOfficeInstall
    {
        OfficeLocalInstallManager LocalInstall = new OfficeLocalInstallManager();
        OfficeWmiInstallManager WmiInstall = new OfficeWmiInstallManager();
        OfficePowershellInstallManager PowershellInstall = new OfficePowershellInstallManager();

        private bool isLocal { get; set; }

        public OfficeInstallManager()
        {
            isLocal = true; 
        }

        public OfficeInstallManager(string[] computerInfo)
        {
            isLocal = false;

            WmiInstall.remoteUser = computerInfo[0];
            WmiInstall.remoteComputerName = computerInfo[1];
            WmiInstall.remoteDomain = computerInfo[2];
            WmiInstall.remotePass = computerInfo[3];



            //need to set Powershell info now..

            try
            {
                //WmiInstall.initConnection(); 
            }
            catch (Exception)
            {

                //try
                //{
                //    PowershellInstall.initConnection();
                //}
                //catch(Exception)

                //{
                //    //let user know that neither wmi nor powershell could connect 
                //}
            }


            try
            {
                WmiInstall.initConnection();
                PowershellInstall.initConnection();
            }
            catch (Exception)
            {//add error handling
            }
        }

        public async Task<OfficeInstallation> CheckForOfficeInstallAsync()
        {
            var result = new OfficeInstallation();

            try
            {
               
                if (isLocal)
                {
                    result = await LocalInstall.CheckForOfficeInstallAsync();
               
                }
                else
                {

                    try
                    {
                        result = await WmiInstall.CheckForOfficeInstallAsync();

                    }
                    catch (Exception)
                    {
                        try
                        {
                            result = await PowershellInstall.CheckForOfficeInstallAsync();
                 
                        }
                        catch (Exception) { }
                   
                    }
                }
            }
            catch (Exception)
            {

            }
          
            return result;
            
           
        }

        public Task<string> GenerateConfigXml()
        {            
            var result = LocalInstall.GenerateConfigXml();
            return result;

        }

        public Task<string> GetOfficeLatestVersion(string branch, OfficeEdition edition)
        {
            
            var result = LocalInstall.GetOfficeLatestVersion(branch, edition);
            return result;
          
      
        }

        public string GetRegistryValue(RegistryKey regKey, string property)
        {

            var result = "";
            try
            {
                if (isLocal)
                {
                    result = LocalInstall.GetRegistryValue(regKey, property);
                }
                else
                {
                    try
                    {
                        result = WmiInstall.GetRegistryValue(regKey, property);
                    }
                    catch (Exception)
                    {
                        try
                        {
                            result = PowershellInstall.GetRegistryValue(regKey, property);
                        }
                        catch (Exception)
                        {
                            //add error handling 
                        }
                    }
                }
            }
            catch (Exception)
            { 
                //add error handling
            }
            return result; 
        }

        public void UninstallOffice(string installVer = "2016")
        {
            LocalInstall.UninstallOffice(installVer);
        }

        public Task UpdateOffice()
        {
            var result = LocalInstall.UpdateOffice();
            return result; 
        }
    }
}
