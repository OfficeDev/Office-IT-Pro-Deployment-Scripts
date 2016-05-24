using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.InstallGenerator.Model;
using Microsoft.Win32;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeInstallManager : IManageOfficeInstall
    {
        OfficeLocalInstallManager LocalInstall = new OfficeLocalInstallManager();
        OfficeWmiInstallManager WmiInstall = new OfficeWmiInstallManager();
        OfficePowershellInstallManager PowershellInstall = new OfficePowershellInstallManager();

        private string[] computerInfo { get; set; }

        private bool isLocal { get; set; }

        public OfficeInstallManager()
        {
            isLocal = true; 

        }

        public OfficeInstallManager(string[] computerInfo)
        {
            isLocal = false;
            this.computerInfo = computerInfo;              

        }

        public  async Task  initConnections()
        {
            WmiInstall.remoteUser = computerInfo[0];
            WmiInstall.remoteComputerName = computerInfo[2];
            WmiInstall.remoteDomain = computerInfo[3];
            WmiInstall.remotePass = computerInfo[1];
            WmiInstall.connectionNamespace = "\\root\\cimv2"; 

            //need to set Powershell info now..           
            PowershellInstall.remoteUser = computerInfo[0];
            PowershellInstall.remoteComputerName = computerInfo[2];
            PowershellInstall.remoteDomain = computerInfo[3];
            PowershellInstall.remotePass = computerInfo[1];


            try
            {
               await WmiInstall.initConnection();
            }
            catch (Exception)
            {
                //
                try
                {
                    PowershellInstall.initConnection();
                }
                catch (Exception)

                {
                    throw (new Exception("Cannot find client"));
                }
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


        public void UninstallOffice(string installVer = "2016")
        {
            LocalInstall.UninstallOffice(installVer);
        }

        public Task UpdateOffice()
        {

            if (isLocal)
            {
                  return LocalInstall.UpdateOffice();
            }
            else
            {


                return WmiInstall.UpdateOffice();
            }

        }

        public async void ChangeOfficeChannel(List<string> updateInfo)
        {


         

   


        }

        //public Task UpdateOffice(List<string> updateInfo)
        //{
  
        //}
    }
}
