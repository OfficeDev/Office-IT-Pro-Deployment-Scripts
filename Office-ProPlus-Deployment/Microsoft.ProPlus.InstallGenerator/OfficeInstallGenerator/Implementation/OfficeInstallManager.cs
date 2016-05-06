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
        OfficeRemoteInstallManager RemoteInstall = new OfficeRemoteInstallManager(); 

        private bool isLocal { get; set; }

        public OfficeInstallManager()
        {
            isLocal = true; 
        }

        public OfficeInstallManager(string computerName)
        {
            isLocal = false; 

        }

        public Task<OfficeInstallation> CheckForOfficeInstallAsync()
        {
           
            if (isLocal)
            {
                var result = LocalInstall.CheckForOfficeInstallAsync();
                return result;

            }
            else
            {
                var result = RemoteInstall.CheckForOfficeInstallAsync();
                return result;

            }


        }

        public Task<string> GenerateConfigXml()
        {            
            var result = LocalInstall.GenerateConfigXml();
            return result;

        }

        public Task<string> GetOfficeLatestVersion(string branch, OfficeEdition edition)
        {
            //TODO implement remote 
            if (isLocal)
            {
                var result = LocalInstall.GetOfficeLatestVersion(branch, edition);
                return result;
            }
            else
            {
                var result = RemoteInstall.GetOfficeLatestVersion(branch, edition);
                return result;
            }
      
        }

        public string GetRegistryValue(RegistryKey regKey, string property)
        {
            var result = LocalInstall.GetRegistryValue(regKey, property);
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
