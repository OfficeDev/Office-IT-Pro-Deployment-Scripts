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
    class OfficeRemoteWmiInstallManager : IManageOfficeInstall
    {
        public Task<OfficeInstallation> CheckForOfficeInstallAsync()
        {
            throw new NotImplementedException();
        }

        public Task<string> GenerateConfigXml()
        {
            throw new NotImplementedException();
        }

        public Task<string> GetOfficeLatestVersion(string branch, OfficeEdition edition)
        {
            throw new NotImplementedException();
        }

        public string GetRegistryValue(RegistryKey regKey, string property)
        {
            throw new NotImplementedException();
        }

        public void UninstallOffice(string installVer = "2016")
        {
            throw new NotImplementedException();
        }

        public Task UpdateOffice()
        {
            throw new NotImplementedException();
        }
    }
}
