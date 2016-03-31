using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGenerator.Model;
using Microsoft.Win32;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeLocalInstallManager
    {

        public async Task<OfficeLocalInstall> CheckForOfficeLocalInstallAsync()
        {
            var localInstall = new OfficeLocalInstall()
            {
                Installed = false
            };

            var officeRegKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            if (officeRegKey != null)
            {
                localInstall.Installed = true;
                localInstall.Version = GetRegistryValue(officeRegKey, "VersionToReport");

                var currentBaseCDNUrl = GetRegistryValue(officeRegKey, "CDNBaseUrl");
 
                var installFile = await GetOfficeInstallFileXml();
                if (installFile == null) return localInstall;

                var currentBranch = installFile.BaseURL.FirstOrDefault(b => b.URL.Equals(currentBaseCDNUrl) &&
                                                                            !b.Branch.ToLower().Contains("business"));
                if (currentBranch != null)
                {
                    localInstall.Channel = currentBranch.Branch;

                    var latestVersion = await GetOfficeLatestVersion(currentBranch.Branch, OfficeEdition.Office32Bit);
                    localInstall.LatestVersion = latestVersion;
                }
            }

            return localInstall;
        }

        public async Task UpdateOffice()
        {
            var installOffice = new InstallOffice();
           



        }

        private async Task<string> GetOfficeLatestVersion(string branch, OfficeEdition edition)
        {
            var ppDownload = new ProPlusDownloader();
            var latestVersion = await ppDownload.GetLatestVersionAsync(branch, edition);
            return latestVersion;
        }

        private async Task<UpdateFiles> GetOfficeInstallFileXml()
        {
            var ppDownload = new ProPlusDownloader();
            var installFiles = await ppDownload.DownloadCabAsync();
            if (installFiles != null)
            {
                var installFile = installFiles.FirstOrDefault();
                if (installFile != null)
                {
                    return installFile;
                }
            }
            return null;
        }

        private string GetRegistryValue(RegistryKey regKey, string property)
        {
            if (regKey != null)
            {
                return regKey.GetValue(property).ToString();
            }
            return "";
        }
    }
}
