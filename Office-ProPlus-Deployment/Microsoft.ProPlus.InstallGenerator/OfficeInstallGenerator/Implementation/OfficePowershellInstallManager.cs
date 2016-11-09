using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGenerator.Model;
using Microsoft.Win32;
using System.Diagnostics;
using System.Management;
using System.Windows.Forms.VisualStyles;
using Microsoft.OfficeProPlus.Downloader;
using System.Management.Automation;
using System.Collections.ObjectModel;
using Microsoft.OfficeProPlus.InstallGenerator.Implementation;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    class OfficePowershellInstallManager : IManageOfficeInstall
    {

        //need to add getters/setters for info needed for connection
        public string remoteUser { get; set; }
        public string remoteComputerName { get; set; }
        public string remoteDomain { get; set; }
        public string remotePass { get; set; }
        public string newVersion { get; set; }
        public string newChannel { get; set; }
        public string connectionNamespace { get; set; }
        public ManagementScope scope { get; set; }

        public void InitConnection()
        {
           //implement me

        }

        public async Task<OfficeInstallation> CheckForOfficeInstallAsync()
        {
            var officeInstance = new OfficeInstallation() { Installed = false };
            string readtext = "";
            try
            {
                
                string PSPath = System.IO.Path.GetTempPath() + remoteComputerName + "PowershellAttemptVersion.txt";
                System.IO.File.Delete(PSPath);
                using (var powerShellInstance = System.Management.Automation.PowerShell.Create())
                {
                    powerShellInstance.AddScript(System.IO.Directory.GetCurrentDirectory() + "\\Resources\\FindVersion.ps1 -machineToRun " + remoteComputerName);
                    var async = powerShellInstance.Invoke();
                }
                readtext = System.IO.File.ReadAllText(PSPath);
                readtext = readtext.Trim();

                officeInstance.Version = readtext.Split('\\')[0];
            
                if (!string.IsNullOrEmpty(officeInstance.Version))
                {
                    officeInstance.Installed = true;
                    var currentBaseCDNUrl = readtext.Split('\\')[1];


                    var installFile = await GetOfficeInstallFileXml();
                    if (installFile == null) return officeInstance;

                    var currentBranch = installFile.BaseURL.FirstOrDefault(b => b.URL.Equals(currentBaseCDNUrl) &&
                                                                                !b.Branch.ToLower().Contains("business"));
                    if (currentBranch != null)
                    {
                        officeInstance.Channel = currentBranch.Branch;

                        var latestVersion =
                            await GetOfficeLatestVersion(currentBranch.Branch, OfficeEdition.Office32Bit);
                        officeInstance.LatestVersion = latestVersion;
                    }


                }
            }
            catch (Exception ex)
            {
                using (System.IO.StreamWriter file =
           new System.IO.StreamWriter(System.IO.Path.GetTempPath() + "failure.txt", true))
                {
                    file.WriteLine(ex.Message);
                }
                throw new Exception(ex.Message);
            }
            return officeInstance;
        }

        public Task<string> GenerateConfigXml()
        {
            throw new NotImplementedException();
        }

        public async Task<string> GetOfficeLatestVersion(string branch, OfficeEdition edition)
        {
            var ppDownload = new ProPlusDownloader();
            var latestVersion = await ppDownload.GetLatestVersionAsync(branch, edition);
            return latestVersion;
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

        public async Task<UpdateFiles> GetOfficeInstallFileXml()
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
    }
}
