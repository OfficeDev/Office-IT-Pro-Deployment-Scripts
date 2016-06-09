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

        public void initConnection()
        {
           //implement me

        }

        public async Task<OfficeInstallation> CheckForOfficeInstallAsync()
        {
            var officeInstance = new OfficeInstallation() { Installed = false };
            string readtext = "";
            try
            {
                string PSPath = System.IO.Directory.GetCurrentDirectory() + "\\Resources\\"+remoteComputerName+"PowershellAttemptVersion.txt";
                System.IO.File.Delete(PSPath);
                Process p = new Process();
                p.StartInfo.FileName = "Powershell.exe";                                //replace path to use local path                            switch out arguments so your program throws in the necessary args
                p.StartInfo.Arguments = @"-ExecutionPolicy Bypass -NoExit -Command ""& {& '" + System.IO.Directory.GetCurrentDirectory() + "\\Resources\\FindVersion.ps1' -machineToRun " + remoteComputerName + "}\"";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.CreateNoWindow = true;
                p.Start();
                p.WaitForExit();
                p.Close();
                readtext = System.IO.File.ReadAllText(PSPath);
                readtext = readtext.Trim();

                officeInstance.Version = readtext.Split('\\')[0];
            }
            catch (Exception)
            {

            }
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

                    var latestVersion = await GetOfficeLatestVersion(currentBranch.Branch, OfficeEdition.Office32Bit);
                    officeInstance.LatestVersion = latestVersion;
                }


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
