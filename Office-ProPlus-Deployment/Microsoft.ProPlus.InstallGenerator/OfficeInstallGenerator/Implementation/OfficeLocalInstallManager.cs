using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
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
                localInstall.Version = GetRegistryValue(officeRegKey, "VersionToReport");
                if (string.IsNullOrEmpty(localInstall.Version)) return localInstall;

                localInstall.Installed = true;
                
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

        public async Task<string> GenerateLocalConfigXml()
        {
            var currentDirectory = Directory.GetCurrentDirectory() + @"\Scripts";
            if (!System.IO.File.Exists(currentDirectory + @"\Generate-ODTConfigurationXML.ps1"))
            {
                currentDirectory = Directory.GetCurrentDirectory() + @"\Project\Scripts";
            }

            var xmlFilePath = Environment.ExpandEnvironmentVariables(@"%temp%\localConfig.xml");

            if (System.IO.File.Exists(xmlFilePath))
            {
                System.IO.File.Delete(xmlFilePath);
            }

            var scriptPath = currentDirectory + @"\Generate-ODTConfigurationXML.ps1";
            var scriptPathTmp = currentDirectory + @"\Tmp-Generate-ODTConfigurationXML.ps1";

            if (!System.IO.File.Exists(scriptPathTmp))
            {
                System.IO.File.Copy(scriptPath, scriptPathTmp, true);
            }

            var scriptUrl = AppSettings.GenerateScriptUrl;

            try
            {
                await Retry.BlockAsync(5, 1, async () =>
                {
                    using (var webClient = new WebClient())
                    {
                        await webClient.DownloadFileTaskAsync(new Uri(scriptUrl), scriptPath);
                    }
                });
            }
            catch (Exception ex) { }

            var n = 1;
            await Retry.BlockAsync(2, 1, async () =>
            {
                var arguments = @"/c Powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle " +
                                @"Hidden -File .\RunGenerateXML.ps1";

                if (n == 2)
                {
                    System.IO.File.Copy(scriptPathTmp, scriptPath, true);
                }

                var p = new Process
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = "cmd",
                        Arguments = arguments,
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        WorkingDirectory = currentDirectory,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    },
                };

                p.Start();
                p.WaitForExit();

                var error = await p.StandardError.ReadToEndAsync();
                if (!string.IsNullOrEmpty(error)) throw (new Exception(error));
                n++;
            });

            await Task.Delay(100);

            if (System.IO.File.Exists(xmlFilePath))
            {
                return System.IO.File.ReadAllText(xmlFilePath);
            }
            return "";
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
                if (regKey.GetValue(property) == null) return "";
                return regKey.GetValue(property).ToString();
            }
            return "";
        }
    }
}
