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
using OfficeInstallGenerator;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeLocalInstallManager : IManageOfficeInstall
    {

        public async Task<OfficeInstallation> CheckForOfficeInstallAsync()
        {
            var localInstall = new OfficeInstallation()
            {
                Installed = false
            };

            var officeRegKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            if (officeRegKey == null)
            {
                officeRegKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\ClickToRun\Configuration");
                if (officeRegKey == null)
                {
                    officeRegKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration");
                }
            }
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

        public async Task<string> GenerateConfigXml()
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

            var configXml = "";

            if (System.IO.File.Exists(xmlFilePath))
            {
                configXml = System.IO.File.ReadAllText(xmlFilePath);
            }

            try
            {
                var installOffice = new InstallOffice();
                var updateUrl = installOffice.GetBaseCdnUrl();
                if (!string.IsNullOrEmpty(updateUrl))
                {
                    var pd = new ProPlusDownloader();
                    var channelName = await pd.GetChannelNameFromUrlAsync(updateUrl, OfficeEdition.Office32Bit);
                    if (!string.IsNullOrEmpty(configXml) && !string.IsNullOrEmpty(channelName))
                    {
                        configXml = installOffice.SetUpdateChannel(configXml, channelName);
                    }
                }
            } catch { }

            return configXml;
        }

        public async Task<string> GetOfficeLatestVersion(string branch, OfficeEdition edition)
        {
            var ppDownload = new ProPlusDownloader();
            var latestVersion = await ppDownload.GetLatestVersionAsync(branch, edition);
            return latestVersion;
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

        public void UninstallOffice(string installVer = "2016")
        {
            
            const string configurationXml = "<Configuration><Remove All=\"TRUE\"/><Display Level=\"Full\" /></Configuration>";

            var tmpPath = Environment.ExpandEnvironmentVariables("%temp%");
            var embededExeFiles = EmbeddedResources.GetEmbeddedItems(tmpPath, @"\.exe$");

            //NOTE: Have this function determine if 2013 ProPlus or 2016 ProPlus is installed and then use the right ODT version            
            var installExe = tmpPath + @"\" + embededExeFiles.FirstOrDefault(f => f.ToLower().Contains("2016"));
            if (installVer == "2013")
            {
                //If 2013 then get the 2013 ODT version
                installExe = tmpPath + @"\" + embededExeFiles.FirstOrDefault(f => f.ToLower().Contains("2013"));
            }
            var xmlPath = tmpPath + @"\configuration.xml";

            if (System.IO.File.Exists(xmlPath)) System.IO.File.Delete(xmlPath);
            System.IO.File.WriteAllText(xmlPath, configurationXml);

            var p = new Process
            {
                StartInfo = new ProcessStartInfo()
                {
                    FileName = installExe,
                    Arguments = "/configure " + tmpPath + @"\configuration.xml",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                },
            };
            p.Start();
            p.WaitForExit();

            var error = p.StandardError.ReadToEnd();
            if (!string.IsNullOrEmpty(error)) throw (new Exception(error));

            if (System.IO.File.Exists(xmlPath)) System.IO.File.Delete(xmlPath);

            foreach (var exeFilePath in embededExeFiles)
            {
                try
                {
                    if (System.IO.File.Exists(tmpPath + @"\" + exeFilePath))
                    {
                        System.IO.File.Delete(tmpPath + @"\" + exeFilePath);
                    }
                }
                catch { }
            }
        }
    
        public string GetRegistryValue(RegistryKey regKey, string property)
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
