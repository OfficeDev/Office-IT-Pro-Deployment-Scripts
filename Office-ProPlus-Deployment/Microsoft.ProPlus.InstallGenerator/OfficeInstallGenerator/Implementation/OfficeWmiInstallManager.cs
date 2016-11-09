using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGenerator.Model;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.Win32;
using System.Management;
using System.Runtime.ExceptionServices;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    class OfficeWmiInstallManager : IManageOfficeInstall
    {

        public string remoteUser { get; set;}
        public string remoteComputerName { get; set;}
        public string remoteDomain { get; set;}
        public string remotePass { get; set;}
        public string newVersion { get; set;}
        public string newChannel { get; set;}
        public string connectionNamespace { get; set; }
        public ManagementScope scope { get; set;}


        public async Task InitConnection()
        {

            var timeOut = new TimeSpan(0, 5, 0);
            var options = new ConnectionOptions
            {
                Impersonation = ImpersonationLevel.Impersonate,
                Timeout = timeOut
            };

            if (remoteDomain != null)
            {
                options.Authority = "NTLMDOMAIN:" + remoteDomain.Trim();
            }

            if (remoteUser != null && remotePass != null)
            {
                options.Username = remoteUser.Trim();
                options.Password = remotePass.Trim();
            }
            
            scope = new ManagementScope("\\\\" + remoteComputerName.Trim() + connectionNamespace, options)
            {
                Options = {EnablePrivileges = true}
            };

            ExceptionDispatchInfo exception = null;
            try
            {
               await Task.Run(() => {
                    try
                    {
                        scope.Connect();
                    }
                    catch(Exception ex)
                    {
                       exception = ExceptionDispatchInfo.Capture(ex);
                    }
              });
              exception?.Throw();
            }
            catch (Exception)
            {
                exception = null;
                try
                {
                    await Task.Run(() =>
                    {
                        try
                        {
                            scope.Connect();
                        }
                        catch (Exception ex)
                        {
                            exception = ExceptionDispatchInfo.Capture(ex);
                        }
                    });

                    exception?.Throw();
                }
                catch (Exception)
                {
                    throw (new Exception("Cannot connect to client"));
                }
            }

        }

        public async Task<OfficeInstallation> CheckForOfficeInstallAsync()
        {         
            var officeInstance = new OfficeInstallation() { Installed = false };
            var officeRegPathKey = @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration";

            officeInstance.Version = await GetRegistryValue(officeRegPathKey, "VersionToReport", "GetStringValue");

            if (string.IsNullOrEmpty(officeInstance.Version))
                {
                    officeRegPathKey = @"SOFTWARE\Microsoft\Office\16.0\ClickToRun\Configuration";
                officeInstance.Version = await GetRegistryValue(officeRegPathKey, "VersionToReport", "GetStringValue");

                    if (string.IsNullOrEmpty(officeInstance.Version))
                    {
                        officeRegPathKey = @"SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration";
                        officeInstance.Version = await GetRegistryValue(officeRegPathKey, "VersionToReport", "GetStringValue");

                    }

                }

                if(!string.IsNullOrEmpty(officeInstance.Version))
                {
                    officeInstance.Installed = true;
                    var currentBaseCDNUrl = await GetRegistryValue(officeRegPathKey, "CDNBaseUrl", "GetStringValue");


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

        private Task<string> GenerateConfigXml()
        {
            throw new NotImplementedException();
        }

        public async Task<string> GetOfficeLatestVersion(string branch, OfficeEdition edition)
        {
            var ppDownload = new ProPlusDownloader();
            var latestVersion = await ppDownload.GetLatestVersionAsync(branch, edition);
            return latestVersion;
        }

        private async Task<string> GetOfficeC2RPath()
        {
            
            await Task.Run(() => {
                var path = @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration";
                var path15 = @"SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration";

                 var result =  GetRegistryValue(path15, "ClientFolder", "GetStringValue").ToString();

                if (string.IsNullOrEmpty(result))
                {
                    result =  GetRegistryValue(path, "ClientFolder", "GetStringValue").ToString();

                }

                return result;
            
            });
            return null;
        }
        

        private async Task<string> GetRegistryValue(string regKey, string valueName, string getmethParam)
        {

            var regValue = "";

            await Task.Run(() =>
            {


                ManagementClass registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);
                ManagementBaseObject inParams = registry.GetMethodParameters(getmethParam);

                inParams["hDefKey"] = 0x80000002;
                inParams["sSubKeyName"] = regKey;
                inParams["sValueName"] = valueName;

                ManagementBaseObject outParams = registry.InvokeMethod(getmethParam, inParams, null);

                try
                {
                    if (outParams.Properties["sValue"].Value.ToString() != null)
                    {
                        regValue = outParams.Properties["sValue"].Value.ToString();
                    }
               } 
                catch (Exception)
                {
                    regValue = null;
                }



            });


            return regValue;
        }

        public void UninstallOffice(string installVer = "2016")
        {
            throw new NotImplementedException();
        }


        public async Task UpdateOffice()
        {


            await InitConnection();
            var currentInstall = await CheckForOfficeInstallAsync();

        }
    }
}
