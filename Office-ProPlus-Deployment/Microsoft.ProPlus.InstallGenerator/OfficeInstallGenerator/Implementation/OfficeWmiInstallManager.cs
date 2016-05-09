using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGenerator.Model;
using Microsoft.Win32;
using System.Management;
namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    class OfficeWmiInstallManager : IManageOfficeInstall
    {

        public string remoteUser { get; set;}
        public string remoteComputerName { get; set;}
        public string remoteDomain { get; set;}
        public string remotePass { get; set;}


        public void initConnection()
        {
            var computerName = remoteComputerName;
            var password = remotePass;

            ConnectionOptions options = new ConnectionOptions();
            options.Authority = "NTLMDOMAIN:" + remoteDomain;
            options.Username = remoteUser;
            options.Password = remotePass;
           

            //ManagementScope scope = new ManagementScope("\\\\"+computerName+"\\root\\cimv2", options);
            ManagementScope scope = new ManagementScope(@"\"+remoteComputerName+@"\root\cimv2", options);
            //scope.Options.EnablePrivileges = true;
            scope.Options.Impersonation = ImpersonationLevel.Impersonate;
            scope.Connect();

            //Query system for Operating System information
            ObjectQuery query = new ObjectQuery(
                "SELECT * FROM Win32_OperatingSystem");
            ManagementObjectSearcher searcher =
                new ManagementObjectSearcher(scope, query);


            ManagementObjectCollection queryCollection = searcher.Get();
            foreach (ManagementObject m in queryCollection)
            {
                // Display the remote computer information
                Console.WriteLine("Computer Name : {0}",
                    m["csname"]);
                Console.WriteLine("Windows Directory : {0}",
                    m["WindowsDirectory"]);
                Console.WriteLine("Operating System: {0}",
                    m["Caption"]);
                Console.WriteLine("Version: {0}", m["Version"]);
                Console.WriteLine("Manufacturer : {0}",
                    m["Manufacturer"]);
            }
        }

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
