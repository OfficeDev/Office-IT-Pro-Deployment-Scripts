using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.InstallGenerator.Model;
using Microsoft.Win32;
using OfficeInstallGenerator;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeInstallManager : IManageOfficeInstall
    {
        OfficeLocalInstallManager LocalInstall = new OfficeLocalInstallManager();
        OfficeWmiInstallManager WmiInstall = new OfficeWmiInstallManager();
        OfficePowershellInstallManager PowershellInstall = new OfficePowershellInstallManager();

        private string _computerName = null;
        private string _domain = null;
        private string _username = null;
        private string _password = null;
        private ConnectionType _connectionType;


        private bool isLocal { get; set; }
        

        public OfficeInstallManager()
        {
            isLocal = true; 
        }

        public OfficeInstallManager(string computerName, string domain = null, string username = null, string password = null)
        {
            isLocal = false;
            _computerName = computerName;
            _domain = domain;
            _username = username;
            _password = password;
        }

        public async Task InitConnections()
        {
            WmiInstall.remoteUser = _username;
            WmiInstall.remoteComputerName = _computerName;
            WmiInstall.remoteDomain = _domain;
            WmiInstall.remotePass = _password;
            WmiInstall.connectionNamespace = "\\root\\cimv2"; 

            //need to set Powershell info now..           
            PowershellInstall.remoteUser = _username;
            PowershellInstall.remoteComputerName = _computerName;
            PowershellInstall.remoteDomain = _domain;
            PowershellInstall.remotePass = _password;

            ExceptionDispatchInfo exception = null;
            try
            {
                await WmiInstall.InitConnection();
                _connectionType = ConnectionType.WMI;
                return;
                //PowershellInstall.InitConnection();
                //_connectionType = ConnectionType.PowerShell;
                //return;
            }
            catch (Exception ex)
            {
                exception = ExceptionDispatchInfo.Capture(ex);
            }

            try
            {
                PowershellInstall.InitConnection();
                _connectionType = ConnectionType.PowerShell;
                return;
                //await WmiInstall.InitConnection();
                //_connectionType = ConnectionType.WMI;
                //return;
            }
            catch (Exception ex)
            {
                exception = ExceptionDispatchInfo.Capture(ex);
            }

            exception?.Throw();
        }

        public async Task<OfficeInstallation> CheckForOfficeInstallAsync()
        {
            var result = new OfficeInstallation();

            if (isLocal)
            {
                result = await LocalInstall.CheckForOfficeInstallAsync();
            }
            else
            {
                switch (_connectionType)
                {
                    case ConnectionType.WMI:
                        result = await WmiInstall.CheckForOfficeInstallAsync();
                        break;
                    case ConnectionType.PowerShell:
                        result = await PowershellInstall.CheckForOfficeInstallAsync();
                        break;
                    default:
                        throw new Exception("Connection Unknown");
                }
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
