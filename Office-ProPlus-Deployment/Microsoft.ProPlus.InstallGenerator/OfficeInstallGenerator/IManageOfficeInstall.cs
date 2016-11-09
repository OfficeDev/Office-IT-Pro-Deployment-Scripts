using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGenerator.Model;
using Microsoft.Win32;


namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
     interface IManageOfficeInstall
    {

        Task<OfficeInstallation> CheckForOfficeInstallAsync();

        Task UpdateOffice();

        void UninstallOffice(string installVer = "2016");

    }
}
