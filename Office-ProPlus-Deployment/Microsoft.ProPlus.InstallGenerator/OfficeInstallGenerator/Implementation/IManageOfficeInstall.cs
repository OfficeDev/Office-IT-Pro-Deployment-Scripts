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

        Task<string> GenerateConfigXml();

        Task<string> GetOfficeLatestVersion(string branch, OfficeEdition edition);

        void UninstallOffice(string installVer = "2016");

        string GetRegistryValue(RegistryKey regKey, string property);
    }
}
