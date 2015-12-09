using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeInstallGenerator;
using System.IO;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeInstallMsiGenerator : IOfficeInstallGenerator
    {

        public IOfficeInstallReturn Generate(IOfficeInstallProperties installProperties)
        {
            var msiPath = installProperties.ExecutablePath;
            var exePath = Path.GetDirectoryName(installProperties.ExecutablePath) + @"\InstallOfficeProPlus.exe";
            try
            {
                var exeGenerator = new OfficeInstallExecutableGenerator();
                installProperties.ExecutablePath = exePath;

                var exeReturn = exeGenerator.Generate(installProperties);
                var exeFilePath = exeReturn.GeneratedFilePath;

                var msiCreatePath = Regex.Replace(msiPath, ".msi$", "", RegexOptions.IgnoreCase);

                var msiGenerator = new MsiGenerator();
                msiGenerator.Generate(new MsiGeneratorProperties()
                {
                    MsiPath = msiCreatePath,
                    ExecutablePath = exePath,
                    Manufacturer = "Microsoft Corporation",
                    Name = "Microsoft Office 365 ProPlus Installer",
                    ProgramFilesPath = @"%ProgramFiles%\Microsoft Office 365 ProPlus Installer",
                    ProgramFiles = new List<string>()
                    {
                        installProperties.ConfigurationXmlPath
                    }
                });

                var installDirectory = new OfficeInstallReturn
                {
                    GeneratedFilePath = msiPath
                };

                return installDirectory;
            }
            finally
            {
                try
                {
                    if (File.Exists(exePath))
                    {
                        File.Delete(exePath);
                    }
                }
                catch { }
            }
        }

        public void InstallOffice(string configurationXml)
        {
            throw new NotImplementedException();
        }
    }
}
