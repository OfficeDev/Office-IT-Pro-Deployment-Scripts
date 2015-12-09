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


        //TODO : Add function that gets the bitness of the install...for now assume
        public IOfficeInstallReturn Generate(IOfficeInstallProperties installProperties)
        {
            var exeGenerator = new OfficeInstallExecutableGenerator();

            installProperties.ExecutablePath = Regex.Replace(installProperties.ExecutablePath, ".msi$", ".exe",
                RegexOptions.IgnoreCase);

            var exeReturn = exeGenerator.Generate(installProperties);
            var exeFilePath = exeReturn.GeneratedFilePath;

            installProperties.ExecutablePath = Regex.Replace(installProperties.ExecutablePath, ".exe$", "",
                RegexOptions.IgnoreCase);


            var msiGenerator = new MsiGenerator();

            msiGenerator.Generate(new MsiGeneratorProperties()
            {
                ExecutablePath = exeFilePath,
                Manufacturer = "Microsoft Corporation",
                Name = "Microsoft Office 365 ProPlus Installer",
                ProgramFilesPath = @"%ProgramFiles%\MS\Microsoft Office 365 ProPlus Installer"
            });

            var installDirectory = new OfficeInstallReturn
            {
                GeneratedFilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            return installDirectory;

        }

        public void InstallOffice(string configurationXml)
        {
            throw new NotImplementedException();
        }
    }
}
