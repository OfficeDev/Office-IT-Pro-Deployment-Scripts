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
                var tmpDir = Environment.ExpandEnvironmentVariables("%temp%");

                var wixDirectory = tmpDir + @"\wixTools";
                var wixZip = ZipExtractor.AssemblyDirectory + @"\wixTools.zip";
                if (!File.Exists(wixZip))
                {
                    var projectPath = Directory.GetCurrentDirectory() + @"\Project\wixTools.zip";
                    if (File.Exists(projectPath))
                    {
                        wixZip = projectPath;
                    }
                }

                if (!Directory.Exists(wixDirectory))
                {
                    ZipExtractor.Extract(wixZip, tmpDir);
                }

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
                    },
                    ProductId = new Guid("8AA11E8A-A882-45CC-B52C-80149B4CF47A"),
                    WixToolsPath = wixDirectory
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
