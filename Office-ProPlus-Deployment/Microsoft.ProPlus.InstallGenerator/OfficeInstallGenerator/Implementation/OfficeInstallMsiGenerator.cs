using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeInstallGenerator;
using System.IO;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Microsoft.OfficeProPlus.InstallGenerator.Extensions;
using Microsoft.OfficeProPlus.MSIGen;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeInstallMsiGenerator : IOfficeInstallGenerator
    {

        public IOfficeInstallReturn Generate(IOfficeInstallProperties installProperties, string remoteLogPath = "")
        {
            var msiPath = installProperties.ExecutablePath;
            var exePath = Path.GetDirectoryName(installProperties.ExecutablePath) + @"\InstallOfficeProPlus.exe";
            try
            {
                var tmpDir = Environment.ExpandEnvironmentVariables(@"%temp%");

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

                string version = null;
                if (installProperties.Version != null)
                {
                    version = installProperties.Version.ToString();
                }

                var exeReturn = exeGenerator.Generate(installProperties, remoteLogPath);
                var exeFilePath = exeReturn.GeneratedFilePath;

                var msiCreatePath = Regex.Replace(msiPath, ".msi$", "", RegexOptions.IgnoreCase);

                var msiInstallProperties = new MsiGeneratorProperties()
                {
                    MsiPath = msiCreatePath,
                    ExecutablePath = exePath,
                    Manufacturer = "Microsoft Corporation",
                    Name = installProperties.ProductName,
                    ProgramFilesPath = installProperties.ProgramFilesPath,
                    ProgramFiles = new List<string>()
                    {
                        installProperties.ConfigurationXmlPath
                    },
                    ProductId = new Guid(installProperties.ProductId),
                    WixToolsPath = wixDirectory,
                    Version = installProperties.Version,
                    UpgradeCode = new Guid(installProperties.UpgradeCode),
                    Language = installProperties.Language,
                    SourceFilePath = installProperties.SourceFilePath
                };
                var msiGenerator = new MsiGenerator();
                msiGenerator.Generate(msiInstallProperties);

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

        private MsiDirectory GetSourceFiles(string sourcePath, string version = null, OfficeClientEdition officeClientEdition = OfficeClientEdition.Office32Bit)
        {
            var lstReturn = new MsiDirectory
            {
                RootPath = sourcePath,
                RelativePath = "",
            };

            var dirInfo = new DirectoryInfo(sourcePath);
            var topFiles = dirInfo.GetFiles("*.*", SearchOption.TopDirectoryOnly);
            foreach (var file in topFiles)
            {
                lstReturn.MsiFiles.Add(new MsiFile() {Path = file.FullName});
            }

            foreach (var directory in dirInfo.GetDirectories())
            {
                GetMsiDirectory(lstReturn, directory.FullName);
            }

            //foreach (var sourceFile in sourceFiles)
            //{
            //    if (!string.IsNullOrEmpty(version))
            //    {
            //        if (!(sourceFile.FullName.ToLower().Contains(version.ToLower()) ||
            //            sourceFile.Name.ToLower() == "v32.cab" ||
            //            sourceFile.Name.ToLower() == "v64.cab"))
            //        {
            //            continue;
            //        }
            //    }

            //    if (officeClientEdition == OfficeClientEdition.Office32Bit)
            //    {
            //        if (sourceFile.Name.ToLower().Contains(".x64."))
            //        {
            //            continue;
            //        }
            //    }
            //    else
            //    {
            //        if (sourceFile.Name.ToLower().Contains(".x86."))
            //        {
            //            continue;
            //        }
            //    }

            //    lstReturn.Add(sourceFile.FullName);
            //}

            return lstReturn;
        }

        public MsiDirectory GetMsiDirectory(MsiDirectory parent, string directoryPath)
        {
            var msiDir = new MsiDirectory
            {
                RootPath = parent.RootPath
            };
            parent.MsiDirectories.Add(msiDir);

            var dirInfo = new DirectoryInfo(directoryPath);

            var relativePath = Regex.Replace(dirInfo.FullName, "^" + parent.RootPath.Replace(@"\", @"\\"), "", RegexOptions.IgnoreCase);
            msiDir.RelativePath = relativePath;
            msiDir.Name = dirInfo.Name;

             var topFiles = dirInfo.GetFiles("*.*", SearchOption.TopDirectoryOnly);

            foreach (var file in topFiles)
            {
                msiDir.MsiFiles.Add(new MsiFile() { Path = file.FullName });
            }

            foreach (var directory in dirInfo.GetDirectories())
            {
                GetMsiDirectory(msiDir, directory.FullName);
            }

            return msiDir;
        }




        public void InstallOffice(string configurationXml)
        {
            throw new NotImplementedException();
        }
    }
}
