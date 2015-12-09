using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class MsiGenerator 
    {

        public MsiGeneratorReturn Generate(MsiGeneratorProperties installProperties)
        {

            installProperties.ExecutablePath = Regex.Replace(installProperties.ExecutablePath, ".msi$", ".exe",
                RegexOptions.IgnoreCase);

            installProperties.ExecutablePath = Regex.Replace(installProperties.ExecutablePath, ".exe$", "",
                RegexOptions.IgnoreCase);

            var project = new Project
            {
                Name = installProperties.Name,
                UI = WUI.WixUI_ProgressOnly,
                Dirs = new[]
                {
                    new Dir(installProperties.ProgramFilesPath)
                },
                Binaries = new[]
                {
                    new Binary(new Id("MSOfficeOneClickInstall"), installProperties.ExecutablePath)
                },
                Actions = new WixSharp.Action[]
                {
                    //Install needs silent tag
                    new BinaryFileAction("MSOfficeOneClickInstall", "", Return.check, When.After, Step.InstallFiles,
                        Condition.NOT_Installed)
                    {
                        Execute = Execute.immediate

                    },
                    new BinaryFileAction("MSOfficeOneClickInstall", "/uninstall", Return.check, When.After,
                        Step.InstallFiles, Condition.Installed)
                    {
                        Execute = Execute.immediate
                    }
                },
                GUID = Guid.NewGuid(),
                ControlPanelInfo = { Manufacturer = installProperties.Manufacturer},
                OutFileName = installProperties.ExecutablePath,
            };

            

            Compiler.WixLocation = @"wixTools\";
            Compiler.BuildMsi(project);

            try
            {
                //System.IO.File.Delete(exeFilePath);
            }
            catch { }

            var installDirectory = new MsiGeneratorReturn
            {
                GeneratedFilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            return installDirectory;

        }


    }
}
