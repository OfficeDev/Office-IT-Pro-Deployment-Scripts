using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using Microsoft.Deployment.WindowsInstaller;
using WixSharp;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class MsiGenerator 
    {

        public MsiGeneratorReturn Generate(MsiGeneratorProperties installProperties)
        {

            installProperties.ExecutablePath = Regex.Replace(installProperties.ExecutablePath, ".msi$", ".exe",
                RegexOptions.IgnoreCase);

            var exePath = installProperties.ExecutablePath;

            installProperties.ExecutablePath = Regex.Replace(installProperties.ExecutablePath, ".exe$", ".msi",
                RegexOptions.IgnoreCase);

            //var project = new Project
            //{
            //    Name = installProperties.Name,
            //    UI = WUI.WixUI_ProgressOnly,
            //    Dirs = new[]
            //    {
            //        new Dir(installProperties.ProgramFilesPath)
            //        {

            //        }
            //    },
            //    Binaries = new[]
            //    {
            //        new Binary(new Id("MSOfficeOneClickInstall"), exePath)
            //    },
            //    Actions = new WixSharp.Action[]
            //    {
            //        new ElevatedManagedAction("InstallOffice", "", Return.check, When.After, 
            //                                  Step.InstallFiles, Condition.NOT_Installed)
            //    },
            //    GUID = Guid.NewGuid(),
            //    ControlPanelInfo = { Manufacturer = installProperties.Manufacturer},
            //    OutFileName = installProperties.ExecutablePath,
            //};


            var project = new Project
            {
                Name = "Microsoft Office 365 ProPlus Installer",
                UI = WUI.WixUI_ProgressOnly,
                Dirs = new[]
                {
                    new Dir(@"%ProgramFiles%\MS\Microsoft Office 365 ProPlus Installer")
                },
                Binaries = new[]
                {
                    new Binary(new Id("MSOfficeOneClickInstall"), exePath)
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
                ControlPanelInfo = { Manufacturer = "Microsoft Corporation" },
                OutFileName = installProperties.ExecutablePath,
            };

            Compiler.WixSdkLocation = @"wixTools\sdk";
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


        public class CustomActions
        {
            [CustomAction]
            public static ActionResult InstallOffice(Session session)
            {
                session.Log("Begin MyAction Hello World");
                return ActionResult.Success;
            }
        }






    }
}
