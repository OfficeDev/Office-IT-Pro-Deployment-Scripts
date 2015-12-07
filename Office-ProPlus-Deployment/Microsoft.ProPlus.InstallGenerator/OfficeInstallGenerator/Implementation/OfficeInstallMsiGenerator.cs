﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeInstallGenerator;
using WixSharp;
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
                    new Binary(new Id("MSOfficeOneClickInstall"), exeFilePath)
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
                ControlPanelInfo = {Manufacturer = "Microsoft Corporation"},
                OutFileName = installProperties.ExecutablePath,
            };

            Compiler.WixLocation = @"wixTools\";
            Compiler.BuildMsi(project);

            try
            {
                //System.IO.File.Delete(exeFilePath);
            }
            catch { }

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
