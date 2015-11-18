using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeInstallGenerator;
using WixSharp;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeInstallMsiGenerator : IOfficeInstallGenerator
    {



        //TODO : Add function that gets the bitness of the install...for now assume
        public IOfficeInstallReturn Generate(IOfficeInstallProperties installProperties)
        {
            var exeGenerator = new OfficeInstallExecutableGenerator();
            var exeReturn = exeGenerator.Generate(installProperties);

            var exeFilePath = exeReturn.GeneratedFilePath;

            var project = new Project()
            {

                Name = "Microsoft Office 365 ProPlus Installer",
                UI = WUI.WixUI_Minimal,

                Dirs = new[]
            {
                new Dir(@"%ProgramFiles%\MS\Microsoft Office 365 ProPlus Installer")
            },

                Binaries = new[]
            {
                new Binary(new Id("MSOfficeOneClickInstall"),exeFilePath)
            },


                Actions = new WixSharp.Action[]
            {
                //Install needs silent tag
                new BinaryFileAction("MSOfficeOneClickInstall","", Return.check, When.After, Step.InstallFiles, Condition.NOT_Installed)
                {
                        Execute = Execute.immediate
                },
                 new BinaryFileAction("MSOfficeOneClickInstall","/uninstall", Return.check, When.After, Step.InstallFiles, Condition.Installed)
                {
                        Execute = Execute.immediate
                }
            }

            };



            project.GUID = Guid.NewGuid();

            Compiler.BuildMsi(project);

            OfficeInstallReturn installDirectory = new OfficeInstallReturn();

            installDirectory.GeneratedFilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            return installDirectory;

        }

    }
}
