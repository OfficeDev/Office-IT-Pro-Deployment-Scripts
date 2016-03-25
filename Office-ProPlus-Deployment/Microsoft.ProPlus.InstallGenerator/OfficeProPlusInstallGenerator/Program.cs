using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.InstallGenerator.Implementation;
using OfficeInstallGenerator;
using System.IO.Compression; 

namespace OfficeProPlusInstallGenerator
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                var installOffice = new InstallOffice();
                installOffice.RunProgram();

                var xmlConfiguration = "";

                var cmdArgs = CmdArguments.GetArguments();
                var xmlArg = cmdArgs.FirstOrDefault(a => a.Name.ToUpper() == "XML");
                if (xmlArg != null) xmlConfiguration = xmlArg.Value;

                Console.WriteLine("Office 365 ProPlus Install Executable Generator");
                Console.WriteLine();

                if (string.IsNullOrEmpty(xmlConfiguration))
                {
                    Console.Write("Configuration Xml File Path: ");
                    xmlConfiguration = Console.ReadLine();
                    Console.WriteLine();
                }

                if (!File.Exists(xmlConfiguration))
                {
                    throw (new Exception("File Does Not Exist: " + xmlConfiguration));
                }

                var p = new OfficeInstallMsiGenerator();
                p.Generate(new OfficeInstallProperties()
                {
                    OfficeVersion = OfficeVersion.Office2016,
                    ConfigurationXmlPath = xmlConfiguration,
                    SourceFilePath = null
                });
            }
            catch (Exception ex)
            {
                var backColor = Console.BackgroundColor;
                var textColor = Console.ForegroundColor;

                Console.BackgroundColor = ConsoleColor.Red;
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("ERROR: " + ex.Message);
                Console.BackgroundColor = backColor;
                Console.ForegroundColor = textColor;
            }
            finally
            {
                Console.WriteLine();
                Console.WriteLine("Done");
                Console.ReadLine();
            }
        }
    }
}
