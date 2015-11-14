using System;
using System.IO;
using System.Linq;

namespace OfficeInstallGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
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

                var p = new OfficeInstallExecutableGenerator();
                p.Generate(OfficeVersion.Office2016, xmlConfiguration, @"E:\Users\rsmith.VCG\Desktop");
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
