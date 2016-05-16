using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using Microsoft.Win32;
using System.Net;
using System.Collections.Specialized;
using System.Deployment.Application;
using System.Web;

namespace ODT_Launcher
{
    public class InstallOffice
    { 
        
        public static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("args: " + args.Count());
                foreach (var arg in args) { Console.WriteLine("argument: "+ arg); }
                Console.ReadLine();
                var install = new InstallOffice();
                //install.RunProgram();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.ToString());
                Console.ReadLine();
            }
        }

        public void RunProgram()
        {
            var installDir = "";
            try
            {
                var url = "";
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    url = ApplicationDeployment.CurrentDeployment?.ActivationUri?.Query;
                }

                var queryString = Support.GetQueryStringParams(url);
                if (!queryString.Any())
                {
                    queryString = Support.GetArguments(Environment.GetCommandLineArgs());
                }

                var xmlServerPath = queryString.FirstOrDefault(q => q.Name.ToLower() == "xml")?.Value;
                var setupServerPath = queryString.FirstOrDefault(q => q.Name.ToLower() == "installer")?.Value;

                var currentDirectory = Environment.ExpandEnvironmentVariables("%temp%");
                installDir = currentDirectory + @"\OfficeProPlusSelfService";
                var loggingPath = currentDirectory + @"\OfficeProPlusSelfServiceLogs";

                Directory.CreateDirectory(installDir);

                Console.Write("Downloading Install Files...");

                Support.FileDownloader(xmlServerPath, installDir + @"\configuration.xml");
                Support.FileDownloader(setupServerPath, installDir + @"\officeSetup.exe");

                Console.WriteLine("Done");

                Support.MinimizeWindow();

                var odtFilePath = installDir + @"\officeSetup.exe";
                var xmlFilePath = installDir + @"\configuration.xml";

                Support.SetLoggingPath(xmlFilePath, loggingPath);

                if (!File.Exists(odtFilePath)) { throw (new Exception("Cannot find ODT Executable")); }
                if (!File.Exists(xmlFilePath)) { throw (new Exception("Cannot find Configuration Xml file")); }

                var p = new Process
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = odtFilePath,
                        Arguments = "/configure " + xmlFilePath,
                        CreateNoWindow = true,
                        UseShellExecute = false
                    },
                };
                p.Start();
                p.WaitForExit();

                Support.WaitForOfficeCtrUpadate();

                var errorMessage = Support.GetOdtErrorMessage(loggingPath);
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    Console.WriteLine(errorMessage.Trim());
                }
            }
            finally
            {
               Support.CleanUp(installDir);
            }
        }

    }
}

