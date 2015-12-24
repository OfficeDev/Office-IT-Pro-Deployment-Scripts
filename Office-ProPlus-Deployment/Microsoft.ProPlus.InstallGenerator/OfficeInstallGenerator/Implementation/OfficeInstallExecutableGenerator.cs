using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Micorosft.OfficeProPlus.ConfigurationXml.Enums;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.CSharp;
using Microsoft.OfficeProPlus.InstallGenerator;
using Microsoft.OfficeProPlus.InstallGenerator.Extensions;
using Microsoft.OfficeProPlus.InstallGenerator.Implementation;
using Microsoft.Win32;



namespace OfficeInstallGenerator
{
    public class OfficeInstallExecutableGenerator : IOfficeInstallGenerator
    {

        public IOfficeInstallReturn Generate(IOfficeInstallProperties installProperties)
        {
            var currentDirectory = Directory.GetCurrentDirectory();
            var embededExeFiles = new List<string>();
            try
            {
                if (Directory.Exists(currentDirectory + @"\Project"))
                {
                    currentDirectory = currentDirectory + @"\Project";
                }

                var codeProvider = new CSharpCodeProvider();
                var icc = codeProvider.CreateCompiler();

                var tmpPath = Environment.ExpandEnvironmentVariables("%temp%");
                var output = currentDirectory + @"\InstallOffice365ProPlus.exe";

                if (!string.IsNullOrEmpty(installProperties.ExecutablePath))
                {
                    output = installProperties.ExecutablePath;
                }

                var parameters = new CompilerParameters
                {
                    GenerateExecutable = true,
                    OutputAssembly = output
                };
                parameters.ReferencedAssemblies.Add("System.dll");
                parameters.ReferencedAssemblies.Add("System.Xml.dll");
                parameters.ReferencedAssemblies.Add("System.Core.dll");
                parameters.ReferencedAssemblies.Add("Microsoft.CSharp.dll");

                embededExeFiles = EmbeddedResources.GetEmbeddedItems(currentDirectory, @"\.exe$");

                File.Copy(installProperties.ConfigurationXmlPath, tmpPath + @"\configuration.xml", true);

                parameters.EmbeddedResources.Add(tmpPath + @"\configuration.xml");

               // parameters.EmbeddedResources.Add(@"\tools\");

                var office2013Setup = DirectoryHelper.GetCurrentDirectoryFilePath("Office2013Setup.exe");
                var office2016Setup = DirectoryHelper.GetCurrentDirectoryFilePath("Office2016Setup.exe"); 

                parameters.EmbeddedResources.Add(installProperties.OfficeVersion == OfficeVersion.Office2013
                    ? office2013Setup
                    : office2016Setup);

                var installOfficeFp = DirectoryHelper.GetCurrentDirectoryFilePath("InstallOffice.cs"); 

                var fileContents = File.ReadAllText(installOfficeFp);
                fileContents = fileContents.Replace("public static void Main1(string[] args)",
                    "public static void Main(string[] args)");

                var configXml = new ConfigXmlParser(tmpPath + @"\configuration.xml");
                var addNode = configXml.ConfigurationXml.Add;
                if (addNode != null && addNode.Version != null)
                {
                    fileContents = fileContents.Replace("//[assembly: AssemblyVersion(\"\")]",
                         "[assembly: AssemblyVersion(\"" + addNode.Version + "\")]");
                }

                if (configXml.ConfigurationXml.Logging == null)
                {
                    configXml.ConfigurationXml.Logging = new ODTLogging();
                }

                configXml.ConfigurationXml.Logging.Level = LoggingLevel.Standard;
                configXml.ConfigurationXml.Logging.Path = "%temp%";

                if (installProperties.SourceFilePath != null)
                {
                    if (!Directory.Exists(installProperties.SourceFilePath + @"\Office"))
                    {
                        throw (new DirectoryNotFoundException("Invalid Source Path: " + installProperties.SourceFilePath));
                    }

                    EmbedSourceFiles(parameters, installProperties.SourceFilePath + @"\Office");
                }

                if (installProperties.OfficeVersion == OfficeVersion.Office2013)
                {
                    fileContents = fileContents.Replace("//[assembly: AssemblyTitle(\"\")]",
                        "[assembly: AssemblyTitle(\"" + "Office 365 ProPlus (2013)" + "\")]");
                    fileContents = fileContents.Replace("//[assembly: AssemblyDescription(\"\")]",
                        "[assembly: AssemblyDescription(\"" + "Office 365 ProPlus (2013)" + "\")]");
                }

                if (installProperties.OfficeVersion == OfficeVersion.Office2016)
                {
                    fileContents = fileContents.Replace("//[assembly: AssemblyTitle(\"\")]",
                        "[assembly: AssemblyTitle(\"" + "Office 365 ProPlus (2016)" + "\")]");
                    fileContents = fileContents.Replace("//[assembly: AssemblyDescription(\"\")]",
                        "[assembly: AssemblyDescription(\"" + "Office 365 ProPlus (2016)" + "\")]");
                }

                var results = icc.CompileAssemblyFromSource(parameters, fileContents);

                if (results.Errors.Count > 0)
                {
                    var strBuilder = new StringBuilder();
                    foreach (CompilerError CompErr in results.Errors)
                    {
                        var errorText = "Line number " + CompErr.Line +
                                        ", Error Number: " + CompErr.ErrorNumber +
                                        ", '" + CompErr.ErrorText + ";" +
                                        Environment.NewLine + Environment.NewLine;
                        strBuilder.AppendLine(errorText);
                    }
                    throw (new Exception(strBuilder.ToString()));
                }

                return new OfficeInstallReturn()
                {
                    GeneratedFilePath = output
                };
            }
            finally
            {
                foreach (var fileName in embededExeFiles)
                {
                    if (File.Exists(currentDirectory + @"\" + fileName))
                    {
                        File.Delete(currentDirectory + @"\" + fileName);
                    }
                }
            }
        }

        public void InstallOffice(string configurationXml)
        {
            var tmpPath = Environment.ExpandEnvironmentVariables("%temp%");
            var embededExeFiles = EmbeddedResources.GetEmbeddedItems(tmpPath, @"\.exe$");

            var installExe = tmpPath + @"\" + embededExeFiles.FirstOrDefault(f => f.ToLower().Contains("2016"));
            var xmlPath = tmpPath + @"\configuration.xml";

            if (File.Exists(xmlPath)) File.Delete(xmlPath);
            File.WriteAllText(xmlPath, configurationXml);

            var p = new Process
            {
                StartInfo = new ProcessStartInfo()
                {
                    FileName = installExe,
                    Arguments = "/configure " + tmpPath + @"\configuration.xml",
                    CreateNoWindow = true,
                    UseShellExecute = false
                },
            };
            p.Start();
            p.WaitForExit();

            if (File.Exists(xmlPath)) File.Delete(xmlPath);

            foreach (var exeFilePath in embededExeFiles)
            {
                try
                {
                    if (File.Exists(tmpPath + @"\" + exeFilePath))
                    {
                        File.Delete(tmpPath + @"\" + exeFilePath);
                    }
                }
                catch { }
            }
        }

        private void EmbedSourceFiles(CompilerParameters parameters, string sourcePath)
        {
            var xmlFilePath = DirectoryHelper.GetCurrentDirectoryFilePath("Files.xml"); 

            var dirInfo = new DirectoryInfo(sourcePath);
            var sourceFiles = dirInfo.GetFiles("*.*", SearchOption.AllDirectories);

            var fileCacher = new FilePathCacher(xmlFilePath);

            foreach (var sourceFile in sourceFiles)
            {
                fileCacher.AddFile(dirInfo.Parent.FullName, sourceFile.FullName);
                parameters.EmbeddedResources.Add(sourceFile.FullName);
            }

            parameters.EmbeddedResources.Add(xmlFilePath);
        }

    }



}
