using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.CSharp;
using Microsoft.OfficeProPlus.InstallGenerator;
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
                var codeProvider = new CSharpCodeProvider();
                var icc = codeProvider.CreateCompiler();

                var tmpPath = Environment.ExpandEnvironmentVariables("%temp%");
                var output = currentDirectory + @"\InstallOffice365ProPlus.exe";

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
                
                parameters.EmbeddedResources.Add(installProperties.OfficeVersion == OfficeVersion.Office2013
                    ? @".\Office2013Setup.exe"
                    : @".\Office2016Setup.exe");


                var fileContents = File.ReadAllText("InstallOffice.cs");
                fileContents = fileContents.Replace("public static void Main1(string[] args)",
                    "public static void Main(string[] args)");

                var configXml = new ConfigXmlParser(tmpPath + @"\configuration.xml");
                var addNode = configXml.ConfigurationXml.Add;
                if (addNode != null && addNode.Version != null)
                {
                    fileContents = fileContents.Replace("//[assembly: AssemblyVersion(\"\")]",
                         "[assembly: AssemblyVersion(\"" + addNode.Version + "\")]");
                }

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

        private void EmbedSourceFiles(CompilerParameters parameters, string sourcePath)
        {
            var currentDirectory = Directory.GetCurrentDirectory();
            var xmlFilePath = currentDirectory + @"\Files.xml";

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
