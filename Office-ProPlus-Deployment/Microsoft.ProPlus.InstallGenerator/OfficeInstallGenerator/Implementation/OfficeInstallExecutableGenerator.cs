using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Enums;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.CSharp;
using Microsoft.OfficeProPlus.InstallGenerator;
using Microsoft.OfficeProPlus.InstallGenerator.Extensions;
using Microsoft.OfficeProPlus.InstallGenerator.Implementation;


namespace OfficeInstallGenerator
{
    public class OfficeInstallExecutableGenerator : IOfficeInstallGenerator
    {
        private List<FileInfo> filesMarkedForDelete = new List<FileInfo>();

        public IOfficeInstallReturn Generate(IOfficeInstallProperties installProperties, string remoteLogPath = "")
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
                parameters.ReferencedAssemblies.Add("System.Windows.Forms.dll");
                parameters.ReferencedAssemblies.Add("Microsoft.CSharp.dll");
                parameters.ReferencedAssemblies.Add("System.Management.dll");

                embededExeFiles = EmbeddedResources.GetEmbeddedItems(currentDirectory, @"\.exe$");

                File.Copy(installProperties.ConfigurationXmlPath, tmpPath + @"\configuration.xml", true);

                var productIdPath = tmpPath + @"\productid.txt";
                var remoteLog = tmpPath + @"\RemoteLog.txt";
                File.WriteAllText(productIdPath, installProperties.ProductId);
                File.WriteAllText(remoteLog, remoteLogPath);

                parameters.EmbeddedResources.Add(tmpPath + @"\configuration.xml");
                parameters.EmbeddedResources.Add(productIdPath);
                parameters.EmbeddedResources.Add(remoteLog);

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

                    //commenting out, trying to go a different path with this, copy out the source files to the same dir as exe file
                    //possibly only copy out if contents of source file greater than 1.5 GB
                    long embeddedFileSize = CalcSize(parameters, installProperties.SourceFilePath + @"\Office", installProperties.BuildVersion, installProperties.OfficeClientEdition);
                    
                    //find file size of embedded files, if less than 1.5 GB, then embed, if greater, copy out as separate folder in same dir as the MSI or exe file
                    if (embeddedFileSize < 1900000000)
                    {
                        EmbedSourceFiles(parameters, installProperties.SourceFilePath + @"\Office",
                            installProperties.BuildVersion, installProperties.OfficeClientEdition);
                    }
                    else
                    {
                        CopyFolder(new DirectoryInfo(installProperties.SourceFilePath), new DirectoryInfo(installProperties.ExecutablePath.Substring(0, installProperties.ExecutablePath.LastIndexOf(@"\"))));
                    }
                    
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
                //delete temp files
                foreach (var fileMarkedForDelete in filesMarkedForDelete)
                {
                    fileMarkedForDelete.Delete();
                }

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

        private void EmbedSourceFiles(CompilerParameters parameters, string sourcePath, string version = null, OfficeClientEdition officeClientEdition = OfficeClientEdition.Office32Bit)
        {
            var embedFileList = new List<string>();
            var xmlFilePath = DirectoryHelper.GetCurrentDirectoryFilePath("Files.xml"); 

            var dirInfo = new DirectoryInfo(sourcePath);
            var sourceFiles = dirInfo.GetFiles("*.*", SearchOption.AllDirectories);

            var fileCacher = new FilePathCacher(xmlFilePath);

            foreach (var sourceFile in sourceFiles)
            {
                if (!string.IsNullOrEmpty(version))
                {
                    if (!(sourceFile.FullName.ToLower().Contains(version.ToLower()) ||
                        sourceFile.Name.ToLower() == "v32.cab" ||
                        sourceFile.Name.ToLower() == "v64.cab"))
                    {
                        continue;
                    }
                }

                if (officeClientEdition == OfficeClientEdition.Office32Bit)
                {
                    if (sourceFile.Name.ToLower().Contains(".x64."))
                    {
                        continue;
                    }
                }
                else
                {
                    if (sourceFile.Name.ToLower().Contains(".x86."))
                    {
                        continue;
                    }
                }
                if (embedFileList.Contains(sourceFile.Name, StringComparer.CurrentCultureIgnoreCase))
                 {
                    sourceFile.CopyTo(sourceFile.DirectoryName + "\\copyof" + sourceFile.Name, true);
                    fileCacher.AddFile(dirInfo.Parent.FullName, sourceFile.DirectoryName + "\\copyof" + sourceFile.Name);
                    parameters.EmbeddedResources.Add(sourceFile.DirectoryName + "\\copyof" + sourceFile.Name);
                    filesMarkedForDelete.Add(new FileInfo(sourceFile.DirectoryName + "\\copyof" + sourceFile.Name));
                 }
                 else
                 {
                    fileCacher.AddFile(dirInfo.Parent.FullName, sourceFile.FullName);
                    parameters.EmbeddedResources.Add(sourceFile.FullName);
                 }
                embedFileList.Add(sourceFile.Name);
            }

            parameters.EmbeddedResources.Add(xmlFilePath);
        }

        private long CalcSize(CompilerParameters parameters, string sourcePath, string version = null, OfficeClientEdition officeClientEdition = OfficeClientEdition.Office32Bit)
        {
            long runningTotalFileSize = 0;
            var xmlFilePath = DirectoryHelper.GetCurrentDirectoryFilePath("Files.xml");

            var dirInfo = new DirectoryInfo(sourcePath);
            var sourceFiles = dirInfo.GetFiles("*.*", SearchOption.AllDirectories);

            var fileCacher = new FilePathCacher(xmlFilePath);

            foreach (var sourceFile in sourceFiles)
            {
                if (!string.IsNullOrEmpty(version))
                {
                    if (!(sourceFile.FullName.ToLower().Contains(version.ToLower()) ||
                        sourceFile.Name.ToLower() == "v32.cab" ||
                        sourceFile.Name.ToLower() == "v64.cab"))
                    {
                        continue;
                    }
                }

                if (officeClientEdition == OfficeClientEdition.Office32Bit)
                {
                    if (sourceFile.Name.ToLower().Contains(".x64."))
                    {
                        continue;
                    }
                }
                else
                {
                    if (sourceFile.Name.ToLower().Contains(".x86."))
                    {
                        continue;
                    }
                }

                //fileCacher.AddFile(dirInfo.Parent.FullName, sourceFile.FullName);

                //parameters.EmbeddedResources.Add(sourceFile.FullName);
                FileInfo file = new FileInfo(sourceFile.FullName);
                runningTotalFileSize += file.Length;
            }

            //parameters.EmbeddedResources.Add(xmlFilePath);
            return runningTotalFileSize;
        }


        public static void CopyFolder(DirectoryInfo source, DirectoryInfo target)
        {
            foreach (DirectoryInfo dir in source.GetDirectories())
                CopyFolder(dir, target.CreateSubdirectory(dir.Name));
            foreach (FileInfo file in source.GetFiles())
                file.CopyTo(Path.Combine(target.FullName, file.Name), true);
        }

    }



}
