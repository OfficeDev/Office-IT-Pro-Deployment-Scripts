using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.OfficeProPlus.InstallGenerator.Events;
using Microsoft.Win32;

//[assembly: AssemblyTitle("")]
//[assembly: AssemblyProduct("")]
//[assembly: AssemblyDescription("")]
//[assembly: AssemblyVersion("")]
//[assembly: AssemblyFileVersion("")]

public class InstallOffice
{

    private XmlDocument _xmlDoc = null;

    public static void Main1(string[] args)
    {
        using (var sw = new StreamWriter(@"C:\OfficeExeLog.txt"))
        {
            
        
        try
        {
            var install = new InstallOffice();
            install.RunProgram();
        }
        catch (Exception ex)
        {
            Console.WriteLine("ERROR: " + ex.ToString());
        }
        finally
        {
            Console.WriteLine();
        }
        }
    }

    public void RunProgram()
    {
        var fileNames = new List<string>();
        var installDir = "";

        try
        {
            MinimizeWindow();

            SilentInstall = false;

            var currentDirectory = Environment.ExpandEnvironmentVariables("%temp%");
            installDir = currentDirectory + @"\OfficeProPlus";

            Directory.CreateDirectory(installDir);
            //Directory.CreateDirectory(Environment.ExpandEnvironmentVariables(@"%temp%\OfficeProPlus\LogFiles"));

            var args = GetArguments();
            if (args.Any())
            {
                if (!HasValidArguments())
                {
                    ShowHelp();
                    return;
                }
            }

            var filesXml = GetTextFileContents("files.xml");
            if (!string.IsNullOrEmpty(filesXml))
            {
                _xmlDoc = new XmlDocument();
                _xmlDoc.LoadXml(filesXml);

            }
            Console.Write("Extracting Install Files...");
            fileNames = GetEmbeddedItems(installDir);
            Console.WriteLine("Done");

            var odtFilePath = installDir + @"\" + fileNames.FirstOrDefault(f => f.ToLower().EndsWith(".exe"));
            var xmlFilePath = installDir + @"\" + fileNames.FirstOrDefault(f => f.ToLower().EndsWith(".xml"));

            SetLoggingPath(xmlFilePath);

            SetSourcePath(xmlFilePath);

            if (!File.Exists(odtFilePath))
            {
                throw (new Exception("Cannot find ODT Executable"));
            }
            if (!File.Exists(xmlFilePath))
            {
                throw (new Exception("Cannot find Configuration Xml file"));
            }

            var runInstall = false;
            if (GetArguments().Any(a => a.Key.ToLower() == "/uninstall"))
            {
                xmlFilePath = UninstallOfficeProPlus(installDir, fileNames);
                runInstall = true;

                if (GetArguments().Any(a => a.Key.ToLower() == "/silent"))
                {
                    SilentInstall = true;
                }
            }
            else if (GetArguments().Any(a => a.Key.ToLower() == "/showxml"))
            {
                Console.Clear();
                var configXml = File.ReadAllText(xmlFilePath);
                Console.WriteLine(BeautifyXml(configXml));
            }
            else if (GetArguments().Any(a => a.Key.ToLower() == "/extractxml"))
            {
                var arg = GetArguments().FirstOrDefault(a => a.Key.ToLower() == "/extractxml");
                if (string.IsNullOrEmpty(arg.Value)) Console.WriteLine("ERROR: Invalid File Path");
                var configXml = BeautifyXml(File.ReadAllText(xmlFilePath));
                File.WriteAllText(arg.Value, configXml);
            }
            else
            {
                Console.WriteLine("Installing Office 365 ProPlus...");
                runInstall = true;
            }

            if (runInstall)
            {

                if (SilentInstall)
                {
                    var doc = new XmlDocument();
                    doc.Load(xmlFilePath);
                    SetConfigSilent(doc);
                    doc.Save(xmlFilePath);
                }

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

                WaitForOfficeCtrUpadate();

                var errorMessage = GetOdtErrorMessage();
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    Console.Error.WriteLine(errorMessage.Trim());
                }
            }
        }
        finally
        {
            CleanUp(installDir);
        }

    }

    private string UninstallOfficeProPlus(string installationDirectory, IEnumerable<string> fileNames)
    {
        Console.WriteLine("Uninstalling Office 365 ProPlus...");

        var doc = new XmlDocument();

        var root = doc.CreateElement("Configuration");

        var remove1 = doc.CreateElement("Remove");
        var all = doc.CreateAttribute("All");
        all.Value = "TRUE";

        remove1.Attributes.Append(all);
        root.AppendChild(remove1);

        doc.AppendChild(root);

        if (SilentInstall)
        {
            SetConfigSilent(doc);
        }

        doc.Save(installationDirectory + @"\configuration.xml");

        return installationDirectory + @"\" + fileNames.FirstOrDefault(f => f.ToLower().EndsWith(".xml"));
    }

    public async Task ChangeOfficeChannel(string targetChannel, string baseUrl)
    {
        var saveBaseUrl = "";
        try
        {
            saveBaseUrl = GetBaseCdnUrl();

            ChangeUpdateSource(baseUrl);
            ChangeBaseCdnUrl(baseUrl);

            RestartC2RSerivce();

            await RunOfficeUpdateAsync(targetChannel);
        }
        catch (Exception ex)
        {
            if (!string.IsNullOrEmpty(saveBaseUrl))
            {
                ChangeBaseCdnUrl(saveBaseUrl);
            }
            throw;
        }
        finally
        {
            ResetUpdateSource();
        }

    }

    public static void RestartC2RSerivce()
    {
        const string serviceName = "ClickToRunSvc";
        var service = new ServiceController(serviceName);
        service.Stop();
        service.WaitForStatus(ServiceControllerStatus.Stopped, new TimeSpan(0, 0, 5));
        service.Start();
        service.WaitForStatus(ServiceControllerStatus.Running, new TimeSpan(0, 0, 5));
    }

    public async Task RunOfficeUpdateAsync(string version)
    {
        await Task.Run(async () => { 
            var c2RPath = GetOfficeC2RPath() + @"\OfficeC2RClient.exe";
            if (File.Exists(c2RPath))
            {
                var arguments =
                    "/update user displaylevel=false forceappshutdown=true updatepromptuser=false updatetoversion=" + version;

                var p = new Process
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = c2RPath,
                        Arguments = arguments,
                        CreateNoWindow = true,
                        UseShellExecute = false
                    },
                };

                p.Start();
                p.WaitForExit();

                await Task.Delay(1000);

                WaitForOfficeCtrUpadateWithError();
            }
        });
    }

    #region Office Operations

    public string GetOfficeC2RPath()
    {
        var officeRegKey = GetOfficeCtrRegPath();
        var configKey = officeRegKey.OpenSubKey(@"Configuration");
        return configKey != null ? GetRegistryValue(configKey, "ClientFolder") : "";
    }

    public bool ProPlusLanguageInstalled(string productId, string language)
    {
        var officeRegKey = GetOfficeCtrRegPath();
        var prodKey = officeRegKey.OpenSubKey(@"ProductReleaseIDs");
        if (prodKey == null) return false;
        var subKeys = prodKey.GetSubKeyNames();
        if (!subKeys.Any()) return false;
        var mainId = subKeys.FirstOrDefault();
        if (mainId == null) return false;
        var mainKey = prodKey.OpenSubKey(mainId);
        if (mainKey == null) return false;
        var prodKeys = mainKey.GetSubKeyNames();
        if (!prodKeys.Any()) return false;

        foreach (var prodKeyName in prodKeys)
        {
            if (!prodKeyName.ToLower().Contains(productId.ToLower())) continue;
            var productKey = mainKey.OpenSubKey(prodKeyName);
            var languages = productKey.GetSubKeyNames();

            if (languages.Any(l => l.ToLower() == language.ToLower()))
            {
                return true;
            }
        }
        return false;
    }

    public string GetOdtErrorMessage()
    {
        var dirInfo = new DirectoryInfo(LoggingPath);
        try
        {

            foreach (var file in dirInfo.GetFiles("*.log"))
            {
                using (var reader = new StreamReader(file.FullName))
                {
                    do
                    {
                        var found = false;
                        var line = reader.ReadLine();
                        if (!line.ToLower().Contains("Prereq::ShowPrereqFailure:".ToLower())) continue;

                        var lineSplit = line.Split(':');
                        foreach (var part in lineSplit)
                        {
                            if (found)
                            {
                                return part;
                            }
                            else
                            {
                                if (part.ToLower().Contains("showprereqfailure"))
                                {
                                    found = true;
                                }
                            }
                        }
                    } while (reader.Peek() > -1);
                }

            }
        }
        catch { }
        finally
        {
            try
            {
                foreach (var file in dirInfo.GetFiles("*.log"))
                {
                    File.Copy(file.FullName, Environment.ExpandEnvironmentVariables(@"%temp%\" + file.Name), true);
                }

                if (Directory.Exists(LoggingPath))
                {
                    Directory.Delete(LoggingPath);
                }
            }
            catch { }
        }
        return null;
    }

    private void SetLoggingPath(string xmlFilePath)
    {
        var tempPath = Environment.ExpandEnvironmentVariables("%temp%");
        const string logFolderName = "OfficeProPlusLogs";
        LoggingPath = tempPath + @"\" + logFolderName;
        if (Directory.Exists(LoggingPath))
        {
            try
            {
                Directory.Delete(LoggingPath);
            }
            catch { }
        }
        Directory.CreateDirectory(LoggingPath);

        var xmlDoc = new XmlDocument();
        xmlDoc.Load(xmlFilePath);

        var loggingNode = xmlDoc.SelectSingleNode("/Configuration/Logging");
        if (loggingNode == null)
        {
            loggingNode = xmlDoc.CreateElement("Logging");
            xmlDoc.DocumentElement.AppendChild(loggingNode);
        }

        SetAttribute(xmlDoc, loggingNode, "Path", LoggingPath);
        xmlDoc.Save(xmlFilePath);
    }

    private void SetSourcePath(string xmlFilePath)
    {
        var tempPath = Environment.ExpandEnvironmentVariables("%temp%");
        const string officeFolderName = "OfficeProPlus";

        var officeFolderPath = tempPath + @"\" + officeFolderName;
        if (Directory.Exists(officeFolderPath + @"\Office"))
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFilePath);

            var addNode = xmlDoc.SelectSingleNode("/Configuration/Add");
            if (addNode != null)
            {
                SetAttribute(xmlDoc, addNode, "SourcePath", officeFolderPath);
                xmlDoc.Save(xmlFilePath);
            }
        }
    }

    private void SetConfigSilent(XmlDocument doc)
    {
        var display = doc.SelectSingleNode("/Configuration/Display");
        if (display == null)
        {
            display = doc.CreateElement("Display");
            doc.AppendChild(display);
        }

        SetAttribute(doc, display, "Level", "None");
        SetAttribute(doc, display, "AcceptEULA", "TRUE");
    }

    #endregion
    
    #region Update Monitoring

    public bool IsUpdateRunning()
    {
        var scenarioTasks = GetRunningScenarioTasks(true);
        if (scenarioTasks == null) return false;
        if (scenarioTasks.Count == 0) return false;

        var anyRunning = scenarioTasks.Any(s => s.State == "TASKSTATE_EXECUTING");
        return anyRunning;
    }
    
    public void WaitForOfficeCtrUpadateWithError(bool showStatus = false)
    {
        if (showStatus) { Console.WriteLine("Waiting for Install to Complete..."); }

        var operationStart = DateTime.Now;
        var totalOperationStart = DateTime.Now;

        //Thread.Sleep(new TimeSpan(0,0,10));

        var startTime = DateTime.Now;
        var installRunning = false;
        var anyCancelled = false;
        var anyFailed = false;
        var tasksDone = new List<string>();
        var tasksRunning = new List<string>();
        var opRunning = false;
        var currentOperationName = "";
        do
        {
            var allComplete = true;

            var scenarioTasks = GetRunningScenarioTasks(true);
            if (scenarioTasks == null) return;
            if (scenarioTasks.Count == 0) return;

            anyCancelled = scenarioTasks.Any(s => s.State == "TASKSTATE_CANCELLED");
            anyFailed = scenarioTasks.Any(s => s.State == "TASKSTATE_FAILED");

            var completeScenarios = scenarioTasks.Where(s => s.State.EndsWith("ED") && !tasksDone.Contains(s.State)).ToList();
            var completeScenariosNames = completeScenarios.Select(s => s.Scenario).ToList();

            if (completeScenarios.Count > 0 && opRunning && completeScenarios.Any(s => tasksRunning.Contains(s.Scenario)))
            {
                if (showStatus)
                {
                    var statusName = completeScenarios[0].State.Split('_')[1];
                    Console.WriteLine(statusName);
                }
                opRunning = false;
                tasksRunning.Clear();
            }

            var anyRunning = scenarioTasks.Any(s => s.State == "TASKSTATE_EXECUTING");
            if (anyRunning) allComplete = false;

            var executingTasks = scenarioTasks.Where(s => s.State == "TASKSTATE_EXECUTING" &&
                                                          !tasksRunning.Contains(s.Scenario)).ToList();
            if (executingTasks.Count > 0)
            {
                installRunning = true;
                var currentOperation = GetCurrentOperation(executingTasks);
                Console.Write(currentOperation + ": ");
                currentOperationName = currentOperation.ToString();

                if (UpdatingOfficeStatus != null)
                {
                    UpdatingOfficeStatus(this, new Events.UpdatingOfficeArgs()
                    {
                        Status = currentOperationName + "..."
                    });
                }

                opRunning = true;
                tasksRunning.AddRange(executingTasks.Select(t => t.Scenario));
            }

            tasksDone = completeScenariosNames;

            if (allComplete) break;
            Thread.Sleep(1000);
        } while (true);

        if (installRunning)
        {
            if (anyFailed)
            {
                throw (new Exception("Update failed"));
            }
            if (anyCancelled)
            {
                throw (new Exception("Update cancelled"));
            }
            if (!anyCancelled && !anyFailed) Console.WriteLine("Install Complete");
        }
        else
        {
            throw (new Exception("Update not running"));
        }
    }

    public void WaitForOfficeCtrUpadate(bool showStatus = false)
    {
        if (showStatus) { Console.WriteLine("Waiting for Install to Complete..."); }

        var operationStart = DateTime.Now;
        var totalOperationStart = DateTime.Now;

        //Thread.Sleep(new TimeSpan(0,0,10));

        var startTime = DateTime.Now;
        var installRunning = false;
        var anyCancelled = false;
        var anyFailed = false;
        var tasksDone = new List<string>();
        var tasksRunning = new List<string>();
        var opRunning = false;
        do
        {
            var allComplete = true;

            var scenarioTasks = GetRunningScenarioTasks();
            if (scenarioTasks == null) return;
            if (scenarioTasks.Count == 0) return;

            anyCancelled = scenarioTasks.Any(s => s.State == "TASKSTATE_CANCELLED");
            anyFailed = scenarioTasks.Any(s => s.State == "TASKSTATE_FAILED");

            var completeScenarios = scenarioTasks.Where(s => s.State.EndsWith("ED") && !tasksDone.Contains(s.State)).ToList();
            var completeScenariosNames = completeScenarios.Select(s => s.Scenario).ToList();

            if (completeScenarios.Count > 0 && opRunning && completeScenarios.Any(s => tasksRunning.Contains(s.Scenario)))
            {
                if (showStatus)
                {
                    Console.WriteLine(completeScenarios[0].State.Split('_')[1]);
                }
                opRunning = false;
                tasksRunning.Clear();
            }

            var anyRunning = scenarioTasks.Any(s => s.State == "TASKSTATE_EXECUTING");
            if (anyRunning) allComplete = false;

            var executingTasks = scenarioTasks.Where(s => s.State == "TASKSTATE_EXECUTING" &&
                                                          !tasksRunning.Contains(s.Scenario)).ToList();
            if (executingTasks.Count > 0)
            {
                installRunning = true;
                var currentOperation = GetCurrentOperation(executingTasks);
                Console.Write(currentOperation + ": ");
                opRunning = true;
                tasksRunning.AddRange(executingTasks.Select(t => t.Scenario));
            }

            tasksDone = completeScenariosNames;

            if (allComplete) break;
            Thread.Sleep(1000);
        } while (true);

        if (installRunning)
        {
            if (anyFailed) Console.WriteLine("Install Failed");
            if (anyCancelled) Console.WriteLine("Install Cancelled");
            if (!anyCancelled && !anyFailed) Console.WriteLine("Install Complete");
        }
        else
        {
            Console.WriteLine("Install Not Running");
        }


    }

    public List<ExecutingScenario> GetRunningScenarioTasks(bool assumeUpdate = false)
    {
        var execScenarios = new List<ExecutingScenario>();
        var executingScenario = GetExecutingScenario();
        if (string.IsNullOrEmpty(executingScenario))
        {
            if (assumeUpdate)
            {
                executingScenario = "Update";
            }
            else
            {
                return new List<ExecutingScenario>();
            }
        }

        var mainRegKey = GetOfficeCtrRegPath();
        if (mainRegKey == null) return new List<ExecutingScenario>();

        var scenarioKey = mainRegKey.OpenSubKey("scenario");
        if (scenarioKey == null) return null;

        var subKeyNames = scenarioKey.GetSubKeyNames();

        foreach (var subKeyName in subKeyNames)
        {
            if (subKeyName.ToUpper() == executingScenario.ToUpper())
            {
                var execScenKey = scenarioKey.OpenSubKey(subKeyName + @"\TasksState");
                if (execScenKey == null) continue;
                var valueNames = execScenKey.GetValueNames();

                foreach (var valueName in valueNames)
                {
                    var strState = "";
                    var state = execScenKey.GetValue(valueName);
                    if (state != null) strState = state.ToString();

                    execScenarios.Add(new ExecutingScenario()
                    {
                        Scenario = valueName,
                        State = strState.ToUpper()
                    });
                }
            }
        }
        return execScenarios;
    }

    public void ResetUpdateSource()
    {
        const string policyPath = @"SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate";
        var policyKey = Registry.LocalMachine.OpenSubKey(policyPath, true);
        if (policyKey != null)
        {
            var saveUpdatePath = GetRegistryValue(policyKey, "saveupdatepath");
            if (!string.IsNullOrEmpty(saveUpdatePath))
            {
                policyKey.SetValue("updatepath", saveUpdatePath, RegistryValueKind.String);
                policyKey.DeleteValue("saveupdatepath");
            }
        }

        var mainRegKey = GetOfficeCtrRegPath();
        if (mainRegKey == null) return;

        var configKey = mainRegKey.OpenSubKey(@"Configuration");
        if (configKey == null) return;

        var saveUpdateUrl = GetRegistryValue(policyKey, "UpdateUrl");
        if (string.IsNullOrEmpty(saveUpdateUrl)) return;

        configKey.SetValue("UpdateUrl", saveUpdateUrl, RegistryValueKind.String);
        configKey.DeleteValue("SaveUpdateUrl");
    }

    public string ChangeUpdateSource(string updateSource)
    {
        var currentupdatepath = "";

        const string policyPath = @"SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate";
        var policyKey = Registry.LocalMachine.OpenSubKey(policyPath, true);
        if (policyKey != null)
        {
            currentupdatepath = GetRegistryValue(policyKey, "updatepath");
            var saveupdatePath = GetRegistryValue(policyKey, "saveupdatepath");
            if (!string.IsNullOrEmpty(currentupdatepath) && !string.IsNullOrEmpty(updateSource))
            {
                if (string.IsNullOrEmpty(saveupdatePath))
                {
                    policyKey.SetValue("saveupdatepath", currentupdatepath, RegistryValueKind.String);
                }
                policyKey.SetValue("updatepath", updateSource, RegistryValueKind.String);
            }
        }

        if (!string.IsNullOrEmpty(currentupdatepath)) return currentupdatepath;

        var mainRegKey = GetOfficeCtrRegPath();
        if (mainRegKey == null) return null;

        var configKey = mainRegKey.OpenSubKey(@"Configuration");
        if (configKey == null) return null;

        currentupdatepath = GetRegistryValue(configKey, "UpdateUrl");
        var saveupdateUrl = GetRegistryValue(configKey, "SaveUpdateUrl");
        if (string.IsNullOrEmpty(currentupdatepath) || string.IsNullOrEmpty(updateSource)) return currentupdatepath;

        if (string.IsNullOrEmpty(saveupdateUrl))
        {
            configKey.SetValue("SaveUpdateUrl", currentupdatepath, RegistryValueKind.String);
        }

        configKey.SetValue("UpdateUrl", updateSource, RegistryValueKind.String);

        return currentupdatepath;
    }

    public string GetBaseCdnUrl()
    {
        var mainRegKey = GetOfficeCtrRegPath();
        if (mainRegKey == null) return "";

        var configKey = mainRegKey.OpenSubKey(@"Configuration", true);
        if (configKey == null) return "";

        return GetRegistryValue(configKey, "CDNBaseUrl");
    }

    public void ChangeBaseCdnUrl(string updateSource)
    {
        var mainRegKey = GetOfficeCtrRegPath();
        if (mainRegKey == null) return;

        var configKey = mainRegKey.OpenSubKey(@"Configuration", true);
        if (configKey == null) return;

        var cdnBaseUrl = GetRegistryValue(configKey, "CDNBaseUrl");
        configKey.SetValue("CDNBaseUrl", updateSource, RegistryValueKind.String);
    }


    public string GetExecutingScenario()
    {
        var mainRegKey = GetOfficeCtrRegPath();
        if (mainRegKey == null) return null;
        var execScenario = mainRegKey.GetValue("ExecutingScenario");
        return execScenario != null ? execScenario.ToString() : null;
    }

    public RegistryKey GetOfficeCtrRegPath()
    {
        var path16 = @"SOFTWARE\Microsoft\Office\ClickToRun";
        var path15 = @"SOFTWARE\Microsoft\Office\15.0\ClickToRun";

        var office16Key = Registry.LocalMachine.OpenSubKey(path16, true);
        var office15Key = Registry.LocalMachine.OpenSubKey(path15, true);

        if (office16Key != null)
        {
            return office16Key;
        }
        else
        {
            if (office15Key != null)
            {
                return office15Key;
            }
        }

        var Hklm32 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);

        office16Key = Hklm32.OpenSubKey(path16, true);
        office15Key = Hklm32.OpenSubKey(path15, true);

        if (office16Key != null)
        {
            return office16Key;
        }
        else
        {
            if (office15Key != null)
            {
                return office15Key;
            }
        }

        return null;
    }

    public CurrentOperation GetCurrentOperation(List<ExecutingScenario> executingTasks)
    {
        if (executingTasks.Any(t => t.Scenario.ToUpper().Contains("DOWNLOAD")))
        {
            return CurrentOperation.Downloading;
        }
        if (executingTasks.Any(t => t.Scenario.ToUpper().Contains("APPLY")))
        {
            return CurrentOperation.Applying;
        }
        if (executingTasks.Any(t => t.Scenario.ToUpper().Contains("FINALIZE")))
        {
            return CurrentOperation.Finalizing;
        }
        if (executingTasks.Any(t => t.Scenario.ToUpper().Contains("INTEGRATE")))
        {
            return CurrentOperation.Finalizing;
        }
        return CurrentOperation.Starting;
    }

    #endregion
    
    #region File Operations

    private void CleanUp(string installDir)
    {
        var dirInfo = new DirectoryInfo(installDir);
        foreach (var file in dirInfo.GetFiles())
        {
            try
            {
                file.Delete();
            }
            catch { }
        }

        foreach (var directory in dirInfo.GetDirectories())
        {
            try
            {
                directory.Delete(true);
            }
            catch { }
        }

        try
        {
            Directory.Delete(installDir, true);
        }
        catch { }
    }

    public string GetTextFileContents(string fileName)
    {
        var resourceName = "";
        var resourceNames = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceNames();
        foreach (var name in resourceNames)
        {
            if (name.ToLower().EndsWith(fileName.ToLower()))
            {
                resourceName = name;
            }
        }

        if (!string.IsNullOrEmpty(resourceName))
        {
            var strReturn = "";
            using (var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            using (var reader = new StreamReader(stream))
            {
                strReturn = reader.ReadToEnd();
            }
            return strReturn;
        }
        return null;
    }

    public List<string> GetEmbeddedItems(string targetDirectory)
    {
        var returnFiles = new List<string>();
        var assemblyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
        if (assemblyPath == null) return returnFiles;

        //var appRoot = new Uri(assemblyPath).LocalPath;
        var assembly = System.Reflection.Assembly.GetExecutingAssembly();
        var assemblyName = assembly.GetName().Name;

        foreach (var resourceStreamName in assembly.GetManifestResourceNames())
        {
            using (var input = assembly.GetManifestResourceStream(resourceStreamName))
            {
                var fileName = Regex.Replace(resourceStreamName, "^" + assemblyName + ".", "", RegexOptions.IgnoreCase);
                fileName = Regex.Replace(fileName, "^Resources.", "", RegexOptions.IgnoreCase);

                returnFiles.Add(fileName);

                var filePath = Path.Combine(targetDirectory, fileName);

                if (File.Exists(filePath)) File.Delete(filePath);

                using (Stream output = File.Create(filePath))
                {
                    CopyStream(input, output);
                }

                var md5Hash = GenerateMD5Hash(filePath);
                MoveFile(targetDirectory, md5Hash, fileName);
            }
        }
        return returnFiles;
    }

    public void MoveFile(string rootDirectory, string md5Hash, string fileName)
    {
        if (_xmlDoc == null) return;
        var fileNode = _xmlDoc.SelectSingleNode("//File[@Hash='" + md5Hash + "' and @FileName='" + fileName + "']");
        if (fileNode == null) return;

        var folderPath = fileNode.Attributes["FolderPath"].Value;
        var xmlfileName = fileNode.Attributes["FileName"].Value;

        Directory.CreateDirectory(rootDirectory + @"\" + folderPath);
        File.Move(rootDirectory + @"\" + fileName, rootDirectory + @"\" + folderPath + @"\" + xmlfileName);
    }

    private static void CopyStream(Stream input, Stream output)
    {
        // Insert null checking here for production
        var buffer = new byte[8192];

        int bytesRead;
        while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
        {
            output.Write(buffer, 0, bytesRead);
        }
    }
    
    #endregion 

    #region Support Functions

    private static string GenerateMD5Hash(string filePath)
    {
        using (var md5 = MD5.Create())
        {
            using (var stream = File.OpenRead(filePath))
            {
                return BitConverter.ToString(md5.ComputeHash(stream)).Replace("-", "").ToLower();
            }
        }
    }

    private void ShowHelp()
    {
        Console.WriteLine("Usage: " + Process.GetCurrentProcess().ProcessName + " [/uninstall] [/showxml] [/extractxml={File Path}]");
        Console.WriteLine();
        Console.WriteLine("  /uninstall\t\t\tRemoves all installed Office 365 ProPlus");
        Console.WriteLine("  \t\t\t\tproducts.");
        Console.WriteLine("  /silent\t\t\tInstalls with prompts");
        Console.WriteLine("  /showxml\t\t\tDisplays the current Office 365 ProPlus");
        Console.WriteLine("  \t\t\t\tconfiguration xml.");
        Console.WriteLine("  /extractxml={File Path}\tExtracts the current Office 365 ProPlus");
        Console.WriteLine("  \t\t\t\tconfiguration xml to the specified file path.");
    }

    private void MinimizeWindow()
    {
        IntPtr winHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
        ShowWindow(winHandle, SW_SHOWMINIMIZED);
    }

    public List<CmdArgument> GetArguments()
    {
        var returnList = new List<CmdArgument>();
        var n = -1;
        foreach (var arg in Environment.GetCommandLineArgs())
        {
            n++;
            if (n == 0) continue;
            var key = arg;
            var value = "";
            if (arg.Contains("="))
            {
                key = arg.Split('=')[0];
                value = arg.Split('=')[1];
            }

            returnList.Add(new CmdArgument()
            {
                Key = key,
                Value = value
            });
        }
        return returnList;
    }

    private string Beautify(XmlDocument doc)
    {
        var sb = new StringBuilder();
        var settings = new XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            NewLineChars = "\r\n",
            NewLineHandling = NewLineHandling.Replace,
            OmitXmlDeclaration = true
        };
        using (var writer = XmlWriter.Create(sb, settings))
        {
            doc.Save(writer);
        }

        var xml = sb.ToString();
        return xml;
    }

    private string BeautifyXml(string xml)
    {
        var doc = new XmlDocument();
        doc.LoadXml(xml);
        return Beautify(doc);
    }

    private void SetAttribute(XmlDocument xmlDoc, XmlNode xmlNode, string name, string value)
    {
        var pathAttr = xmlNode.Attributes[name];
        if (pathAttr == null)
        {
            pathAttr = xmlDoc.CreateAttribute(name);
            xmlNode.Attributes.Append(pathAttr);
        }
        pathAttr.Value = value;
    }

    private string GetRegistryValue(RegistryKey regKey, string property)
    {
        if (regKey != null)
        {
            if (regKey.GetValue(property) == null) return "";
            return regKey.GetValue(property).ToString();
        }
        return "";
    }
    
    //private void extractWixTools(string installDir)
    //{
    //    string zipPath = installDir + @"\tools.zip";
    //    string extractPath = "\tools";

    //    using (ZipArchive archive = ZipFile.OpenRead(zipPath))
    //    {
    //        foreach (ZipArchiveEntry entry in archive.Entries)
    //        {
    //            entry.ExtractToFile(Path.Combine(extractPath+entry.FullName));
    //        }
    //    }
    //}

    #endregion

    #region Properties

    private bool HasValidArguments()
    {
        return !GetArguments().Any(a => (a.Key.ToLower() != "/uninstall" &&
                                         a.Key.ToLower() != "/showxml" &&
                                         a.Key.ToLower() != "/silent" &&
                                         a.Key.ToLower() != "/extractxml"));
    }

    public string LoggingPath { get; set; }

    public bool SilentInstall { get; set; }

    public static string ResourcePath
    {
        get
        {
            return Directory.GetCurrentDirectory();
        }
    }


    #endregion

    #region WindowsFunctions

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    public const int SW_SHOWMINIMIZED = 2;

    #endregion

    #region Events

    public event Events.UpdatingOfficeEventHandler UpdatingOfficeStatus = null;

    #endregion

}

public class ODTLogFile
{
    public string FilePath { get; set; }

    public DateTime ModifiedTime { get; set; }
}

public enum CurrentOperation
{
    Starting = 0,
    Downloading = 1,
    Applying = 2,
    Finalizing = 3
}

public class ExecutingScenario
{

    public string Scenario { get; set; }

    public string State { get; set; }

}

public class CmdArgument
{
    public string Key { get; set; }

    public string Value { get; set; }
}
