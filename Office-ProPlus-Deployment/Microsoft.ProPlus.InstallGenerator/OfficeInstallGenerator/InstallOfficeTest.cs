using System;
using System.Collections.Generic;
using System.Configuration;
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
using System.Globalization;
using System.Management;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.Win32;
using File = System.IO.File;

//[assembly: AssemblyTitle("")]
//[assembly: AssemblyProduct("")]
//[assembly: AssemblyDescription("")]
//[assembly: AssemblyVersion("")]
//[assembly: AssemblyFileVersion("")]

public class InstallOffice2
{

    private XmlDocument _xmlDoc = null;

    public static void Main1(string[] args)
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

    public void RunProgram()
    {
        var fileNames = new List<string>();
        var installDir = "";

        try
        {
            MinimizeWindow();

            SilentInstall = false;
            var startTime = DateTime.Now;

            var currentDirectory = Environment.ExpandEnvironmentVariables("%public%");
            installDir = currentDirectory + @"\OfficeProPlus";

            Directory.CreateDirectory(installDir);

            var args = GetArguments();
            if (args.Any() && Environment.UserName != "SYSTEM")
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
           

            if (!File.Exists(odtFilePath)) { throw (new Exception("Cannot find ODT Executable")); }
            if (!File.Exists(xmlFilePath)) { throw (new Exception("Cannot find Configuration Xml file")); }

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

                if (GetArguments().Any(a => a.Key.ToLower() == "/silent"))
                {
                    SilentInstall = true;
                }
            }

            if (runInstall)
            {

                if (SilentInstall)
                {
                    Console.WriteLine("Running Silent Install...");
                    var doc = new XmlDocument();
                    doc.Load(xmlFilePath);
                    SetConfigSilent(doc);
                    doc.Save(xmlFilePath);
                }

                try
                {
                    var officeProducts = GetOfficeVersion();
                    if (officeProducts != null)
                    {
                        var doc = new XmlDocument();
                        doc.Load(xmlFilePath);
                        SetClientEdition(doc, officeProducts.Edition);
                        doc.Save(xmlFilePath);
                    }
                } catch { }

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


    public OfficeInstalledProducts GetOfficeVersion()
    {
        var installKeys = new List<string>()
        {
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        };

        var officeKeys = new List<string>()
        {
            @"SOFTWARE\Microsoft\Office",
            @"SOFTWARE\Wow6432Node\Microsoft\Office"
        };

        string osArchitecture = null;
        var osClass = new ManagementClass("Win32_OperatingSystem");
        foreach (var queryObj in osClass.GetInstances())
        {
            foreach (var prop in queryObj.Properties)
            {
                if (prop.Name?.ToLower() == "OSArchitecture".ToLower())
                {
                    osArchitecture = prop.Value?.ToString();
                    if (osArchitecture == null) osArchitecture = "";
                }
            }
        }

        // $results = new-object PSObject[] 0;

        // foreach ($computer in $ComputerName) {
        //    if ($Credentials) {
        //       $os=Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials
        //    } else {
        //       $os=Get-WMIObject win32_operatingsystem -computername $computer
        //    }

        //    $osArchitecture = $os.OSArchitecture

        //    if ($Credentials) {
        //       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials
        //} else {
        //       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer
        //}

        var officePathReturn = GetOfficePathList();

        foreach (var regKey in installKeys)
        {
            var keyList = new List<string>();
            var keys = Registry.LocalMachine.OpenSubKey(regKey)?.GetSubKeyNames();

            foreach (var key in keys) {
                 var path = regKey + @"\" + key;
                 var installPath = Registry.LocalMachine.OpenSubKey(path)?.GetValue("InstallLocation")?.ToString();
                
                 if (string.IsNullOrEmpty(installPath))
                 {
                    continue;
                 }

                var buildType = "64-Bit";               
                if (osArchitecture == "32-bit") {
                    buildType = "32-Bit";
                }

                if (regKey.ToUpper().Contains("Wow6432Node".ToUpper()))
                {
                    buildType = "32-Bit";
                }

                if (Regex.Match(key, "{.{8}-.{4}-.{4}-1000-0000000FF1CE}").Success) {
                    buildType = "64-Bit";
                }

                if (Regex.Match(key, "{.{8}-.{4}-.{4}-0000-0000000FF1CE}").Success)
                {
                    buildType = "64-Bit";
                }

                var modifyPath = Registry.LocalMachine.OpenSubKey(path)?.GetValue("ModifyPath")?.ToString();

                if (!string.IsNullOrEmpty(modifyPath) ) {
                    if (modifyPath.ToLower().Contains("platform=x86")) {
                        buildType = "32-Bit";
                    }

                    if (modifyPath.ToLower().Contains("platform=x64"))
                    {
                        buildType = "64-Bit";
                    }
                }

                var primaryOfficeProduct = false;
                var officeProduct = false;
                foreach (var officeInstallPath in officePathReturn.PathList)
                {
                    if (!string.IsNullOrEmpty(officeInstallPath)) {
                        var installReg = "^" + installPath.Replace(@"\", @"\\");
                        installReg = installReg.Replace("(", @"\(");
                        installReg = installReg.Replace(@")", @"\)");

                        if (Regex.Match(officeInstallPath, installReg).Success)
                        {
                            officeProduct = true;
                        }
                    }
                }

                if (!officeProduct)
                {
                    continue;
                }

                var name = Registry.LocalMachine.OpenSubKey(path)?.GetValue("DisplayName")?.ToString();
                if (name == null) name = "";

                if (officePathReturn.ConfigItemList.Contains(key.ToUpper()) && name.ToUpper().Contains("MICROSOFT OFFICE"))
                {
                    primaryOfficeProduct = true;
                }

                var version = Registry.LocalMachine.OpenSubKey(path)?.GetValue("DisplayVersion")?.ToString();
                modifyPath = Registry.LocalMachine.OpenSubKey(path)?.GetValue("ModifyPath")?.ToString();

                var cltrUpdatedEnabled = "";
                var cltrUpdateUrl = "";
                var clientCulture = "";

                var clickToRun = false;
                if (officePathReturn.ClickToRunPathList.Contains(installPath?.ToUpper()))
                {
                    clickToRun = true;
                    if (name.ToUpper().Contains("MICROSOFT OFFICE"))
                    {
                        primaryOfficeProduct = true;
                    }

                    foreach (var cltr in officePathReturn.ClickToRunList)
                    {
                        if (!string.IsNullOrEmpty(cltr.InstallPath))
                        {
                            if (cltr.InstallPath.ToUpper() == installPath.ToUpper())
                            {
                                if (cltr.Bitness == "x64") {
                                    buildType = "64-Bit";
                                }
                                if (cltr.Bitness == "x86") {
                                    buildType = "32-Bit";
                                }
                                clientCulture = cltr.ClientCulture;
                            }
                        }
                    }
               }

                var offInstall = new OfficeInstall
                {
                    DisplayName = name,
                    Version = version,
                    InstallPath = installPath,
                    ClickToRun = clickToRun,
                    Bitness = buildType,
                    ClientCulture = clientCulture
                };
                officePathReturn.ClickToRunList.Add(offInstall);
                //}
            }

        }

        var returnList = officePathReturn.ClickToRunList.Distinct().ToList();

        return new OfficeInstalledProducts()
        {
            OfficeInstalls = returnList,
            OSArchitecture = osArchitecture
        };
    }

    public OfficePathsReturn GetOfficePathList()
    {
        var officeKeys = new List<string>()
        {
            @"SOFTWARE\Microsoft\Office",
            @"SOFTWARE\Wow6432Node\Microsoft\Office"
        };

        var clickToRunList = new List<OfficeInstall>();

        var pathReturn = new OfficePathsReturn()
        {
            ClickToRunList = new List<OfficeInstall>(),
            VersionList = new List<string>(),
            PathList = new List<string>(),
            PackageList = new List<string>(),
            ClickToRunPathList = new List<string>(),
            ConfigItemList = new List<string>()
        };


        foreach (var regKey in officeKeys)
        {
            var officeVersion = GetRegistrySubKeys(regKey);
            var c2RRegPath = regKey + @"\ClickToRun\Configuration";
            var c2R16Key = GetRegistryKey(c2RRegPath);
            if (c2R16Key != null)
            {
                clickToRunList.Add(new OfficeInstall
                {
                    InstallPath = GetRegistryValue(c2RRegPath, "InstallationPath"),
                    Bitness = GetRegistryValue(c2RRegPath, "Platform"),
                    ClientCulture = GetRegistryValue(c2RRegPath, "ClientCulture"),
                    ClickToRun = true
                });
            }

            foreach (var key in officeVersion)
            {
                var match = Regex.Match(key, @"\d{2}\.\d");
                if (match.Success)
                {
                    if (!pathReturn.VersionList.Contains(key))
                    {
                        pathReturn.VersionList.Add(key);
                    }

                    var path = regKey + @"\" + key;
                    var configPath = path + @"\Common\Config";

                    var configKey = Registry.LocalMachine.OpenSubKey(configPath);
                    var configItems = configKey?.GetSubKeyNames();

                    if (configItems != null)
                    {
                        pathReturn.ConfigItemList.AddRange(from configId in configItems where !string.IsNullOrEmpty(configId) select configId.ToUpper());
                    }

                    var cltr = new OfficeInstall();

                    var packagePath = path + @"\Common\InstalledPackages";
                    var clickToRunPath = path + @"\ClickToRun\Configuration";
                    var officeLangResourcePath = path + @"\Common\LanguageResources";
                    var cultures = CultureInfo.GetCultures(CultureTypes.AllCultures);

                    var virtualInstallPath = GetRegistryValue(clickToRunPath, "InstallationPath")?.ToString();
                    var mainLangId = GetRegistryValue(officeLangResourcePath, "SKULanguage")?.ToString();
                    if (string.IsNullOrEmpty(mainLangId))
                    {
                        var mainlangCulture = cultures.FirstOrDefault(c => c.LCID.ToString() == mainLangId);

                        if (mainlangCulture != null)
                        {
                            cltr.ClientCulture = mainlangCulture.Name;
                        }
                    }

                    var officeLangPath = path + @"\Common\LanguageResources\InstalledUIs";
                    var offLangPathKey = Registry.LocalMachine.OpenSubKey(officeLangPath);
                    var langValues = offLangPathKey?.GetSubKeyNames();

                    CultureInfo langCulture = null;
                    if (langValues != null)
                    {
                        foreach (var langValue in langValues)
                        {
                            langCulture = cultures.FirstOrDefault(c => c.LCID.ToString() == langValue);
                        }
                    }

                    if (string.IsNullOrEmpty(virtualInstallPath))
                    {
                        clickToRunPath = regKey + @"\ClickToRun\Configuration";
                        virtualInstallPath = GetRegistryValue(clickToRunPath, "InstallationPath");
                    }

                    if (!string.IsNullOrEmpty(virtualInstallPath))
                    {
                        if (!pathReturn.ClickToRunPathList.Contains(virtualInstallPath?.ToUpper()))
                        {
                            pathReturn.ClickToRunPathList.Add(virtualInstallPath?.ToUpper());
                        }

                        cltr.InstallPath = virtualInstallPath;
                        cltr.Bitness = GetRegistryValue(clickToRunPath, "Platform");
                        cltr.ClientCulture = GetRegistryValue(clickToRunPath, "ClientCulture");
                        cltr.ClickToRun = true;
                        clickToRunList.Add(cltr);
                    }

                    var packageItems = Registry.LocalMachine.OpenSubKey(packagePath)?.GetSubKeyNames();
                    var officeItems = Registry.LocalMachine.OpenSubKey(path)?.GetSubKeyNames();

                    foreach (var itemKey in officeItems)
                    {
                        var itemPath = path + @"\" + itemKey;
                        var installRootPath = itemPath + @"\InstallRoot";

                        //HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot

                        var filePath = GetRegistryValue(installRootPath, "Path");

                        if (string.IsNullOrEmpty(filePath)) continue;

                        if (!pathReturn.PathList.Contains(filePath))
                        {
                            pathReturn.PathList.Add(filePath);
                        }
                    }

                    if (packageItems != null)
                    {
                        foreach (var packageGuid in packageItems)
                        {
                            var packageItemPath = packagePath + @"\" + packageGuid;
                            var packageName = Registry.LocalMachine.OpenSubKey(packageItemPath)?.GetValue(null)?.ToString();
                            if (!pathReturn.PackageList.Contains(packageName))
                            {
                                if (!string.IsNullOrEmpty(packageName))
                                {
                                    pathReturn.PackageList.Add(packageName.Replace(" ", "").ToLower());
                                }
                            }
                        }
                    }

                }
            }
        }

        return pathReturn;
    }


    private RegistryKey GetRegistryKey(string keyPath)
    {
        var key = Registry.LocalMachine.OpenSubKey(keyPath);
        return key;
        //using (var key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(keyPath))
        //{
        //    var value = key?.GetValue(property)?.ToString();
        //    if (!string.IsNullOrEmpty(value))
        //    {
        //        return value;
        //    }
        //}
    }

    private List<string> GetRegistrySubKeys(string keyPath)
    {
        using (var key = Registry.LocalMachine.OpenSubKey(keyPath))
        {
            var subKeyList = key?.GetSubKeyNames().ToList();
            if (subKeyList != null) return subKeyList;
        }

        //using (var key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(keyPath))
        //{
        //    var value = key?.GetValue(property)?.ToString();
        //    if (!string.IsNullOrEmpty(value))
        //    {
        //        return value;
        //    }
        //}

        return new List<string>();
    }

    private string GetRegistryValue(string keyPath, string property)
    {
        using (var key = Registry.LocalMachine.OpenSubKey(keyPath))
        {
            var value = key?.GetValue(property)?.ToString();
            if (!string.IsNullOrEmpty(value))
            {
                return value;
            }
        }

        //using (var key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(keyPath))
        //{
        //    var value = key?.GetValue(property)?.ToString();
        //    if (!string.IsNullOrEmpty(value))
        //    {
        //        return value;
        //    }
        //}
        return "";
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

    private bool HasValidArguments()
    {
        return !GetArguments().Any(a => (a.Key.ToLower() != "/uninstall" &&
                                         a.Key.ToLower() != "/showxml" &&
                                         a.Key.ToLower() != "/silent" &&
                                         a.Key.ToLower() != "/extractxml"));
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

    private void SetClientEdition(XmlDocument doc, OfficeClientEdition edition)
    {
        var add = doc.SelectSingleNode("/Configuration/Add");
        if (add == null) return;
        switch (edition)
        {
            case OfficeClientEdition.Office32Bit:
                SetAttribute(doc, add, "OfficeClientEdition", "32");
                break;
            case OfficeClientEdition.Office64Bit:
                SetAttribute(doc, add, "OfficeClientEdition", "64");
                break;
        }
    }

    private void SetConfigSilent(XmlDocument doc)
    {
        var display = doc.SelectSingleNode("/Configuration/Display");
        if (display == null)
        {
            display = doc.CreateElement("Display");
            doc.DocumentElement.AppendChild(display);
        }

        SetAttribute(doc, display, "Level", "None");
        SetAttribute(doc, display, "AcceptEULA", "TRUE");
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

    public static string ResourcePath
    {
        get
        {
            return Directory.GetCurrentDirectory();
        }
    }

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

    public List<ExecutingScenario> GetRunningScenarioTasks()
    {
        var execScenarios = new List<ExecutingScenario>();
        var executingScenario = GetExecutingScenario();
        if (string.IsNullOrEmpty(executingScenario)) return new List<ExecutingScenario>();

        var mainRegPath = GetOfficeCtrRegPath();
        if (string.IsNullOrEmpty(mainRegPath)) return new List<ExecutingScenario>();

        var scenarioPath = mainRegPath + @"\scenario";

        var scenarioKey = Registry.LocalMachine.OpenSubKey(scenarioPath);
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

    public string GetExecutingScenario()
    {
        var mainRegPath = GetOfficeCtrRegPath();
        if (mainRegPath == null) return null;
        var configKey = Registry.LocalMachine.OpenSubKey(mainRegPath);
        if (configKey == null) return null;
        var execScenario = configKey.GetValue("ExecutingScenario");
        return execScenario != null ? execScenario.ToString() : null;
    }

    public string GetOfficeCtrRegPath()
    {
        var path16 = @"SOFTWARE\Microsoft\Office\ClickToRun";
        var path15 = @"SOFTWARE\Microsoft\Office\15.0\ClickToRun";


        var office16Key = Registry.LocalMachine.OpenSubKey(path16);
        var office15Key = Registry.LocalMachine.OpenSubKey(path15);


        if (office16Key != null)
        {
            return path16;
        }
        else
        {
            if (office15Key != null)
            {
                return path15;
            }
        }
        return null;
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
                    File.Copy(file.FullName, Environment.ExpandEnvironmentVariables(@"%public%\" + file.Name), true);
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
        var tempPath = Environment.ExpandEnvironmentVariables("%public%");
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
        var tempPath = Environment.ExpandEnvironmentVariables("%public%");
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



    public string LoggingPath { get; set; }

    public bool SilentInstall { get; set; }

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
        return CurrentOperation.Starting;
    }

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

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    public const int SW_SHOWMINIMIZED = 2;

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
}


public class OfficeInstalledProducts
{
    public List<OfficeInstall> OfficeInstalls { get; set; }

    public string OSArchitecture { get; set; }

    public OfficeClientEdition Edition
    {
        get
        {
            if (OSArchitecture != null)
            {
                if (OSArchitecture.Contains("32"))
                {
                    return OfficeClientEdition.Office32Bit;
                }
            }

            var installed64Bits = OfficeInstalls.Count(i => i.Bitness != null && i.Bitness.Contains("64"));
            var installed32Bits = OfficeInstalls.Count(i => i.Bitness != null && i.Bitness.Contains("32"));
            return installed64Bits > installed32Bits ? OfficeClientEdition.Office64Bit : OfficeClientEdition.Office32Bit;

        }
    }    
}

public class OfficeInstall
{
    public string DisplayName { get; set; }

    public bool ClickToRun { get; set; }

    public string Bitness { get; set; }

    public string Version { get; set; }

    public string InstallPath { get; set; }

    public string ClientCulture { get; set; }

}


public class OfficePathsReturn
{
    public List<string> VersionList { get; set; }

    public List<string> PathList { get; set; }

    public List<string> PackageList { get; set; }

    public List<string> ClickToRunPathList { get; set; }

    public List<string> ConfigItemList { get; set; }

    public List<OfficeInstall> ClickToRunList { get; set; }
}