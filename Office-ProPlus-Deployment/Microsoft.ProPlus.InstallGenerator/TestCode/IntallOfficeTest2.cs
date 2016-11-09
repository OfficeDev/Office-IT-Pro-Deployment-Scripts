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
using System.Globalization;
using System.Management;
using System.Windows;
using System.Windows.Forms;
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
        SilentInstall = false;
        try
        {
            //MinimizeWindow();

            Initialize();

            SetLoggingPath(XmlFilePath);
            SetSourcePath(XmlFilePath);

            var runInstall = false;
            switch (Operation)
            {
                case OperationType.ShowXml:
                    ShowXml(XmlFilePath);
                    break;
                case OperationType.ExtractXml:
                    ExtractXml(XmlFilePath);
                    break;
                case OperationType.Uninstall:
                    XmlFilePath = UninstallOfficeProPlus(InstallDirectory, FileNames);
                    runInstall = true;
                    break;
                case OperationType.Install:
                    runInstall = true;
                    UpdateLanguagePackInstall(XmlFilePath);
                    break;
            }

            if (runInstall)
            {
                RunInstall(OdtFilePath, XmlFilePath);
            }
        }
        finally
        {
            //CleanUp(InstallDirectory);
        }

    }

    #region Main Operations

    private int RunInstall(string odtFilePath, string xmlFilePath)
    {
        try
        {
            StopWus();

            if (DetermineIfLanguageInstalled(false))
            {
                return 0;
            }

            Console.WriteLine("Installing Office 365 ProPlus...");
            try
            {
                var officeProducts = GetOfficeVersion();
                if (officeProducts != null)
                {
                    foreach (var product in officeProducts.OfficeInstalls)
                    {
                        var displayName = product.DisplayName;
                        if (displayName == null) displayName = "";
                    }

                    Console.WriteLine(officeProducts.Edition);

                    var doc = new XmlDocument();
                    doc.Load(xmlFilePath);

                    if (officeProducts.Edition.HasValue)
                    {
                        SetClientEdition(doc, officeProducts.Edition.Value);
                    }

                    doc.Save(xmlFilePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Get Version Error: " + ex.Message);
            }

            var ucMapiProcess = Process.GetProcessesByName("ucmapi.exe");
            if (ucMapiProcess != null && ucMapiProcess.Length > 0)
            {
                foreach (var ucmProcess in ucMapiProcess)
                {
                    ucmProcess.Kill();
                }
            }

            var doc1 = new XmlDocument();
            doc1.Load(xmlFilePath);
            if (SilentInstall)
            {
                Console.WriteLine("Running Silent Install...");
                SetDisplayLevel(doc1);
            }
            else
            {
                SetDisplayLevel(doc1, false);
            }
            doc1.Save(xmlFilePath);

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

            DetermineIfLanguageInstalled();

            return p.ExitCode;
        }
        catch (Exception ex)
        {
            RollBackInstall();
            throw;
        }
        finally
        {
            StartWus();
        }
    }

    public bool DetermineIfLanguageInstalled(bool throwException = true)
    {
        if (!IsLanguagePackInstall(XmlFilePath)) return false;
        var languages = GetLanguagePackLanguages(XmlFilePath);

        foreach (var language in languages)
        {
            var languageInstalled = IsLanguageInstalled(XmlFilePath, language);
            if (Operation == OperationType.Install)
            {
                if (!languageInstalled)
                {
                    if (throwException)
                    {
                        throw (new Exception("Language not Installed: " + language));
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return true;
                }
            }
            else if (Operation == OperationType.Uninstall)
            {
                if (languageInstalled)
                {
                    if (throwException)
                    {
                        throw (new Exception("Language still Installed: " + language));
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return true;
                }
            }
        }
        return false;
    }

    public void RollBackInstall()
    {
        if (!string.IsNullOrEmpty(ProductId))
        {
            Console.WriteLine("Rolling Back Install...");
            ProductId = ProductId.Replace("{", "").Replace("}", "");

            RegistryKey installKey = null;
            RegistryKey uninstallKey = null;
            for (var t = 1; t <= 3; t++)
            {
                switch (t)
                {
                    case 1:
                        uninstallKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                        break;
                    case 2:
                        uninstallKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                                     .OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                        break;
                    case 3:
                        uninstallKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                                     .OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                        break;
                }

                if (uninstallKey != null)
                {
                    foreach (var subKey in uninstallKey.GetSubKeyNames())
                    {
                        var tmpKey = uninstallKey.OpenSubKey(subKey);
                        if (tmpKey == null) continue;
                        if (tmpKey.GetValue("Comments") == null) continue;
                        var comments = tmpKey.GetValue("Comments").ToString();
                        comments = comments.Replace("{", "").Replace("}", "");
                        if (comments.ToUpper() == ProductId.ToUpper())
                        {
                            installKey = tmpKey;
                            break;
                        }
                    }
                }

            }

            if (installKey != null)
            {
                var tmpUninstallString = installKey.GetValue("UninstallString");
                if (tmpUninstallString == null) return;
                var uninstallString = tmpUninstallString.ToString();

                uninstallString = Regex.Replace(uninstallString, @"/I", @"/X", RegexOptions.IgnoreCase);

                var p = new Process
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = "cmd",
                        Arguments = "/c start cmd /c " + uninstallString + " /qb",
                        UseShellExecute = true
                    },
                };
                p.Start();
                p.WaitForExit();
            }
        }
    }

    private string UninstallOfficeProPlus(string installationDirectory, IEnumerable<string> fileNames)
    {
        var doc = new XmlDocument();

        if (IsLanguagePackInstall(XmlFilePath))
        {
            Console.WriteLine("Uninstalling Office 365 ProPlus Language Pack...");

            UpdateLanguagePackInstall(XmlFilePath, true);
            doc.Load(XmlFilePath);
        }
        else
        {
            Console.WriteLine("Uninstalling Office 365 ProPlus...");

            var root = doc.CreateElement("Configuration");
            var remove1 = doc.CreateElement("Remove");
            var all = doc.CreateAttribute("All");
            all.Value = "TRUE";
            remove1.Attributes.Append(all);
            root.AppendChild(remove1);
            doc.AppendChild(root);
        }

        if (SilentInstall)
        {
            SetDisplayLevel(doc);
        }

        doc.Save(installationDirectory + @"\configuration.xml");

        return installationDirectory + @"\" + fileNames.FirstOrDefault(f => f.ToLower().EndsWith(".xml"));
    }

    private void ShowXml(string xmlFilePath)
    {
        Console.Clear();
        var configXml = File.ReadAllText(xmlFilePath);
        Console.WriteLine(BeautifyXml(configXml));
    }

    private void ExtractXml(string xmlFilePath)
    {
        var arg = GetArguments().FirstOrDefault(a => a.Key.ToLower() == "/extractxml");
        if (string.IsNullOrEmpty(arg.Value)) Console.WriteLine("ERROR: Invalid File Path");
        var configXml = BeautifyXml(File.ReadAllText(xmlFilePath));
        File.WriteAllText(arg.Value, configXml);
    }

    private void Initialize()
    {
        FindTempFilesPath();

        InstallDirectory = TempFilesPath + @"\OfficeProPlus";
        Directory.CreateDirectory(InstallDirectory);

        OS = Environment.OSVersion;

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
        FileNames = GetEmbeddedItems(InstallDirectory);
        Console.WriteLine("Done");

        OdtFilePath = InstallDirectory + @"\" + FileNames.FirstOrDefault(f => f.ToLower().EndsWith("setup.exe"));
        XmlFilePath = InstallDirectory + @"\" + FileNames.FirstOrDefault(f => f.ToLower().EndsWith(".xml"));

        var sourceFilePath = GetArguments().FirstOrDefault(a => a.Key.ToLower() == "/sourfilepath");

        var chkPath = InstallDirectory + @"\Office\Data";
        if (!Directory.Exists(chkPath) && sourceFilePath != null)
        {
            if (!string.IsNullOrEmpty(sourceFilePath.Value))
            {
                SetSourcePath(XmlFilePath, sourceFilePath.Value);
            }
        }

        var productIdFile = InstallDirectory + @"\" + FileNames.FirstOrDefault(f => f.ToLower().EndsWith("productid.txt"));
        if (!string.IsNullOrEmpty(productIdFile))
        {
            if (File.Exists(productIdFile))
            {
                ProductId = File.ReadAllText(productIdFile);
            }
        }

        if (!File.Exists(OdtFilePath)) { throw (new Exception("Cannot find ODT Executable")); }
        if (!File.Exists(XmlFilePath)) { throw (new Exception("Cannot find Configuration Xml file")); }

        if (GetArguments().Any(a => a.Key.ToLower() == "/silent"))
        {
            SilentInstall = true;
        }

        var productArg = GetArguments().FirstOrDefault(a => a.Key.ToLower() == "/productid");
        if (productArg != null)
        {
            ProductId = productArg.Value;
        }

        if (GetArguments().Any(a => a.Key.ToLower() == "/uninstall"))
        {
            Operation = OperationType.Uninstall;
        }
        else if (GetArguments().Any(a => a.Key.ToLower() == "/showxml"))
        {
            Operation = OperationType.ShowXml;
        }
        else if (GetArguments().Any(a => a.Key.ToLower() == "/extractxml"))
        {
            Operation = OperationType.ExtractXml;
        }
        else
        {
            Operation = OperationType.Install;
        }
    }

    private void FindTempFilesPath()
    {
        var windirTemp = Environment.ExpandEnvironmentVariables(@"%windir%\Temp");
        if (Directory.Exists(windirTemp))
        {
            try
            {
                var tempDirectory = windirTemp + @"\OfficeProPlus";
                Directory.CreateDirectory(tempDirectory);
                TempFilesPath = windirTemp;
                return;
            }
            catch (Exception ex)
            {

            }
        }
        TempFilesPath = Environment.ExpandEnvironmentVariables("%public%");
    }

    #endregion

    #region Configuration XML

    private void SetDisplayLevel(XmlDocument doc, bool silent = true)
    {
        var display = doc.SelectSingleNode("/Configuration/Display");
        if (display == null)
        {
            display = doc.CreateElement("Display");
            doc.DocumentElement.AppendChild(display);
        }

        if (silent)
        {
            SetAttribute(doc, display, "Level", "None");
        }
        else
        {
            SetAttribute(doc, display, "Level", "Full");
        }
        SetAttribute(doc, display, "AcceptEULA", "TRUE");
    }

    private OfficeClientEdition GetClientEdition(XmlDocument doc)
    {
        var add = doc.SelectSingleNode("/Configuration/Add");
        if (add == null) return OfficeClientEdition.Office32Bit;
        if (doc.Attributes == null || doc.Attributes.Count == 0) return OfficeClientEdition.Office32Bit;

        var currentValue = GetAttribute(doc, add, "OfficeClientEdition");
        if (currentValue == "32")
        {
            return OfficeClientEdition.Office32Bit;
        }
        else
        {
            return OfficeClientEdition.Office64Bit;
        }
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

    private void SetLoggingPath(string xmlFilePath)
    {
        const string logFolderName = "OfficeProPlusLogs";
        LoggingPath = TempFilesPath + @"\" + logFolderName;
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

    private void SetSourcePath(string xmlFilePath, string sourcePath = null)
    {
        const string officeFolderName = "OfficeProPlus";

        var officeFolderPath = TempFilesPath + @"\" + officeFolderName;
        if (Directory.Exists(officeFolderPath + @"\Office"))
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFilePath);

            var addNode = xmlDoc.SelectSingleNode("/Configuration/Add");
            if (addNode != null)
            {
                SetAttribute(xmlDoc, addNode, "SourcePath", sourcePath ?? officeFolderPath);

                xmlDoc.Save(xmlFilePath);
            }
        }
    }

    private bool IsLanguagePackInstall(string xmlFilePath)
    {
        var xmlDoc = new XmlDocument();
        xmlDoc.Load(xmlFilePath);

        var languagePack = xmlDoc.SelectSingleNode("/Configuration/Add/Product[@ID='LanguagePack']");
        if (languagePack != null)
        {
            return true;
        }

        languagePack = xmlDoc.SelectSingleNode("/Configuration/Remove/Product[@ID='LanguagePack']");
        if (languagePack != null)
        {
            return true;
        }

        return false;
    }

    private void UpdateLanguagePackClientCulture(string xmlFilePath, string clientCulture, bool convertToRemove = false)
    {
        var xmlDoc = new XmlDocument();
        xmlDoc.Load(xmlFilePath);

        var addNode = xmlDoc.SelectSingleNode("/Configuration/Add");

        var languagePack = xmlDoc.SelectSingleNode("/Configuration/Add/Product[@ID='LanguagePack']") ??
                           xmlDoc.SelectSingleNode("/Configuration/Remove/Product[@ID='LanguagePack']");
        if (languagePack == null) return;

        var languageList = new List<string>();
        var languages = languagePack.SelectNodes("./Language");
        foreach (XmlNode language in languages)
        {
            var langId = GetAttribute(xmlDoc, language, "ID");
            if (!string.IsNullOrEmpty(langId))
            {
                languageList.Add(langId);
            }
        }

        for (var i = 0; i < languages.Count; i++)
        {
            var language = languages[i];
            languagePack.RemoveChild(language);
        }

        if (convertToRemove && addNode != null)
        {
            var removeNode = xmlDoc.CreateElement("Remove");
            languagePack = xmlDoc.CreateElement("Product");
            SetAttribute(xmlDoc, languagePack, "ID", "LanguagePack");
            removeNode.AppendChild(languagePack);
            xmlDoc.DocumentElement.PrependChild(removeNode);
        }

        if (!convertToRemove)
        {
            AddLanguage(xmlDoc, languagePack, clientCulture);
        }
        else
        {
            if (addNode != null)
            {
                xmlDoc.DocumentElement.RemoveChild(addNode);
            }
        }

        var otherLangs = languageList.Where(l => l.ToLower() != clientCulture.ToLower());
        if (otherLangs.Any())
        {
            foreach (var otherLang in otherLangs)
            {
                AddLanguage(xmlDoc, languagePack, otherLang);
            }
        }

        xmlDoc.Save(xmlFilePath);
    }

    private List<string> GetLanguagePackLanguages(string xmlFilePath)
    {
        var languageList = new List<string>();
        var xmlDoc = new XmlDocument();
        xmlDoc.Load(xmlFilePath);

        var languagePack = xmlDoc.SelectSingleNode("/Configuration/Add/Product[@ID='LanguagePack']") ??
                           xmlDoc.SelectSingleNode("/Configuration/Remove/Product[@ID='LanguagePack']");
        if (languagePack == null) return new List<string>();

        var languages = languagePack.SelectNodes("./Language");
        foreach (XmlNode language in languages)
        {
            var langId = GetAttribute(xmlDoc, language, "ID");
            if (!string.IsNullOrEmpty(langId))
            {
                languageList.Add(langId);
            }
        }

        return languageList;
    }


    private void AddLanguage(XmlDocument xmlDoc, XmlNode productNode, string languageId)
    {
        var languageNode = xmlDoc.CreateElement("Language");
        SetAttribute(xmlDoc, languageNode, "ID", languageId);
        productNode.AppendChild(languageNode);
    }

    #endregion

    #region File Operations

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

    #region Office Registry

    public void UpdateLanguagePackInstall(string xmlFilePath, bool convertToRemove = false)
    {
        var regPath = GetOfficeCtrRegPath();
        var configKey = Registry.LocalMachine.OpenSubKey(regPath + @"\Configuration");
        if (configKey == null) return;

        var clientCulture = GetRegistryValue(regPath + @"\Configuration", "ClientCulture");
        if (string.IsNullOrEmpty(clientCulture)) return;

        if (IsLanguagePackInstall(xmlFilePath))
        {
            UpdateLanguagePackClientCulture(xmlFilePath, clientCulture, convertToRemove);
        }
    }

    public bool IsLanguageInstalled(string xmlFilePath, string language)
    {
        var regPath = GetOfficeCtrRegPath();
        var prodKey = Registry.LocalMachine.OpenSubKey(regPath + @"\ProductReleaseIDs");
        if (prodKey == null) return false;

        foreach (var subKey in prodKey.GetSubKeyNames())
        {
            var cultureKey = prodKey.OpenSubKey(subKey + @"\culture");
            if (cultureKey == null) continue;

            var langKeys = cultureKey.GetSubKeyNames();

            var langExists = langKeys.Any(k => k.ToLower().StartsWith(language.ToLower()));
            if (langExists)
            {
                return true;
            }
        }
        return false;
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
                if (prop.Name == null) continue;
                if (prop.Name.ToLower() == "OSArchitecture".ToLower())
                {
                    if (prop.Value == null) continue;
                    osArchitecture = prop.Value.ToString() ?? "";
                    break;
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
            var keys = GetRegistrySubKeys(regKey);

            foreach (var key in keys)
            {
                var path = regKey + @"\" + key;
                var installPath = GetRegistryValue(path, "InstallLocation");

                if (string.IsNullOrEmpty(installPath))
                {
                    continue;
                }

                var buildType = "64-Bit";
                if (osArchitecture == "32-bit")
                {
                    buildType = "32-Bit";
                }

                if (regKey.ToUpper().Contains("Wow6432Node".ToUpper()))
                {
                    buildType = "32-Bit";
                }

                if (Regex.Match(key, "{.{8}-.{4}-.{4}-1000-0000000FF1CE}").Success)
                {
                    buildType = "64-Bit";
                }

                if (Regex.Match(key, "{.{8}-.{4}-.{4}-0000-0000000FF1CE}").Success)
                {
                    buildType = "64-Bit";
                }


                var modifyPath = GetRegistryValue(path, "ModifyPath");
                if (!string.IsNullOrEmpty(modifyPath))
                {
                    if (modifyPath.ToLower().Contains("platform=x86"))
                    {
                        buildType = "32-Bit";
                    }

                    if (modifyPath.ToLower().Contains("platform=x64"))
                    {
                        buildType = "64-Bit";
                    }
                }


                var officeProduct = false;
                foreach (var officeInstallPath in officePathReturn.PathList)
                {
                    if (!string.IsNullOrEmpty(officeInstallPath))
                    {
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


                var name = GetRegistryValue(path, "DisplayName");
                if (name == null) name = "";

                if (officePathReturn.ConfigItemList.Contains(key.ToUpper()) && name.ToUpper().Contains("MICROSOFT OFFICE"))
                {
                    //primaryOfficeProduct = true;
                }

                var version = GetRegistryValue(path, "DisplayVersion");
                modifyPath = GetRegistryValue(path, "ModifyPath");

                var clientCulture = "";

                if (installPath == null) installPath = "";

                var clickToRun = false;
                if (officePathReturn.ClickToRunPathList.Contains(installPath.ToUpper()))
                {
                    clickToRun = true;
                    if (name.ToUpper().Contains("MICROSOFT OFFICE"))
                    {
                        //primaryOfficeProduct = true;
                    }

                    foreach (var cltr in officePathReturn.ClickToRunList)
                    {
                        if (!string.IsNullOrEmpty(cltr.InstallPath))
                        {
                            if (cltr.InstallPath.ToUpper() == installPath.ToUpper())
                            {
                                if (cltr.Bitness == "x64")
                                {
                                    buildType = "64-Bit";
                                }
                                if (cltr.Bitness == "x86")
                                {
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

                    var configItems = GetRegistrySubKeys(configPath);
                    if (configItems != null)
                    {
                        pathReturn.ConfigItemList.AddRange(from configId in configItems where !string.IsNullOrEmpty(configId) select configId.ToUpper());
                    }

                    var cltr = new OfficeInstall();

                    var packagePath = path + @"\Common\InstalledPackages";
                    var clickToRunPath = path + @"\ClickToRun\Configuration";
                    var officeLangResourcePath = path + @"\Common\LanguageResources";
                    var cultures = CultureInfo.GetCultures(CultureTypes.AllCultures);

                    var virtualInstallPath = GetRegistryValue(clickToRunPath, "InstallationPath");
                    var mainLangId = GetRegistryValue(officeLangResourcePath, "SKULanguage");
                    if (string.IsNullOrEmpty(mainLangId))
                    {
                        var mainlangCulture = cultures.FirstOrDefault(c => c.LCID.ToString() == mainLangId);

                        if (mainlangCulture != null)
                        {
                            cltr.ClientCulture = mainlangCulture.Name;
                        }
                    }

                    var officeLangPath = path + @"\Common\LanguageResources\InstalledUIs";
                    var langValues = GetRegistrySubKeys(officeLangPath);

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
                        if (virtualInstallPath == null) virtualInstallPath = "";
                        if (!pathReturn.ClickToRunPathList.Contains(virtualInstallPath.ToUpper()))
                        {
                            pathReturn.ClickToRunPathList.Add(virtualInstallPath.ToUpper());
                        }

                        cltr.InstallPath = virtualInstallPath;
                        cltr.Bitness = GetRegistryValue(clickToRunPath, "Platform");
                        cltr.ClientCulture = GetRegistryValue(clickToRunPath, "ClientCulture");
                        cltr.ClickToRun = true;
                        clickToRunList.Add(cltr);
                    }

                    var packageItems = GetRegistrySubKeys(packagePath);
                    var officeItems = GetRegistrySubKeys(path);

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
                            var packageName = GetRegistryValue(packageItemPath, null);
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
        using (var configKey = Registry.LocalMachine.OpenSubKey(mainRegPath))
        {
            if (configKey == null) return null;
            var execScenario = configKey.GetValue("ExecutingScenario");
            return execScenario != null ? execScenario.ToString() : null;
        }
    }

    public string GetOfficeCtrRegPath()
    {
        var path16 = @"SOFTWARE\Microsoft\Office\ClickToRun";
        var path15 = @"SOFTWARE\Microsoft\Office\15.0\ClickToRun";

        using (var office16Key = Registry.LocalMachine.OpenSubKey(path16))
        using (var office15Key = Registry.LocalMachine.OpenSubKey(path15))
        {
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
    }

    #endregion

    #region Properties

    public string TempFilesPath { get; set; }

    public OperationType Operation { get; set; }

    public OperatingSystem OS { get; set; }

    public string ProductId { get; set; }

    public string RollBackFilePath { get; set; }

    public string OdtFilePath { get; set; }

    public string XmlFilePath { get; set; }

    public List<string> FileNames { get; set; }

    public string InstallDirectory { get; set; }

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

    #region Operations

    private void StopWus()
    {
        if (OS.Version.Major == 6 && OS.Version.Minor == 1)
        {
            ToggleWus(true);
        }
    }

    private void StartWus()
    {
        if (OS.Version.Major == 6 && OS.Version.Minor == 1)
        {
            ToggleWus(true);
        }
    }

    private void ToggleWus(bool toggle)
    {
        var args = "/C net start wuauserv";
        if (toggle)
        {
            args = "/C net stop wuauserv";
        }

        var p = new Process
        {
            StartInfo =
            {
                FileName = "CMD.exe",
                Arguments = args
            }
        };

        p.Start();
        p.WaitForExit();
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
                    File.Copy(file.FullName, TempFilesPath + @"\" + file.Name, true);
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
        //var dirInfo = new DirectoryInfo(installDir);
        //foreach (var file in dirInfo.GetFiles())
        //{
        //    try
        //    {
        //        file.Delete();
        //    }
        //    catch { }
        //}

        //foreach (var directory in dirInfo.GetDirectories())
        //{
        //    try
        //    {
        //        directory.Delete(true);
        //    }
        //    catch { }
        //}

        //try
        //{
        //    Directory.Delete(installDir, true);
        //}
        //catch { }
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

    #endregion

    #region XML

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

    private string GetAttribute(XmlDocument xmlDoc, XmlNode xmlNode, string name)
    {
        var pathAttr = xmlNode.Attributes[name];
        if (pathAttr != null)
        {
            var value = xmlNode.Attributes[name].Value;
            if (!string.IsNullOrEmpty(value)) return value.ToString();
        }
        return "";
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

    #endregion

    #region Registry

    private RegistryKey GetRegistryKey(string keyPath)
    {
        var key = Registry.LocalMachine.OpenSubKey(keyPath, true);
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
        using (var key = Registry.LocalMachine.OpenSubKey(keyPath, true))
        {
            if (key == null) return new List<string>();
            var subKeyList = key.GetSubKeyNames().ToList();
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
        using (var key = Registry.LocalMachine.OpenSubKey(keyPath, true))
        {
            if (key == null) return "";
            var objValue = key.GetValue(property);
            if (objValue == null) return "";
            var value = objValue.ToString();
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

    #endregion

    #region Windows Functions

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    public const int SW_SHOWMINIMIZED = 2;

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

public class OfficeInstalledProducts
{
    public List<OfficeInstall> OfficeInstalls { get; set; }

    public string OSArchitecture { get; set; }

    private OfficeClientEdition _setClientEdition = OfficeClientEdition.Office32Bit;
    public OfficeClientEdition? Edition
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

            var installsFiltered = OfficeInstalls.Where(i => !i.DisplayName.ToLower().Contains("mui") &&
                                                             !i.DisplayName.ToLower().Contains("shared") &&
                                                             !i.DisplayName.ToLower().Contains("license")).ToList();
            if (installsFiltered.Any(i => i.Bitness.Contains("32")))
            {
                return OfficeClientEdition.Office32Bit;
            }

            if (installsFiltered.Any(i => i.Bitness.Contains("64")))
            {
                return OfficeClientEdition.Office64Bit;
            }

            return null;
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

public enum OfficeClientEdition
{
    Office32Bit = 0,
    Office64Bit = 1
}

public enum OperationType
{
    Install = 0,
    Uninstall = 1,
    ShowXml = 2,
    ExtractXml = 3
}