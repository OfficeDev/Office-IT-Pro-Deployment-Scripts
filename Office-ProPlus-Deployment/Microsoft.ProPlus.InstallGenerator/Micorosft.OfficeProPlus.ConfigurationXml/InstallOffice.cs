using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.IO;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using Microsoft.Win32;
//[assembly: AssemblyTitle("")]
//[assembly: AssemblyProduct("")]
//[assembly: AssemblyDescription("")]
//[assembly: AssemblyVersion("")]
//[assembly: AssemblyFileVersion("")]

class InstallOffice
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
            var currentDirectory = Environment.ExpandEnvironmentVariables("%temp%");
            installDir = currentDirectory + @"\Office365ProPlus";
            Directory.CreateDirectory(installDir);

            var filesXml = GetTextFileContents("files.xml");
            _xmlDoc = new XmlDocument();
            _xmlDoc.LoadXml(filesXml);

            Console.Write("Extracting Install Files...");
            fileNames = GetEmbeddedItems(installDir);
            Console.WriteLine("Done");

            var odtFilePath = installDir + @"\" + fileNames.FirstOrDefault(f => f.ToLower().EndsWith(".exe"));
            var xmlFilePath = installDir + @"\" + fileNames.FirstOrDefault(f => f.ToLower().EndsWith(".xml"));

            if (!File.Exists(odtFilePath)) { throw (new Exception("Cannot find ODT Executable")); }
            if (!File.Exists(xmlFilePath)) { throw (new Exception("Cannot find Configuration Xml file")); }

            Console.WriteLine("Installing Office 365 ProPlus...");
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
        }
        finally
        {
            CleanUp(installDir);
        }
    }

    public string GetTextFileContents(string fileName)
    {
        var resourceName = fileName;
        var resourceNames = Assembly.GetExecutingAssembly().GetManifestResourceNames();
        foreach (var name in resourceNames)
        {
            if (name.ToLower().EndsWith(fileName.ToLower()))
            {
                resourceName = name;
            }
        }

        var strReturn = "";
        using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
        using (var reader = new StreamReader(stream))
        {
            strReturn = reader.ReadToEnd();
        }
        return strReturn;
    }

    public List<string> GetEmbeddedItems(string targetDirectory)
    {
        var returnFiles = new List<string>();
        var assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase);
        if (assemblyPath == null) return returnFiles;

        //var appRoot = new Uri(assemblyPath).LocalPath;
        var assembly = Assembly.GetExecutingAssembly();
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
