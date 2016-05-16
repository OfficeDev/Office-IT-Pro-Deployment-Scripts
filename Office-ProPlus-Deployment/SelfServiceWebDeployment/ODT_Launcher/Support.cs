using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Deployment.Application;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Xml;
using Microsoft.Win32;

namespace ODT_Launcher
{
    public  class Support
    {

        public static List<QueryStringItem> GetQueryStringParams(string url)
        {
            if (string.IsNullOrEmpty(url)) return new List<QueryStringItem>();

            if (url.Contains("#")) url = url.Split('#')[1];
            if (url.Contains("?")) url = url.Split('?')[1];

            var qscoll = HttpUtility.ParseQueryString(url);

            return qscoll.AllKeys.Select(s => new QueryStringItem()
            {
                Name = s,
                Value = qscoll[s]
            }).ToList();
        }

        public static List<QueryStringItem> GetArguments(string[] args)
        {
            var returnList = new List<QueryStringItem>();
            if (args == null) return returnList;

            foreach (var arg in args)
            {
                if (!arg.Contains("=")) continue;
                var key = arg.Split('=')[0];
                var value = arg.Split('=')[1];

                returnList.Add(new QueryStringItem()
                {
                    Name = key,
                    Value = value
                });
            }

            return returnList;
        }

        public static void FileDownloader(string remoteFile, string localFile)
        {
            using (var client = new WebClient())
            {
                client.DownloadFile(remoteFile, localFile);
            }
        }

        public static void CleanUp(string installDir)
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

        public static List<ExecutingScenario> GetRunningScenarioTasks()
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

        public static string GetExecutingScenario()
        {
            var mainRegPath = GetOfficeCtrRegPath();
            if (mainRegPath == null) return null;
            var configKey = Registry.LocalMachine.OpenSubKey(mainRegPath);
            if (configKey == null) return null;
            var execScenario = configKey.GetValue("ExecutingScenario");
            return execScenario != null ? execScenario.ToString() : null;
        }

        public static string GetOfficeCtrRegPath()
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

        public static string GetOdtErrorMessage(string loggingPath)
        {
            var dirInfo = new DirectoryInfo(loggingPath);
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

                    if (Directory.Exists(loggingPath))
                    {
                        Directory.Delete(loggingPath);
                    }
                }
                catch { }
            }
            return null;
        }

        public static void SetLoggingPath(string xmlFilePath, string loggingPath)
        {
            if (Directory.Exists(loggingPath))
            {
                try
                {
                    Directory.Delete(loggingPath);
                }
                catch { }
            }
            Directory.CreateDirectory(loggingPath);

            var xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFilePath);

            var loggingNode = xmlDoc.SelectSingleNode("/Configuration/Logging");
            if (loggingNode == null)
            {
                XmlElement newLoggingNode = xmlDoc.CreateElement("Logging");
                newLoggingNode.SetAttribute("Level", "Standard");
                newLoggingNode.SetAttribute("Path", "");
                xmlDoc.DocumentElement.AppendChild(newLoggingNode);
                SetAttribute(xmlDoc, newLoggingNode, "Path", loggingPath);

            }
            else
            {
                SetAttribute(xmlDoc, loggingNode, "Path", loggingPath);

            }

            xmlDoc.Save(xmlFilePath);
            System.Diagnostics.Debug.WriteLine("Sdaf");
        }

        private static void SetAttribute(XmlDocument xmldoc, XmlNode xmlNode, string name, string value)
        {
            var pathAttr = xmlNode.Attributes[name];
            if (pathAttr == null)
            {
                pathAttr = xmldoc.CreateAttribute(name);
                xmlNode.Attributes.Append(pathAttr);
            }
            pathAttr.Value = value;
        }

        public static CurrentOperation GetCurrentOperation(List<ExecutingScenario> executingTasks)
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

        public static void WaitForOfficeCtrUpadate(bool showStatus = false)
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
        
        public static void MinimizeWindow()
        {
            IntPtr winHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            ShowWindow(winHandle, SW_SHOWMINIMIZED);
        }

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        public const int SW_SHOWMINIMIZED = 2;


    }

    public class QueryStringItem
    {
        public string Name { get; set; }

        public string Value { get; set; }
    }
}
