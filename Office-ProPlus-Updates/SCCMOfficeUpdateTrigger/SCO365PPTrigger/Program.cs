//*********************************************************
// THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
// IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
// PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
//*********************************************************
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;

namespace Contoso.SCO365PPTrigger
{
    /// <summary>
    /// This application is designed to be run remotely on a users Office 365 Pro Plus PC, via a SSCM package containing the source media for
    /// an Office 365 Pro Plus build. This ensures that the PC updates Office 365 Pro Plus from that SCCM package on the DP closest to it.
    /// </summary>
    class Program
    {

        const string C2RClientExe = "officec2rclient.exe";
        const string C2RClientArg = "/update user";
        const string keyPath16 = @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"; //path to registy on Office 16 (2016)
        const string keyPath15 = @"SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"; //path to registy on Office 15 (2013)
        const string UpdatePathKey = "UpdateUrl"; //reg key which contains URL to get Office updates from
        //const string VersionKey = "VersionToReport"; //may use this to deteremine version of Office installed in the future
        const string OfficeC2RClientPathKey = "ClientFolder"; //folder path containts the officec2rclient.exe file. Different on 2013 and 2016.

        //params from command line
        static bool _enableLogging = false;
        static string _officec2rClientArgs = string.Empty; //hold additional officec2rclient.exe params supplied by SCCM package config

        //needed for log file path
        static string _logFilePath = string.Empty; //will be used if logging is enabled

        //return int is used to have an exit code, which SCCM can then report on, so you can see why process failed        
        static int Main(string[] args)
        {

            if (SetupArguments(args))
            {
                //get version of office i.e. 15 or 16
                int officeVersion = GetOfficeVersion();
                LogMessage("Found Office version " + officeVersion);

                if (officeVersion == -1)
                {
                    LogMessage("No office 15 or 16 365 Pro Plus found in registry or on PC, exiting process");
                    return 1002; //exit with no office key found
                }

                string newDPPath = string.Empty;


                //get the path from where the exe is current running i.e. closest DP share
                newDPPath = GetDPPath();
                LogMessage("DP execution path is " + newDPPath);


                if (newDPPath.Length == 0)
                {
                    LogMessage("Exiting update process, failed to get DP path");
                    return 1005; //exit with failure to get DP path
                }

                //get what is currently in the registry for updatepath
                string currentUpdatePath = GetOfficeUpdateUrl(officeVersion);
                LogMessage("UpdateURL in reg is " + currentUpdatePath);

                //if the new and the reg are not the same, we must update registry
                if (newDPPath.CompareTo(currentUpdatePath) != 0)
                {

                    LogMessage("New UpdateUrl path and current registry update path are not same, need to update registry");
                    bool updateSuccess = SetUpdateUrl(officeVersion, newDPPath);
                    if (!updateSuccess)
                    {
                        LogMessage("Exiting update process, failed to update registry with new DP path");
                        return 1003; //exit with failure to update registry key
                    }

                    //check if new value was set correct ie.. read it again from reg and see if the same
                    string updatedPath = GetOfficeUpdateUrl(officeVersion);
                    if (updatedPath.CompareTo(newDPPath) != 0)
                    {
                        LogMessage("Exiting update process, regstry value is not the same as DP path");
                        return 1004; //exit with failure to update registry key
                    }

                }

                //start process             
                bool startSuccess = LaunchUpdateProcess(officeVersion);
                if (!startSuccess)
                {
                    LogMessage("Finished Failure, could not launch officec2rclient.exe");
                    return 1006; //cannot launch process exit code
                }

                LogMessage("Finished Success");
                return 0; //success exit code

            }
            else
                return 1001; //error setting up args

        }

        /// <summary>
        /// Look for params passed on command line. If /q show help section, otherwise setup environment ready for execution
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        static bool SetupArguments(string[] args)
        {

            try
            {
                //show command line params if pass in /?
                if (args.Length == 1)
                    if (args[0].Trim().CompareTo(@"/?") == 0) //show params for usage
                    {
                        ShowCommandLineParams();
                        return false;//exit
                    }

                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].Equals("-EnableLogging", StringComparison.InvariantCultureIgnoreCase) && (i + 1 < args.Length) && !args[i + 1].StartsWith("-"))
                        _enableLogging = Convert.ToBoolean(args[i + 1]); //could blow up if not convertable, make sure not to handle exception
                    else if (args[i].Equals("-C2RArgs", StringComparison.InvariantCultureIgnoreCase) && (i + 1 < args.Length) && !args[i + 1].StartsWith("-"))
                        _officec2rClientArgs = args[i + 1];

                }

                if (_enableLogging)
                {
                    _logFilePath = Path.Combine(Path.GetTempPath(), "sccmofficeupdater_" + DateTime.Now.ToString("ddMMyyHHmmss") + ".log"); //will blow up with Environment var is NULL
                    LogMessage("Log file generated by custom executable to trigger Office 365 Pro Plus to update, via SCCM DP");
                    LogMessage("Log file path is set to " + _logFilePath);
                    // Path.GetTempPath
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Show help for command line
        /// </summary>
        static void ShowCommandLineParams()
        {
            Console.WriteLine("Custom tool to ensure Office 365 Pro Plus updates from package on SCCM DP. Does this by ");
            Console.WriteLine("editing PC registry with updateurl to SCCM DP location, and then calls officec2rclient.exe.");
            Console.WriteLine("Note: Requires .Net Framework 3.5 or higher, and to be run as admin on user's PC.");
            Console.WriteLine();
            Console.WriteLine("Examples for usage");
            Console.WriteLine("------------------");
            Console.WriteLine();
            Console.WriteLine("SCO365PPTrigger.exe");
            Console.WriteLine("SCO365PPTrigger.exe -EnableLogging true");
            Console.WriteLine("SCO365PPTrigger.exe -EnableLogging true -C2RArgs \"officec2rclient.exe params\"");
            Console.WriteLine();
            Console.WriteLine("Parameters Help");
            Console.WriteLine("---------------");
            Console.WriteLine();
            Console.WriteLine("SCO365PPTrigger.exe run without any params, will update Ofice 365 Pro Plus UpdateUrl registy path ");
            Console.WriteLine(" with the SCCM DP package path and then calls \"officec2rclient.exe /update user\". ");
            Console.WriteLine(" There are 2 optional parameters shown below: ");
            Console.WriteLine();
            Console.WriteLine("-EnableLogging       Set to true or false. If true will create a log file in the users system temp ");
            Console.WriteLine("                     folder. Log file name looks like \"sccmupdater_*.log\". Default value is false");
            Console.WriteLine("-C2RArgs             Add params from officec2rclient.exe, enclose in quotes. Allows you ");
            Console.WriteLine("                     to control update experience e.g. \"updatetoversion=ver displaylevel=false\".");
            Console.WriteLine("                     Use the officec2rclient.exe documentation to see full list of supported params.");
            Console.WriteLine();

        }


        /// <summary>
        /// Get Office 365 Pro Plus updateurl from registry
        /// </summary>
        /// <param name="OfficeVersion">15 or 16 supported</param>
        /// <returns></returns>
        static string GetOfficeUpdateUrl(int OfficeVersion)
        {

            try
            {
                string regPath = string.Empty;
                if (OfficeVersion == 16)
                    regPath = keyPath16;
                else
                    regPath = keyPath15;

                string updateUrl = RegistryHelpers.GetRegistryValue(regPath, UpdatePathKey) as string;
                return updateUrl.ToLower();
            }
            catch (Exception ex)
            {
                LogMessage("Error getting UpdateURL from registry");
                LogMessage(string.Format("Exception occured:{0}, {1}", ex.Message, ex.StackTrace.ToString()));
                return string.Empty;
            }
        }


        /// <summary>
        ///  Returns executing path as lower case string
        /// </summary>
        /// <returns></returns>
        static string GetDPPath()
        {
            try
            {


                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                return Path.GetDirectoryName(exePath).ToLower();

            }
            catch (Exception ex)
            {
                LogMessage("Error assembling DP path from running executable");
                LogMessage(string.Format("Exception occured:{0}, {1}", ex.Message, ex.StackTrace.ToString()));
                return string.Empty;
            }
        }


        /// <summary>
        /// Determine which Office version is installed on users PC
        /// </summary>
        /// <returns>15, or 16, or -1 if no Office / failure </returns>
        static int GetOfficeVersion()
        {

            try
            {
                //must be better way of doing this
                string c2rclientPath = RegistryHelpers.GetRegistryValue(keyPath16, OfficeC2RClientPathKey) as string;
                if (c2rclientPath != null)
                {
                    if (File.Exists(Path.Combine(c2rclientPath, C2RClientExe)))
                    {
                        return 16;
                    }
                }
                else
                    LogMessage("Couldnt open registy key for Office 16 " + keyPath16 + " value " + OfficeC2RClientPathKey);

                c2rclientPath = RegistryHelpers.GetRegistryValue(keyPath15, OfficeC2RClientPathKey) as string;
                if (c2rclientPath != null)
                {
                    if (File.Exists(Path.Combine(c2rclientPath, C2RClientExe)))
                    {
                        return 15;
                    }
                }
                else
                    LogMessage("Couldnt open registy key for Office 15 " + keyPath15 + " value " + OfficeC2RClientPathKey);

                return -1;
            }
            catch (Exception ex)
            {
                LogMessage("Error trying to open registry keys to identify Office version installed");
                LogMessage(string.Format("Exception occured:{0}, {1}", ex.Message, ex.StackTrace.ToString()));
                return -1;
            }
        }


        /// <summary>
        /// Update registry with new UpdateUrl path, Office version required as paths are different
        /// </summary>
        /// <param name="OfficeVersion"></param>
        /// <param name="UpdatePath"></param>
        /// <returns></returns>
        static bool SetUpdateUrl(int OfficeVersion, string UpdatePath)
        {
            string regPath = string.Empty;
            if (OfficeVersion == 16)
                regPath = keyPath16;
            else
                regPath = keyPath15;

            try
            {
                RegistryHelpers.SetRegistryValue(regPath, UpdatePathKey, UpdatePath);
                LogMessage("Registry update success at " + regPath + "\\" + UpdatePathKey);
                return true;
            }
            catch (Exception ex)
            {
                LogMessage("Failed to update " + regPath + "\\" + UpdatePathKey);
                LogMessage(string.Format("Exception occured:{0}, {1}", ex.Message, ex.StackTrace.ToString()));
                return false;
            }
        }

        /// <summary>
        /// Start officec2rclient.exe with supplied params
        /// </summary>
        /// <param name="OfficeVersion"></param>
        /// <returns></returns>
        static bool LaunchUpdateProcess(int OfficeVersion)
        {
            string _officeConfigPath = string.Empty;

            if (OfficeVersion == 16)
                _officeConfigPath = keyPath16;
            else
                _officeConfigPath = keyPath15;

            string exePath = RegistryHelpers.GetRegistryValue(_officeConfigPath, OfficeC2RClientPathKey) as string;


            if (File.Exists(Path.Combine(exePath, C2RClientExe)))
            {
                // Launch it
                Process p;
                try
                {
                    p = new Process();
                    p.StartInfo.WorkingDirectory = exePath;
                    p.StartInfo.FileName = C2RClientExe;
                    p.StartInfo.Arguments = C2RClientArg + " " + _officec2rClientArgs; //do we need to cater for quotes
                    LogMessage("Starting " + p.StartInfo.FileName + " with params " + p.StartInfo.Arguments);
                    p.Start();
                    return true;
                }
                catch (Exception ex)
                {
                    LogMessage("Failed to launch C2RClientExe");
                    LogMessage(string.Format("Exception occured:{0}, {1}", ex.Message, ex.StackTrace.ToString()));
                    return false;
                }
            }
            else
            {
                LogMessage(string.Format("OfficeC2RClient.exe doesn't exist at {0}", exePath));
                return false;
            }

        }


        /// <summary>
        /// Log message via command line, optionaly will create log file
        /// </summary>
        /// <param name="Message"></param>
        static void LogMessage(string Message)
        {

            try
            {

                Console.WriteLine(Message);
                //if logging is enabled and log file path exists File.
                if (_enableLogging && _logFilePath.Length > 0)
                {
                    File.AppendAllText(_logFilePath, Environment.NewLine + DateTime.Now + " : " + Message);
                }


            }
            catch (Exception ex)
            {

                Console.WriteLine("Error writing to log file. " + ex.Message);

            }
        }



    }

    public class RegistryHelpers
    {

        public static RegistryKey GetRegistryKey()
        {
            return GetRegistryKey(null);
        }

        public static RegistryKey GetRegistryKey(string keyPath)
        {

            //returns null if it cannot find key
            return Registry.LocalMachine.OpenSubKey(keyPath, true);

        }

        public static object GetRegistryValue(string keyPath, string keyName)
        {

            RegistryKey registry = GetRegistryKey(keyPath);

            if (registry == null)
                return null;
            else
            {
                object tmpKeyVal = registry.GetValue(keyName);
                registry.Close();
                return tmpKeyVal;
            }
        }

        public static void SetRegistryValue(string keyPath, string keyName, object keyValue)
        {
            RegistryKey registry = GetRegistryKey(keyPath);
            registry.SetValue(keyName, keyValue);
            registry.Close();

        }
    }

}

