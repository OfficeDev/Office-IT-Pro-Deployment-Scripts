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
using System.Management;

//[assembly: AssemblyTitle("")]
//[assembly: AssemblyProduct("")]
//[assembly: AssemblyDescription("")]
//[assembly: AssemblyVersion("")]
//[assembly: AssemblyFileVersion("")]

public class InstallOfficeWmi
{
    public string remoteUser { get; set; }
    public uint ProcessId { get; set; }
    public bool EventArrived { get; set;}
    public string remoteComputerName { get; set; }
    public string remoteDomain { get; set; }
    public string remotePass { get; set; }
    public string newVersion { get; set; }
    public string newChannel { get; set; }
    public string connectionNamespace { get; set; }
    public ManagementScope scope { get; set; }
    public ManagementScope scope2 { get; set; }



    public async Task ChangeOfficeChannel(string targetChannel, string baseUrl)
    {
        var saveBaseUrl = "";
        try
        {

            await initConnection();
            saveBaseUrl = GetBaseCdnUrl();

            await Task.Run(() => { ChangeUpdateSource(baseUrl); });
            await Task.Run(() => { ChangeBaseCdnUrl(baseUrl); });

            await Task.Run(() => { RestartC2RSerivce(); });

            await RunOfficeUpdateAsync(targetChannel);
        }
        catch (Exception ex)
        {
            if (!string.IsNullOrEmpty(saveBaseUrl))
            {
                ChangeBaseCdnUrl(saveBaseUrl);
            }
            throw(new Exception(ex.Message));
        }
        finally
        {
            ResetUpdateSource();
        }

    }

    public async Task initConnection()
    {

        var timeOut = new TimeSpan(0, 5, 0);
        ConnectionOptions options = new ConnectionOptions();
        options.Authority = "NTLMDOMAIN:" + remoteDomain.Trim();
        options.Username = remoteUser.Trim();
        options.Password = remotePass.Trim();
        options.Impersonation = ImpersonationLevel.Impersonate;
        options.Timeout = timeOut;



        scope = new ManagementScope("\\\\" + remoteComputerName.Trim() + connectionNamespace, options);
        scope.Options.EnablePrivileges = true;

        scope2 = new ManagementScope("\\\\" + remoteComputerName.Trim() + "\\root\\default", options);
        scope2.Options.EnablePrivileges = true;

        try
        {
            await Task.Run(() => { scope.Connect(); });
            //await Task.Run(() => { scope2.Connect(); });
        }

        catch (Exception)
        {
            await Task.Run(() => { scope.Connect(); });
            //await Task.Run(() => { scope2.Connect(); });
        }

    }


    private void scopeDisconnect(ManagementScope currentScope)
    {
        currentScope.Connect();
    }


    public  void RestartC2RSerivce()
    {

        try
        {

            CreateRegistryValue(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "LocalAccountTokenFilterPolicy", "1");
            SetRegistryValue(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "LocalAccountTokenFilterPolicy", "SetDWORDValue", "1");

            SelectQuery query = new SelectQuery("select * from Win32_Service where name='ClickToRunSvc'");

            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query))
            {

                ManagementObjectCollection collection = searcher.Get();

                foreach (ManagementObject service in collection)

                {

                    if (service["Started"].Equals(true))

                    {

                        var outparams = service.InvokeMethod("StopService", null);
                        Thread.Sleep(2000);
                        service.InvokeMethod("StartService", null);

                    }

                }

            }


        }
        catch (Exception ex)
        {
           throw (new Exception("Cannot restart ClickToRunSvc. "+ex.Message));

        }

    }


    #region Office Operations

    public string GetOfficeC2RPath()
    {
        var configKey = "";

        try
        {
            var officeRegKey = GetOfficeCtrRegPath().Result;

            configKey = GetRegistryValue(officeRegKey + "\\Configuration", "ClientFolder").Result;
        }
        catch(Exception ex)
        {
            throw (new Exception("Cannot find c2r path. " + ex.Message));
        }
        

        return configKey;
    }
    public async Task RunOfficeUpdateAsync(string version)
    {
        await Task.Run(async () => {
            try
            {
                var c2RPath = GetOfficeC2RPath() + @"\OfficeC2RClient.exe /update user displaylevel=true forceappshutdown=true updatepromptuser=false updatetoversion=" + version;
                var mainRegKey = GetOfficeCtrRegPath().Result;
                var c2rExe = new[] { c2RPath };
                var wmiProcess = new ManagementClass(scope, new ManagementPath("Win32_Process"), new ObjectGetOptions());
                ManagementBaseObject inParams = wmiProcess.GetMethodParameters("Create");
                inParams["CommandLine"] = c2RPath;

                await Task.Run(() => { return wmiProcess.InvokeMethod("Create", inParams, null); });

                Thread.Sleep(1000);

                var executingScenario = GetRegistryValue(mainRegKey, "ExecutingScenario").Result;

                while(executingScenario != null)
                {
                    Thread.Sleep(1000);
                    executingScenario = GetRegistryValue(mainRegKey, "ExecutingScenario").Result;
                }


                var updateStatus = GetRegistryValue(mainRegKey, "LastScenarioResult").Result;

                if(updateStatus != "Success")
                {
                    throw (new Exception("Channel/version change was not successful"));
                }

            }
            catch(Exception ex)
            {
                throw (new Exception(ex.Message));
            }
        });
    }


    #endregion

    #region Update Monitoring
    public void ProcessStoppedEventArrived(object sender, EventArrivedEventArgs e)
    {
        if ((uint)e.NewEvent.Properties["ProcessId"].Value == ProcessId)
        {
            var exitCode = (int)(uint)e.NewEvent.Properties["ExitStatus"].Value;

        }
    }


    public void ResetUpdateSource()
    {
        const string policyPath = @"SOFTWARE\Policies\Microsoft\office\16.0\common\";
        var policyKey = GetRegistryBaseKey(policyPath, "officeupdate", "EnumKey");
        //var policyKey = Registry.LocalMachine.OpenSubKey(policyPath, true);
        if (policyKey != null)
        {
            var saveUpdatePath = GetRegistryValue(policyKey.ToString(), "saveupdatepath").Result;
            if (!string.IsNullOrEmpty(saveUpdatePath.ToString()))
            {
                SetRegistryValue(policyPath + "\\officeupdate", "updatePath", "SetStringValue", saveUpdatePath);
                SetRegistryValue(policyPath + "\\officeupdate", "saveupdatepath","DeleteValue", null);
                //policyKey.SetValue("updatepath", saveUpdatePath, RegistryValueKind.String);
                //policyKey.DeleteValue("saveupdatepath");
            }
        }

        var mainRegKey = GetOfficeCtrRegPath().Result;
        if (mainRegKey == null) return;

        var configKey = GetRegistryBaseKey(mainRegKey, "Configuration", "EnumKey");
        if (configKey == null) return;

        var saveUpdateUrl = GetRegistryValue(configKey.ToString(), "SaveUpdateUrl").Result;
        if (string.IsNullOrEmpty(saveUpdateUrl)) return;

        SetRegistryValue(mainRegKey + "\\Configuration", "UpdateUrl",  "SetStringValue",saveUpdateUrl);
        SetRegistryValue(mainRegKey + "\\Configuration", "SaveUpdateUrl","DeleteValue", null);
        //configKey.SetValue("UpdateUrl", saveUpdateUrl, RegistryValueKind.String);
        //configKey.DeleteValue("SaveUpdateUrl");
    }

    public string ChangeUpdateSource(string updateSource)
    {
        try
        {

       
        var currentupdatepath = "";

        const string policyPath = @"SOFTWARE\Policies\Microsoft\office\16.0\common\";
        var policyKey =  GetRegistryBaseKey(policyPath, "officeupdate", "EnumKey");
        if (policyKey != null)
        {
            currentupdatepath = GetRegistryValue(policyKey.ToString(), "updatepath").Result;
            var saveupdatePath = GetRegistryValue(policyKey.ToString(), "saveupdatepath").Result;
            if (!string.IsNullOrEmpty(currentupdatepath) && !string.IsNullOrEmpty(updateSource))
            {
                if (string.IsNullOrEmpty(saveupdatePath.ToString()))
                {
                    SetRegistryValue(policyPath + "officeUpdate", "saveupdatepath","SetStringValue", currentupdatepath);
                    //policyKey.SetValue("saveupdatepath", currentupdatepath, RegistryValueKind.String);
                }
                SetRegistryValue(policyPath + "officeUpdate", "updatepath", "SetStringValue", updateSource);
                //policyKey.SetValue("updatepath", updateSource, RegistryValueKind.String);

            }
        }

        if (!string.IsNullOrEmpty(currentupdatepath)) return currentupdatepath;

        var mainRegKey = GetOfficeCtrRegPath().Result;
        if (mainRegKey == null) return null;

        //var configKey = mainRegKey.OpenSubKey(@"Configuration", true);
        var configKey = GetRegistryBaseKey(mainRegKey, "Configuration", "EnumKey");
        if (configKey == null) return null;

        currentupdatepath = GetRegistryValue(mainRegKey + @"\Configuration", "UpdateUrl").Result;
        var saveupdateUrl = GetRegistryValue(mainRegKey + @"\Configuration", "SaveUpdateUrl").Result;
        if (string.IsNullOrEmpty(currentupdatepath) || string.IsNullOrEmpty(updateSource)) return currentupdatepath;

        if (string.IsNullOrEmpty(saveupdateUrl.ToString()))
        {

            SetRegistryValue(mainRegKey + @"\Configuration", "UpdateUrl", "SetStringValue", currentupdatepath);
            //configKey.SetValue("SaveUpdateUrl", currentupdatepath, RegistryValueKind.String);
        }

        SetRegistryValue(mainRegKey + @"\Configuration", "UpdateUrl", "SetStringValue", updateSource);
        //configKey.SetValue("UpdateUrl", updateSource, RegistryValueKind.String);

        return currentupdatepath;
        }
        catch(Exception ex)
        {
            throw (new Exception("Cannot change update source. "+ex.Message));
        }
    }

    public string GetBaseCdnUrl()
    {
      try
        {
            var mainRegKey = GetOfficeCtrRegPath().Result;
            if (mainRegKey == null) return "";


            //var configKey = GetRegistryValue(mainRegKey.ToString()+"Configuration","CDNBaseUrl");
            var configKey = GetRegistryBaseKey(mainRegKey, "Configuration", "EnumKey");
            //var configKey = mainRegKey.OpenSubKey(@"Configuration", true);
            if (configKey == null) return "";

            return GetRegistryValue(mainRegKey + "\\Configuration", "CDNBaseUrl").Result;
        }
        catch (Exception ex)
        {
            throw (new Exception("Cannot get base cdn url. " + ex.Message));
        }
       

       

    }

    public void ChangeBaseCdnUrl(string updateSource)
    {
        try
        {

    
        var mainRegKey = GetOfficeCtrRegPath().Result;
        if (mainRegKey == null) return;

        //var configKey = mainRegKey.OpenSubKey(@"Configuration", true);
        var configKey = GetRegistryBaseKey(mainRegKey, "Configuration", "EnumKey");
        if (configKey == null) return;

        var cdnBaseUrl = GetRegistryValue(mainRegKey+"\\"+configKey, "CDNBaseUrl").Result;
        //configKey.SetValue("CDNBaseUrl", updateSource, RegistryValueKind.String);
        SetRegistryValue(mainRegKey + "\\" + configKey, "CDNBaseUrl", "SetStringValue", updateSource);
        }
        catch(Exception ex)
        {
            throw (new Exception("Cannot change base cdn url. " + ex.Message));
        }
    }

    public async Task<string> GetOfficeCtrRegPath()
    {
        var path16 = @"SOFTWARE\Microsoft\Office\";
        var path15 = @"SOFTWARE\Microsoft\Office\15.0\";
      
            var office16Key = GetRegistryBaseKey(path16, "ClickToRun","EnumKey");
            var office15Key = GetRegistryBaseKey(path15, "ClickToRun","EnumKey");
            //var office16Key = Registry.LocalMachine.OpenSubKey(path16, true);
            //var office15Key = Registry.LocalMachine.OpenSubKey(path15, true);

            if (office16Key != null)
            {
                return path16+"ClickToRun";
            }
            else
            {
                if (office15Key != null)
                {
                    return path15 + "ClickToRun";
                }
            }

            //var Hklm32 = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, remoteComputerName, RegistryView.Registry32);
            //var Hklm32 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);

            //office16Key = Hklm32.OpenSubKey(path16, true);
            //office15Key = Hklm32.OpenSubKey(path15, true);
            office16Key = @"\SOFTWARE\Wow6432Node\Microsoft\Office\";
            office15Key = @"\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\";


            if (office16Key != null)
                {
                    return path16 + "ClickToRun";
                }
                else
                {
                    if (office15Key != null)
                    {
                    return path15 + "ClickToRun";
                    }
                }

                return null;
   

    }


    #endregion
    


    #region Support Functions

    private async  Task<string> GetRegistryValue(string regKey, string valueName)
    {
        string value = null;
        await Task.Run(() =>
        {
            ManagementClass registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);
            ManagementBaseObject inParams = registry.GetMethodParameters("GetStringValue");

            inParams["hDefKey"] = 0x80000002;
            inParams["sSubKeyName"] = regKey;
            inParams["sValueName"] = valueName;

            ManagementBaseObject outParams = registry.InvokeMethod("GetStringValue", inParams, null);

            try
            {
                if (outParams.Properties["sValue"].Value != null)
                {
                    value = outParams.Properties["sValue"].Value.ToString();
                }
            }
            catch (Exception)
            {
                return null;
            }
            return value;

        });

        return value;

    }

    private async void CreateRegistryValue(string regKey, string valueName, string value)
    {
      
            await Task.Run(() =>
            {
                


                    ManagementClass registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);
                    ManagementBaseObject inParams = registry.GetMethodParameters("SetDWORDValue");

                    inParams["hDefKey"] = 0x80000002;
                    inParams["sSubKeyName"] = regKey;
                    inParams["sValueName"] = valueName;

                    ManagementBaseObject outParams = registry.InvokeMethod("SetDWORDValue", inParams, null);
                

                //var result = outParams.Properties["ReturnValue"].Value.ToString();

                //if (result != "0")
                //{
                //    throw (new Exception("Cannot create registry Value: " + valueName + " at path: " + regKey + " return code: " + result));
                //}
                //}
                //catch (Exception ex)
                //{
                //}




            });

    }

    private String[] GetSubKeyNames(string parentKey)
    {
        ManagementClass registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);
        ManagementBaseObject inParams = registry.GetMethodParameters("EnumKey");

        inParams["hDefKey"] = 0x80000002;
        inParams["sSubKeyName"] = parentKey;

        ManagementBaseObject outParams = registry.InvokeMethod("EnumKey", inParams, null);

        try
        {
            var subKeyNames = (String[])outParams.Properties["sNames"].Value;

            return subKeyNames; 

        }
        catch (Exception)
        {
            return null;
        }


        return null;
    }

    private String[] GetValueNames(string keyName)
    {
        ManagementClass registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);
        ManagementBaseObject inParams = registry.GetMethodParameters("EnumKey");

        inParams["hDefKey"] = 0x80000002;
        inParams["sSubKeyName"] = keyName;

        ManagementBaseObject outParams = registry.InvokeMethod("EnumKey", inParams, null);

        try
        {
            var valueNames = (String[])outParams.Properties["sValueName"].Value;

            return valueNames;

        }
        catch (Exception)
        {
            return null;
        }


        return null;
    }

    private string GetRegistryBaseKey(string parentKey, string childKey, string getmethParam)
    {
       
            ManagementClass registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);
            ManagementBaseObject inParams = registry.GetMethodParameters(getmethParam);

            inParams["hDefKey"] = 0x80000002;
            inParams["sSubKeyName"] = parentKey;

            ManagementBaseObject outParams = registry.InvokeMethod(getmethParam, inParams, null);

            try
            {
                var subKeyNames = (String[])outParams.Properties["sNames"].Value;

                foreach (var key in subKeyNames)
                {
                    if (key == childKey)                    {
                        return key;
                    }
                }


            }
            catch (Exception)
            {
                return null;
            }
       

            return null;
    
    }

    private void SetRegistryValue(string keyPath, string valueName, string method, string keyValue = null)
    {
        try
        {

     
        ManagementClass registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);
        ManagementBaseObject inParams = registry.GetMethodParameters(method);

        if(keyValue != null)
        {
            inParams["hDefKey"] = 0x80000002;
            inParams["sSubKeyName"] = keyPath;
            inParams["sValueName"] = valueName;
            inParams["sValue"] = keyValue;
        }
        else
        {
            inParams["hDefKey"] = 0x80000002;
            inParams["sSubKeyName"] = keyPath;
            inParams["sValueName"] = valueName;
        }

        var outParams = registry.InvokeMethod(method, inParams,null);
        }
        catch (Exception)
        {

        }



    }
    #endregion


}
