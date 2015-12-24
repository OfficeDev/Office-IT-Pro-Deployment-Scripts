//css_ref ..\..\WixSharp.dll;
//css_ref System.Core.dll;
//css_ref ..\..\Wix_bin\SDK\Microsoft.Deployment.WindowsInstaller.dll;

using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.OfficeProPlus.InstallGenerator;
using Microsoft.Win32;
using RegistryReader;
using WixSharp;
using System;
using File = WixSharp.File;

public class MsiGenerator 
{

    public MsiGeneratorReturn Generate(MsiGeneratorProperties installProperties)
    {
        var project = new ManagedProject(installProperties.Name)
        {
            UI = WUI.WixUI_ProgressOnly,
            Actions = new WixSharp.Action[]
            {
                new SetPropertyAction("InstallDirectory", installProperties.ProgramFilesPath),
                new ElevatedManagedAction("InstallOffice", Return.check, When.After, Step.InstallFiles, Condition.NOT_Installed), 
                new ElevatedManagedAction("UninstallOffice", Return.check, When.Before, Step.RemoveFiles, Condition.BeingRemoved), 
            },
            Properties = new[] 
            { 
                new Property("InstallDirectory", "empty"),
            }
        };

        var files = new List<WixSharp.File>();
        foreach (var filePath in installProperties.ProgramFiles)
        {
            files.Add(new WixSharp.File(filePath));
        }

        files.Add(new WixSharp.File(installProperties.ExecutablePath));

        project.Dirs = new[]
        {
            new Dir(installProperties.ProgramFilesPath, files.ToArray())
        };

        project.GUID = installProperties.ProductId;
        project.ControlPanelInfo = new ProductInfo() {Manufacturer = "Microsoft Corporation"};
        project.OutFileName = installProperties.MsiPath;

        project.Load += project_Load;
        project.AfterInstall += project_AfterInstall;

        if (!string.IsNullOrEmpty(installProperties.WixToolsPath))
        {
            Compiler.WixLocation = installProperties.WixToolsPath + @"\";
            Compiler.WixSdkLocation = installProperties.WixToolsPath + @"\sdk\";
        }
        else
        {
            Compiler.WixLocation = @"wixTools\";
            Compiler.WixSdkLocation = @"wixTools\sdk\";
        }

        Compiler.BuildMsi(project);

        var installDirectory = new MsiGeneratorReturn
        {
            GeneratedFilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        };

        return installDirectory;
    }

    private void project_Load(SetupEventArgs e)
    {
        if (e.IsUISupressed)
        {
            
        }
    }

    private void project_AfterInstall(SetupEventArgs e)
    {
        var errorMessage = GetOdtErrorMessage();
        if (e.IsInstalling)
        {
            if (errorMessage != null)
            {
                //MessageBox.Show(errorMessage);
                e.Result = ActionResult.Failure;
                return;
            }
            else
            {
                e.Result = ActionResult.Success;  
            }
        }
        else if (e.IsRepairing)
        {
            RepairOffice(e);
        }
        else if (e.IsUninstalling)
        {
            VerifyOfficeUninstalled(e);
        }
        else
        {
            e.Result = ActionResult.Success;
        }
    }

    public string GetOdtErrorMessage()
    {
        var tempPath = Environment.ExpandEnvironmentVariables("%temp%");
        const string logFolderName = "OfficeProPlusLogs";
        var loggingPath = tempPath + @"\" + logFolderName;

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
        catch (Exception ex)
        {

        }
        finally
        {
            try
            {
                if (Directory.Exists(loggingPath))
                {
                    Directory.Delete(loggingPath, true);
                }
            }
            catch { }
        }
        return null;
    }


    public void VerifyOfficeUninstalled(SetupEventArgs e)
    {
        string officePath = null;
        const string regPath = @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration";
        try
        {
            var officeRegKey = Registry.LocalMachine.OpenSubKey(regPath);
            if (officeRegKey != null)
            {
                officePath = officeRegKey.GetValue("ClientFolder").ToString();
            }
            else
            {
                officePath = RegistryWOW6432.GetRegKey64(RegHive.HKEY_LOCAL_MACHINE, regPath, "ClientFolder") ??
                             RegistryWOW6432.GetRegKey32(RegHive.HKEY_LOCAL_MACHINE, regPath, "ClientFolder");
            }
        }
        catch { }

        if (!string.IsNullOrEmpty(officePath))
        {
            e.Result = ActionResult.Failure;
            return;
        }

        e.Result = ActionResult.Success;
    }

    public void RepairOffice(SetupEventArgs e)
    {
        string officePath = null;

        const string regPath = @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration";

        var officeRegKey = Registry.LocalMachine.OpenSubKey(regPath);
        if (officeRegKey != null)
        {
            officePath = officeRegKey.GetValue("ClientFolder").ToString();
        }
        else
        {
            officePath = RegistryWOW6432.GetRegKey64(RegHive.HKEY_LOCAL_MACHINE, regPath, "ClientFolder") ??
                         RegistryWOW6432.GetRegKey32(RegHive.HKEY_LOCAL_MACHINE, regPath, "ClientFolder");
        }

        if (officePath == null)
        {
            e.Result = ActionResult.Success;
            return;
        }

        var officeFilePath = officePath + @"\OfficeClickToRun.exe";

        if (!System.IO.File.Exists(officeFilePath))
        {
            e.Result = ActionResult.Success;
            return;
        }

        var p = new Process
        {
            StartInfo = new ProcessStartInfo()
            {
                FileName = officeFilePath,
                Arguments = "scenario=Repair DisplayLevel=True",
                CreateNoWindow = true,
                UseShellExecute = false
            },
        };
        p.Start();
        p.WaitForExit();

        e.Result = ActionResult.Success;
    }

}

public class CustomActions
{
    [CustomAction]
    public static ActionResult InstallOffice(Session session)
    {
        try
        {
            var installDir = session.CustomActionData["INSTALLDIR"];
            if (installDir == null) return ActionResult.Failure;

            var isSilent = false;

            try
            {
                var uiLevel = session.CustomActionData["UILevel"];
                if (uiLevel == "2" || uiLevel == "3")
                {
                    isSilent = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            var p = new Process
            {
                StartInfo = new ProcessStartInfo()
                {
                    FileName = installDir + @"\InstallOfficeProPlus.exe",
                    CreateNoWindow = true,
                    UseShellExecute = false
                },
            };

            if (isSilent)
            {
                p.StartInfo.Arguments = "/silent";
            }

            p.Start();
            p.WaitForExit();

            return ActionResult.Success;
        }
        catch (Exception ex)
        {
            return ActionResult.Failure;
        }
    }

    [CustomAction]
    public static ActionResult UninstallOffice(Session session)
    {
        try
        {
            var installDir = session.CustomActionData["INSTALLDIR"];
            if (installDir == null) return ActionResult.Failure;

            var p = new Process
            {
                StartInfo = new ProcessStartInfo()
                {
                    FileName = installDir + @"\InstallOfficeProPlus.exe",
                    Arguments = "/uninstall",
                    CreateNoWindow = true,
                    UseShellExecute = false
                },
            };
            p.Start();
            p.WaitForExit();

            return ActionResult.Success;
        }
        catch (Exception ex)
        {
            return ActionResult.Failure;
        }
    }


}
