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
using System.Runtime.CompilerServices;
using WindowsInstaller;
using Microsoft.OfficeProPlus.MSIGen;
using WixSharp.CommonTasks;
using File = WixSharp.File;

public class MsiGenerator 
{
    public string OriginalMsiPath { get; set; }
    public MsiGeneratorReturn Generate(MsiGeneratorProperties installProperties)
    {
        var project = new ManagedProject(installProperties.Name)
        {
            UI = WUI.WixUI_ProgressOnly,
            Actions = new WixSharp.Action[]
            {
                new SetPropertyAction("InstallDirectory", installProperties.ProgramFilesPath),
                new ElevatedManagedAction(CustomActions.InstallOffice, Return.check, When.After, Step.InstallFiles, Condition.NOT_Installed),
                new ElevatedManagedAction(CustomActions.UninstallOffice, Return.check, When.Before, Step.RemoveFiles, Condition.BeingRemoved),
            },
            Properties = new[] 
            { 
                new Property("InstallDirectory", "empty"),
                new Property()
                {
                    Name = "ProductGuid",
                    Value = installProperties.ProductId.ToString()
                }
                
            }
        };

        project.Media.AttributesDefinition+= ";CompressionLevel=high";


        var files = new List<WixSharp.File>();
        foreach (var filePath in installProperties.ProgramFiles)
        {
            files.Add(new WixSharp.File(filePath));
        }
        files.Add(new WixSharp.File(installProperties.ExecutablePath));


        var rootDir = new Dir(installProperties.ProgramFilesPath, files.ToArray());
        project.Dirs = new[]
        {
            rootDir
        };

        project.GUID = installProperties.ProductId;
        project.ControlPanelInfo = new ProductInfo()
        {
            Manufacturer = installProperties.Manufacturer,
            Comments = installProperties.ProductId.ToString()
        };
        project.OutFileName = installProperties.MsiPath;
        project.UpgradeCode = installProperties.UpgradeCode;
        project.Version = installProperties.Version;
        project.MajorUpgrade = new MajorUpgrade()
        {
            DowngradeErrorMessage = "A later version of [ProductName] is already installed. Setup will now exit.",
            AllowDowngrades = false,
            AllowSameVersionUpgrades = false
        };

        //project.Platform = Platform.x64;

        //project.MajorUpgradeStrategy.RemoveExistingProductAfter = null;

        project.Load += project_Load;
        project.AfterInstall += project_AfterInstall;
        //project.InstallScope = InstallScope.perMachine;
        

        if (!string.IsNullOrEmpty(installProperties.Language))
        {
            project.Language = installProperties.Language;
        }
        
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

        var returnValue  = Compiler.BuildMsi(project);

        var installDirectory = new MsiGeneratorReturn
        {
            GeneratedFilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        };

        return installDirectory;
    }

    private Dir AddProgramFiles(string rootPath, Dir parentDir, MsiDirectory directory)
    {
        var path = rootPath + @"\" + directory.RelativePath;

        var newDir = new Dir()
        {
            Name = directory.Name
        };

        var fileLength = parentDir.Files.Length;
        var newFileList = new File[fileLength + directory.MsiFiles.Count];
        parentDir.Files.CopyTo(newFileList, 0);

        var startIndex = parentDir.Files.Length;
        foreach (var file in directory.MsiFiles)
        {
            newFileList[startIndex] = new File(file.Path);
            startIndex ++;
        }

        newDir.Files = newFileList;

        foreach (var subDir in directory.MsiDirectories)
        {
            AddProgramFiles(rootPath, newDir, subDir);
        }

        var length = parentDir.Dirs.Length;
        var newList = new Dir[length + 1];
        parentDir.Dirs.CopyTo(newList, 0);
        newList[length] = newDir;
        parentDir.Dirs = newList;

        return newDir; 
    }

    private void project_Load(SetupEventArgs e)
    {
        if (Directory.Exists(@"C:\Windows\Temp\OfficeProPlus"))
            Directory.Delete(@"C:\Windows\Temp\OfficeProPlus", true);
        string launchLocation = e.MsiFile;
        string officeFolder = "";
        foreach (var currentDirectory in Directory.GetDirectories(launchLocation.Substring(0, launchLocation.LastIndexOf(@"\"))))
        {
            if (currentDirectory.ToLower().EndsWith("office"))
            {
                officeFolder = currentDirectory;
            }
        }
        if (!string.IsNullOrEmpty(officeFolder))
        {
            //copy files to install location
            CopyFolder(new DirectoryInfo(officeFolder), new DirectoryInfo(@"C:\Windows\Temp\OfficeProPlus\Office"));
        }

        if (e.IsUISupressed)
        {
            
        }
    }


    public static void CopyFolder(DirectoryInfo source, DirectoryInfo target)
    {
        foreach (DirectoryInfo dir in source.GetDirectories())
            CopyFolder(dir, target.CreateSubdirectory(dir.Name));
        foreach (FileInfo file in source.GetFiles())
            file.CopyTo(Path.Combine(target.FullName, file.Name), true);


    }

    private void project_AfterInstall(SetupEventArgs e)
    {
        var errorMessage = GetOdtErrorMessage();
        if (e.IsInstalling)
        {
            if (errorMessage != null)
            {
                e.Result = ActionResult.Success;
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
            //VerifyOfficeUninstalled(e);
            e.Result = ActionResult.Success;
        }
        else
        {
            e.Result = ActionResult.Success;
        }       
    }

    public string GetOdtErrorMessage()
    {
        var tempPath = Environment.ExpandEnvironmentVariables("%public%");
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
            if (installDir == null)
            {
                MessageBox.Show("No Install Directory");
                return ActionResult.Failure;
            }

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

            var arguments = "/uninstall";

            if (isSilent)
            {
                arguments += " /silent";
            }

            var p = new Process
            {
                StartInfo = new ProcessStartInfo()
                {
                    FileName = installDir + @"\InstallOfficeProPlus.exe",
                    Arguments = arguments,
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
