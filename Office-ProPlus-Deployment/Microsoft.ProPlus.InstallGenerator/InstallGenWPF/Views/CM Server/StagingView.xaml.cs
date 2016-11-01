using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MetroDemo;
using MetroDemo.Events;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config

{
    /// <summary>
    /// Interaction logic for DeploymentStagingView.xaml
    /// </summary>
    public partial class DeploymentStagingView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;
        public event MessageEventHandler ErrorMessage;



        public DeploymentStagingView()
        {

            InitializeComponent();
        }

        private void DeploymentStagingView_OnLoaded(object sender, RoutedEventArgs e)
        {

        }

        private void StagingPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (ImgProgramCreated.IsVisible)
            {
                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = true
                });
            }
        }

        private async void DeployButton_OnClick(object sender, RoutedEventArgs e)
        {
            await StageDeployment();         
        }



        #region helpers

        private async Task StartCreatePrograms()
        {
            try
            {
                GlobalObjects.ViewModel.BlockNavigation = true;
                StageButton.IsEnabled = false;

                Dispatcher.Invoke(() =>
                {
                    WaitCreatingProgram.Visibility = Visibility.Visible;
                });

                await CreatePrograms();

                WaitCreatingProgram.Visibility = Visibility.Collapsed;
                ImgProgramCreated.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {

                WaitCreatingProgram.Visibility = Visibility.Collapsed;
                ImgProgramCreateFail.Visibility = Visibility.Visible;
                StageButton.IsEnabled = true;

                LogErrorMessage(ex); ;
            }
        }

        private async Task CreatePrograms()
        {
            await Task.Run(() =>
            {
                var n = 1;
                var CMConfig = GlobalObjects.ViewModel.CmPackage;
                var currentDirectory = Directory.GetCurrentDirectory() + @"\Scripts";

                var scriptPath = currentDirectory + $"\\Setup-CMOfficeDeployment.ps1";
                var scriptPathTmp = currentDirectory + $"\\Tmp-Setup-CMOfficeDeployment.ps1";

                foreach (var program in CMConfig.Programs)
                {


                    var channels = new List<string>();
                    var bitnesses = new List<string>();

                    var arguments =
                         $"/c Powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden . .\\Setup-CMOfficeDeployment.ps1;Create-CMOfficeDeploymentProgram -Channels ";

                    program.Channels.ForEach(c =>
                    {
                        if (!channels.Contains(c.Branch.NewName.ToString()))
                        {
                            channels.Add(c.Branch.NewName.ToString());
                        }
                    });

                    program.Bitnesses.ToList().ForEach(b =>
                    {
                        if (!bitnesses.Contains(b.Name))
                        {
                            bitnesses.Add(b.Name);
                        }
                    });

                    channels.ForEach(c =>
                    {
                        if (channels.IndexOf(c) < channels.Count - 1)
                        {
                            arguments += $"{c},";
                        }
                        else
                        {
                            arguments += c;
                        }
                    });

                    arguments += $" -DeploymentType {program.DeploymentType}";

                    arguments += " -Bitness ";

                    if (bitnesses.Count == 2)
                    {
                        arguments += "Both";
                    }
                    else
                    {
                        arguments += bitnesses[0];
                    }

                    if (GlobalObjects.ViewModel.CmPackage.CMPSModulePath != "")
                    {
                        arguments += $" -CMPSModulePath {GlobalObjects.ViewModel.CmPackage.CMPSModulePath}";
                    }

                    if (program.DeploymentType == DeploymentType.DeployWithScript)
                    {
                        EditDeploymentScript(program);
                    }

                    if (program.ScriptName != "")
                    {
                        arguments += $" -ScriptName {program.ScriptName}";
                    }

                    if (program.ConfigurationXml != "")
                    {
                        arguments += $" -ConfigurationXml {program.ConfigurationXml}";
                    }


                    if (GlobalObjects.ViewModel.CmPackage.SiteCode != "")
                    {
                        arguments += $" -SiteCode {GlobalObjects.ViewModel.CmPackage.SiteCode} ";
                    }

                    program.CollectionNames.ToList().ForEach( async c =>
                    {

                        var agrumentsCopy = arguments; 

                        if (program.CustomName != "")
                        {
                            agrumentsCopy += $" -CustomName {program.CustomName}-{c.Replace(' ','-')}";
                        }

                        await Retry.BlockAsync(2, 1, async () =>
                        {
                            var tcs = new TaskCompletionSource<bool>();

                            if (n == 2)
                            {
                                System.IO.File.Copy(scriptPathTmp, scriptPath, true);
                            }


                            var p = new Process
                            {
                                StartInfo = new ProcessStartInfo()
                                {
                                    FileName = "cmd",
                                    Arguments = agrumentsCopy,
                                    CreateNoWindow = true,
                                    UseShellExecute = false,
                                    WorkingDirectory = currentDirectory,
                                    RedirectStandardOutput = true,
                                    RedirectStandardError = true,

                                },
                            };

                            p.EnableRaisingEvents = true;

                            //p.Exited += (sender, args) =>
                            //{
                            //    tcs.SetResult(true);
                            //    p.Dispose();
                            //};

                            p.Start();
                            p.WaitForExit();

                            var error = await p.StandardError.ReadToEndAsync();
                            if (!string.IsNullOrEmpty(error)) throw (new Exception(error));
                            n++;
                        });
                    });
                }

            });
        }

        private async Task StartCreatePackages()
        {
            try
            {
                GlobalObjects.ViewModel.BlockNavigation = true;
                StageButton.IsEnabled = false;

                Dispatcher.Invoke(() =>
                {
                    WaitCreatingPackage.Visibility = Visibility.Visible;
                });

                await CreatePackages();

                WaitCreatingPackage.Visibility = Visibility.Collapsed;
                ImgPackageCreated.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {

                WaitCreatingPackage.Visibility = Visibility.Collapsed;
                ImgPackageCreateFail.Visibility = Visibility.Visible;
                StageButton.IsEnabled = true;

                LogErrorMessage(ex);
            }
        }

        private async Task StartDownloadFiles()
        {
            try
            {
                GlobalObjects.ViewModel.BlockNavigation = true;
                StageButton.IsEnabled = false;

                Dispatcher.Invoke(() =>
                {
                    WaitFilesDownloading.Visibility = Visibility.Visible;
                });

                //await DownloadScripts();

                if (GlobalObjects.ViewModel.CmPackage.DeploymentSource == DeploymentSource.CDN)
                {
                    await DownloadChannelFiles();
                }

                WaitFilesDownloading.Visibility = Visibility.Collapsed;
                ImgFilesDownloaded.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {

                WaitFilesDownloading.Visibility = Visibility.Collapsed;
                ImgFilesDownloadFail.Visibility = Visibility.Visible;
                StageButton.IsEnabled = true;

                LogErrorMessage(ex);
            }
        }

        private async Task CreatePackages()
        {
            var n = 1;
            var CMConfig = GlobalObjects.ViewModel.CmPackage;
            var currentDirectory = Directory.GetCurrentDirectory() + @"\Scripts";

            var scriptPath = currentDirectory + $"\\Setup-CMOfficeDeployment.ps1";
            var scriptPathTmp = currentDirectory + $"\\Tmp-Setup-CMOfficeDeployment.ps1";

            var channels = new List<string>();
            var bitnesses = new List<string>();

            var arguments =
                 $"/c Powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden . .\\Setup-CMOfficeDeployment.ps1;Create-CMOfficePackage -Channels ";


            foreach (var program in CMConfig.Programs)
            {

                program.Channels.ForEach(c =>
                {
                    if (!channels.Contains(c.Branch.NewName.ToString()))
                    {
                        channels.Add(c.Branch.NewName.ToString());
                    }
                    //else
                    //{
                    //    GlobalObjects.ViewModel.CmPackage.MoveFiles = false;
                    //}
                });

                program.Bitnesses.ToList().ForEach(b =>
                {
                    if (!bitnesses.Contains(b.Name))
                    {
                        bitnesses.Add(b.Name);
                    }
                });
            }

            channels.ForEach(c =>
            {
                if (channels.IndexOf(c) < channels.Count - 1)
                {
                    arguments += $"{c},";
                }
                else
                {
                    arguments += c;
                }
            });


            if (GlobalObjects.ViewModel.CmPackage.DeploymentDirectory != "")
            {
                arguments += $" -OfficeSourceFilesPath  {GlobalObjects.ViewModel.CmPackage.DeploymentDirectory} ";
            }
        
            arguments += " -Bitness ";

            if (bitnesses.Count == 2)
            {
                arguments += "Both";
            }
            else
            {
                arguments += bitnesses[0];
            }

            arguments += $" -MoveSourceFiles ${GlobalObjects.ViewModel.CmPackage.MoveFiles} -SiteCode {GlobalObjects.ViewModel.CmPackage.SiteCode} -UpdateOnlyChangedBits ${GlobalObjects.ViewModel.CmPackage.UpdateOnlyChangedBits}";

            if (GlobalObjects.ViewModel.CmPackage.CustomPackageShareName.Trim() != "")
            {
                arguments +=
                    $" -CustomPackageShareName {GlobalObjects.ViewModel.CmPackage.CustomPackageShareName.Trim().Replace(' ','-')}";
            }

            if (GlobalObjects.ViewModel.CmPackage.CMPSModulePath != "")
            {
                arguments += $" -CMPSModulePath {GlobalObjects.ViewModel.CmPackage.CMPSModulePath}";
            }

            await Retry.BlockAsync(2, 1, async () =>
            {
                var tcs = new TaskCompletionSource<bool>();

                if (n == 2)
                {
                    System.IO.File.Copy(scriptPathTmp, scriptPath, true);
                }


                var p = new Process
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = "cmd",
                        Arguments = arguments,
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        WorkingDirectory = currentDirectory,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,

                    },
                };

                p.EnableRaisingEvents = true;

                //p.Exited += (sender, args) =>
                //{
                //    tcs.SetResult(true);
                //    p.Dispose();
                //};

                p.Start();
                p.WaitForExit();


                var error = await p.StandardError.ReadToEndAsync();
                if (!string.IsNullOrEmpty(error)) throw (new Exception(error));
                n++;
            });
        }

        private async Task StageDeployment()
        {
            ImgFilesDownloadFail.Visibility = Visibility.Collapsed;
            ImgFilesDownloaded.Visibility = Visibility.Collapsed;
            ImgProgramCreated.Visibility = Visibility.Collapsed;
            ImgProgramCreateFail.Visibility = Visibility.Collapsed;
            ImgPackageCreated.Visibility = Visibility.Collapsed; 
            ImgPackageCreateFail.Visibility = Visibility.Collapsed;
            
            try
            {
             
                GlobalObjects.ViewModel.BlockNavigation = true;
                StageButton.IsEnabled = false;

                await StartDownloadFiles();
                await StartCreatePackages();         
                await StartCreatePrograms();

                GlobalObjects.ViewModel.BlockNavigation = false;
                StageButton.IsEnabled = true;

                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = true
                });
            }
            catch (Exception ex)
            {
                StageButton.IsEnabled = true;

                LogErrorMessage(ex);
            }
           
        }

        private async Task DownloadChannelFiles()
        {

            await Task.Run(() =>
            {

            
            var n = 1;
            var driveName = GetDrive();
            var CMConfig = GlobalObjects.ViewModel.CmPackage;
            var currentDirectory = Directory.GetCurrentDirectory() + @"\Scripts";

            var scriptPath = currentDirectory + $"\\Setup-CMOfficeDeployment.ps1";
            var scriptPathTmp = currentDirectory + $"\\Tmp-Setup-CMOfficeDeployment.ps1";



            foreach (var program in CMConfig.Programs)
            {
                var channels = new List<string>();
                var languages = new List<string>();
                var bitnesses = new List<string>();
                var arguments =
                    $"/c Powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden . .\\Setup-CMOfficeDeployment.ps1;Download-CMOfficeChannelFiles ";

                program.Languages.ToList().ForEach(l =>
                {
                    if (!languages.Contains(l.Id))
                    {
                        languages.Add(l.Id);
                    }
                });

                program.Bitnesses.ToList().ForEach(b =>
                {
                    if (!bitnesses.Contains(b.Name))
                    {
                        bitnesses.Add(b.Name);
                    }
                });


                if (GlobalObjects.ViewModel.CmPackage.DeploymentSource != DeploymentSource.CDN &&
                    GlobalObjects.ViewModel.CmPackage.DeploymentDirectory != "")
                {
                    arguments +=
                        $" -OfficeFilesPath {GlobalObjects.ViewModel.CmPackage.DeploymentDirectory} -Languages ";
                }
                else
                {
                    GlobalObjects.ViewModel.CmPackage.DeploymentDirectory = driveName + "OfficeChannelFiles";
                    arguments += $" -OfficeFilesPath  {driveName}OfficeChannelFiles -Languages ";
                }

                languages.ForEach(l =>
                {
                    if (languages.IndexOf(l) < languages.Count - 1)
                    {
                        arguments += $"{l},";
                    }
                    else
                    {
                        arguments += l;
                    }
                });

                arguments += " -Bitness ";

                if (bitnesses.Count == 2)
                {
                    arguments += "Both";
                }
                else
                {
                    arguments += bitnesses[0];
                }

                program.Channels.ForEach(async c =>
                {
                    var argumentsCopy = arguments;
                    if (c.SelectedVersion == BranchVersion.Previous)
                        argumentsCopy += $" -Version {c.Branch.Versions[1].Version} ";

                    argumentsCopy += $" -Channels {c.Branch.NewName}";

                    await Retry.BlockAsync(2, 1, async () =>
                    {
                        var tcs = new TaskCompletionSource<bool>();

                        if (n == 2)
                        {
                            System.IO.File.Copy(scriptPathTmp, scriptPath, true);
                        }


                        var p = new Process
                        {
                            StartInfo = new ProcessStartInfo()
                            {
                                FileName = "cmd",
                                Arguments = argumentsCopy,
                                CreateNoWindow = true,
                                UseShellExecute = false,
                                WorkingDirectory = currentDirectory,
                                RedirectStandardOutput = true,
                                RedirectStandardError = true,

                            },
                        };

                        p.EnableRaisingEvents = true;

                        //p.Exited += (sender, args) =>
                        //{
                        //    tcs.SetResult(true);
                        //    p.Dispose();
                        //};

                        p.Start();
                        p.WaitForExit();

                        var error = await p.StandardError.ReadToEndAsync();
                        if (!string.IsNullOrEmpty(error)) throw (new Exception(error));
                        n++;
                    });
                });
            }
            });
        }

        private async Task DownloadScripts()
        {
            var downloadUrls = GlobalObjects.ViewModel.CmPackage.DownloadUrls;
            var currentDirectory = Directory.GetCurrentDirectory() + @"\Scripts";

            foreach (var url in downloadUrls)
            {
                var splitUrl = url.Url.Split('/');

                if (url.Url.Split('/')[splitUrl.Length - 2] == "DeploymentFiles")
                {
                    currentDirectory = Directory.GetCurrentDirectory() + @"\Scripts\DeploymentFiles";
                }

                var scriptPath = currentDirectory + $"\\{url.Name}";
                var scriptPathTmp = currentDirectory + $"\\Tmp-{url.Name}";


                if (!File.Exists(scriptPathTmp))
                {
                    File.Copy(scriptPath, scriptPathTmp, true);
                }

                var scriptUrl = url.Url;

                try
                {
                    await Retry.BlockAsync(5, 1, async () =>
                    {
                        using (var webClient = new WebClient())
                        {
                            await webClient.DownloadFileTaskAsync(new Uri(scriptUrl), scriptPath);
                        }
                    });
                }
                catch (Exception ex) { }
            }

        }

        private void LogErrorMessage(Exception ex)
        {
            ex.LogException(false);
            if (ErrorMessage != null)
            {
                ErrorMessage(this, new MessageEventArgs()
                {
                    Title = "Error",
                    Message = ex.Message
                });
            }
        }

        private string GetDrive()
        {
            var allDrives = DriveInfo.GetDrives().Where(d =>
                        d.DriveType != DriveType.CDRom && d.DriveType != DriveType.Removable &&
                        d.DriveType != DriveType.Unknown).ToList();

            var maxDrive = allDrives[0];
             
            allDrives.ForEach(d =>
            {
                if (maxDrive.IsReady && d.IsReady)
                    if (d.AvailableFreeSpace > maxDrive.AvailableFreeSpace)
                    maxDrive = d;
            });

            return maxDrive.Name;
        }

        private void EditDeploymentScript(CmProgram program)
        {
            var targetXmlPath = @".\Scripts\DeploymentFiles\DefaultConfiguration.xml";
            var deploymentScriptPath = @".\Scripts\DeploymentFiles\CM-OfficeDeploymentScript.ps1";
            var deploymentScriptCopyPath = @".\Scripts\DeploymentFiles\";

            var powershellScript = File.ReadAllLines(deploymentScriptPath).ToList();
            var installLineNum = powershellScript.IndexOf(" Install-OfficeClickToRun -TargetFilePath $targetFilePath");

            //Exclude-Applications -TargetFilePath $targetFilePath -ExcludeApps @("Access","Excel","Groove","InfoPath","Lync","OneDrive","OneNote","Outlook","PowerPoint","Project","Publisher","SharePointDesigner","Visio","Word")
            var excludeAppsCommand = "Exclude-Applications -TargetFilePath $targetFilePath -ExcludeApps @(";

            //Add-ProductSku -TargetFilePath $targetFilePath -Languages $languages -ProductIDs O365ProPlusRetail,O365BusinessRetail,VisioProRetail,ProjectProRetail
            var addAppsCommand = "Add-ProductSku -TargetFilePath $targetFilePath -Languages $languages -ProductIDs ";
            var languages = @"$languages = """; 

            //Add-ProductLanguage -TargetFilePath $targetFilePath -ProductIDs All -Languages fr-fr,it-it 
            var addLanguagesCommand = "Add-ProductLanguage -TargetFilePath $targetFilePath -ProductIDs All -Languages ";

            var excludeApps = new List<string>();
            var additionalApps = new List<string>();

            program.Products.ToList().ForEach(p =>
            {
                if (p.ProductAction == ProductAction.Exclude)
                    excludeApps.Add(p.Id);
                else
                {
                    additionalApps.Add(p.Id);
                }
            });

            if (excludeApps.Count > 0)
            {
                excludeApps.ForEach(e =>
                {
                    excludeAppsCommand += @" """ + e + @""" ";

                    if (excludeApps.IndexOf(e) != excludeApps.Count - 1)
                        excludeAppsCommand += ",";
                    else
                    {
                        excludeAppsCommand += ")";
                    }
                });
            }

            if (additionalApps.Count > 0)
            {
                additionalApps.ForEach(a =>
                {
                    addAppsCommand += a;
                    if (additionalApps.IndexOf(a) != additionalApps.Count - 1)
                        addAppsCommand += ",";
                });
            }

            program.Languages.ToList().ForEach(l =>
            {
                addLanguagesCommand += l.Id;
                languages += l.Id;
                if (program.Languages.IndexOf(l) != program.Languages.Count - 1)
                    addLanguagesCommand += ",";
            });

            languages += @"""";

            powershellScript[installLineNum - 1] = excludeAppsCommand.Replace("\\", "");
            powershellScript[installLineNum - 2] = addAppsCommand.Replace("\\", "");
            powershellScript[installLineNum - 3] = addLanguagesCommand.Replace("\\", "");
            powershellScript[installLineNum - 4] = languages.Replace("\\", "");

            //if (program.ScriptName != "")
            //    program.ScriptName = GlobalObjects.ViewModel.CmPackage.Programs.IndexOf(program) + "-" +
            //                         program.ScriptName;
            //else
            //{
            program.ScriptName = GlobalObjects.ViewModel.CmPackage.Programs.IndexOf(program) +
                                    "-CM-OfficeDeploymentScript.ps1";
            //}

            File.WriteAllLines(
                  deploymentScriptCopyPath + program.ScriptName, powershellScript);
        }

        #endregion
    }
}
