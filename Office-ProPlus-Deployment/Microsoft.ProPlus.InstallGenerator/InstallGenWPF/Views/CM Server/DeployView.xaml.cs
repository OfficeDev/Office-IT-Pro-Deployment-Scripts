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
    public partial class DeployView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;
        public event MessageEventHandler ErrorMessage;



        public DeployView()
        {

            InitializeComponent();
        }

        private void DeploymentStagingView_OnLoaded(object sender, RoutedEventArgs e)
        {

        }

        private void DeployingPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
        }

        private async void DeployButton_OnClick(object sender, RoutedEventArgs e)
        {
            await BeginDeploy(); 
        }

        #region helpers

        private async Task BeginDeploy()
        {
            ImgFilesDistributed.Visibility = Visibility.Collapsed;
            ImgFilesDistributingFailed.Visibility = Visibility.Collapsed;
            ImgDeployed.Visibility = Visibility.Collapsed;
            ImgDeployFail.Visibility = Visibility.Collapsed;

            try
            {
                GlobalObjects.ViewModel.BlockNavigation = true;
                DeployButton.IsEnabled = false;

                await StartDistributeFiles();
                await StartDeployFiles();

                DeployButton.IsEnabled = true;

                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = true
                });

                GlobalObjects.ViewModel.BlockNavigation = false;

                //TODO hide Next button, need to create event 
            }
            catch (Exception ex)
            {
                GlobalObjects.ViewModel.BlockNavigation = false;

                LogErrorMessage(ex);
            }
           
        }

        private async Task DistributeFiles()
        {         
            var n = 1;
            var CMConfig = GlobalObjects.ViewModel.CmPackage; 
            var currentDirectory = Directory.GetCurrentDirectory() + @"\Scripts";

            var scriptPath = currentDirectory + $"\\Setup-CMOfficeDeployment.ps1";
            var scriptPathTmp = currentDirectory + $"\\Tmp-Setup-CMOfficeDeployment.ps1";

            var channels = new List<string>();
            var arguments = $"/c Powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden . .\\Setup-CMOfficeDeployment.ps1;Distribute-CMOfficePackage -Channels ";


            foreach (var program in CMConfig.Programs)
            {
                program.Channels.ForEach(c =>
                {
                    if (!channels.Contains(c.Branch.NewName.ToString()))
                    {
                        channels.Add(c.Branch.NewName.ToString());
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

            //arguments +=
            //    $" -DeploymentExpiryDurationInDays {GlobalObjects.ViewModel.CmPackage.DeploymentExpiryDurationInDays} ";
            arguments += $" -SiteCode {GlobalObjects.ViewModel.CmPackage.SiteCode} ";

            if (GlobalObjects.ViewModel.CmPackage.CMPSModulePath != "")
                arguments += $" -CMPSModulePath {GlobalObjects.ViewModel.CmPackage.CMPSModulePath} ";

            if (GlobalObjects.ViewModel.CmPackage.DistributionPoint != "")
            {
                arguments += $" -DistributionPoint {GlobalObjects.ViewModel.CmPackage.DistributionPoint}";
            }
            else
            {
                arguments += $" -DistributionPointGroupName  {GlobalObjects.ViewModel.CmPackage.DistributionPointGroupName}";
            }

            arguments += " -WaitForDistributionToFinish $true ";

            await Retry.Block(2, 1, async () =>
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

                p.Exited += (sender, args) =>
                {
                    tcs.SetResult(true);
                    p.Dispose();
                };

                p.Start();


                var error = await p.StandardError.ReadToEndAsync();
                if (!string.IsNullOrEmpty(error)) throw (new Exception(error));
                n++;
            });
        }

        private async Task DeployPrograms()
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
                var arguments = $"/c Powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden . .\\Setup-CMOfficeDeployment.ps1;Deploy-CMOfficeProgram -Channel ";


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

                arguments += " -Bitness ";

                if (bitnesses.Count == 2)
                {
                    arguments += "Both";
                }
                else
                {
                    arguments += bitnesses[0];
                }

                arguments += $" -ProgramType {program.DeploymentType} ";

                //if (program.DeploymentType == DeploymentType.DeployWithScript)
                //{
                //    //EditDeploymentScript(program);
                //}
                //else
                //{
                //    //edit configuration.xml
                //}

                arguments += $" -SiteCode {GlobalObjects.ViewModel.CmPackage.SiteCode}";
                arguments += $" -DeploymentPurpose {program.DeploymentPurpose} ";


                if (GlobalObjects.ViewModel.CmPackage.CMPSModulePath != "")
                    arguments += $" -CMPSModulePath {GlobalObjects.ViewModel.CmPackage.CMPSModulePath} ";

               

                program.CollectionNames.ToList().ForEach(async c =>
                {
                    var argumentsCopy = arguments;
                    argumentsCopy += $" -Collection '{c}' ";

                    if (program.CustomName != "")
                        argumentsCopy += $" -CustomName {program.CustomName}-{c.Replace(' ','-')}";


                   await Retry.Block(2, 1, async () =>
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

                        p.Exited += (sender, args) =>
                        {
                            tcs.SetResult(true);
                            p.Dispose();
                        };

                        p.Start();


                        var error = await p.StandardError.ReadToEndAsync();
                        if (!string.IsNullOrEmpty(error)) throw (new Exception(error));
                        n++;
                    });
                });
            }

        }

        private async Task StartDistributeFiles()
        {
            try
            {
                GlobalObjects.ViewModel.BlockNavigation = true;

                Dispatcher.Invoke(() =>
                {
                    WaitFilesDistributing.Visibility = Visibility.Visible;
                });

                await DistributeFiles();

                WaitFilesDistributing.Visibility = Visibility.Collapsed;
                ImgFilesDistributed.Visibility = Visibility.Visible;

            }
            catch (Exception ex)
            {

                WaitFilesDistributing.Visibility = Visibility.Collapsed;
                ImgFilesDistributingFailed.Visibility = Visibility.Visible;
                DeployButton.IsEnabled = true;

                LogErrorMessage(ex); ;
            }
        }

        private async Task StartDeployFiles()
        {
            try
            {
                GlobalObjects.ViewModel.BlockNavigation = true;

                Dispatcher.Invoke(() =>
                {
                    WaitDeployingPrograms.Visibility = Visibility.Visible;
                });

                await DeployPrograms();

                WaitDeployingPrograms.Visibility = Visibility.Collapsed;
                ImgDeployed.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {

                WaitDeployingPrograms.Visibility = Visibility.Collapsed;
                ImgDeployFail.Visibility = Visibility.Visible;
                DeployButton.IsEnabled = true;

                LogErrorMessage(ex); ;
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





        #endregion


    }
}
