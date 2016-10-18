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
            var cm = GlobalObjects.ViewModel.CmPackage;

        }

        private async void DeployButton_OnClick(object sender, RoutedEventArgs e)
        {
            await DeployPackage(); 
        }

        #region helpers

        private async Task DeployPackage()
        {




            try
            {
                GlobalObjects.ViewModel.BlockNavigation = true;
                StageButton.IsEnabled = false;

                Dispatcher.Invoke(() =>
                {
                    WaitFilesDownloading.Visibility = Visibility.Visible;
                });

                await DownloadScripts();
                await DownloadChannelFiles();
                //await DownloadChannelFiles();
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

        private async Task DownloadChannelFiles()
        {         
            var n = 1;
            var CMConfig = GlobalObjects.ViewModel.CmPackage; 
            var currentDirectory = Directory.GetCurrentDirectory() + @"\Scripts";

            var scriptPath = currentDirectory + $"\\Setup-CMOfficeDeployment.ps1";
            var scriptPathTmp = currentDirectory + $"\\Tmp-Setup-CMOfficeDeployment.ps1";

            var arguments = $"/c Powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden . .\\Setup-CMOfficeDeployment.ps1;Download-CMOfficeChannelFiles -Channels ";
            var channels = new List<string>();
            var languages = new List<string>();
            var bitnesses = new List<string>();

            CMConfig.Programs.ToList().ForEach(p =>
            {
                p.Channels.ForEach(c =>
                {
                    if (!channels.Contains(c.Branch.NewName.ToString()))
                    {
                        channels.Add(c.Branch.NewName.ToString());
                    }
                });

                p.Languages.ToList().ForEach(l =>
                {
                    if (!languages.Contains(l.Id))
                    {
                        languages.Add(l.Id);
                    }
                });

                p.Bitnesses.ToList().ForEach(b =>
                {
                    if (!bitnesses.Contains(b.Name))
                    {
                        bitnesses.Add(b.Name);
                    }
                });

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

            arguments += " -OfficeFilesPath C:\\OfficeChannels -Languages ";

            languages.ForEach(c =>
            {
                if (languages.IndexOf(c) < languages.Count - 1)
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

        #endregion
    }
}
