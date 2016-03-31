using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Media;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Implementation;
using Microsoft.OfficeProPlus.InstallGenerator.Model;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using OfficeInstallGenerator;
using OfficeInstallGenerator.Model;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace MetroDemo.ExampleViews
{
    /// <summary>V
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class LocalView : UserControl
    {
        private LanguagesDialog languagesDialog = null;
        private CancellationTokenSource _tokenSource = new CancellationTokenSource();

        public event TransitionTabEventHandler TransitionTab;
        public event MessageEventHandler InfoMessage;
        public event MessageEventHandler ErrorMessage;

        private Task _downloadTask = null;
        private int _cachedIndex = 0;
        private DateTime _lastUpdated;

        private List<Channel> items = null;
        private DownloadAdvanced advancedSettings = null;

        private OfficeLocalInstall LocalInstall { get; set; }


        public LocalView()
        {
            InitializeComponent();
        }

        private void LocalView_Loaded(object sender, RoutedEventArgs e)             
        {
            try
            {
                if (MainTabControl == null) return;
                MainTabControl.SelectedIndex = 0;

                if (GlobalObjects.ViewModel == null) return;

                GlobalObjects.ViewModel.PropertyChangeEventEnabled = false;
                LoadXml();
                GlobalObjects.ViewModel.PropertyChangeEventEnabled = true;

                LoadViewState().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async Task LoadViewState()
        {
            try
            {
                ImgLatestUpdateFail.Visibility = Visibility.Collapsed;
                UpdateStatus.Foreground = (Brush) FindResource("MessageBrush");
                ErrorRow.Visibility = Visibility.Collapsed;

                var installGenerator = new OfficeLocalInstallManager();
                LocalInstall = await installGenerator.CheckForOfficeLocalInstallAsync();

                var installedRows = Visibility.Collapsed;
                if (LocalInstall.Installed)
                {
                    InstallOffice.Visibility = Visibility.Collapsed;
                    ImgOfficeInstalled.Visibility = Visibility.Visible;
                    installedRows = Visibility.Visible;
                    VersionLabel.Content = LocalInstall.Version;

                    ProductBranch.SelectedItem = LocalInstall.Channel;

                    if (LocalInstall.LatestVersionInstalled)
                    {
                        ImgLatestUpdate.Visibility = Visibility.Visible;
                        UpdateOffice.Visibility = Visibility.Collapsed;
                        UpdateStatus.Visibility = Visibility.Collapsed;
                        UpdateButtonColumn.Width = new GridLength(45, GridUnitType.Pixel);
                    }
                    else
                    {
                        UpdateStatus.Content = "New version available  (" + LocalInstall.LatestVersion + ")";
                        UpdateButtonColumn.Width = new GridLength(90, GridUnitType.Pixel);
                        UpdateStatus.Visibility = Visibility.Visible;
                        ImgLatestUpdate.Visibility = Visibility.Collapsed;
                        UpdateOffice.Visibility = Visibility.Visible;
                    }
                    
                }
                else
                {
                    InstallOffice.Visibility = Visibility.Visible;
                    ImgOfficeInstalled.Visibility = Visibility.Collapsed;
                    installedRows = Visibility.Collapsed;
                }

                UpdateRow.Visibility = installedRows;
                VersionRow.Visibility = installedRows;
                ChannelRow.Visibility = installedRows;
                ModifyInstallRow.Visibility = installedRows;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        public async Task RunUpdateOffice()
        {
            await Task.Run(async () =>
            {
                try
                {
                    Dispatcher.Invoke(() =>
                    {
                        WaitUpdateImage.Visibility = Visibility.Visible;
                        ImgLatestUpdate.Visibility = Visibility.Collapsed;
                        UpdateOffice.Visibility = Visibility.Collapsed;
                        UpdateButtonColumn.Width = new GridLength(50, GridUnitType.Pixel);
                        UpdateStatus.Content = "Updating...";
                    });

                    var installOffice = new InstallOffice();
                    installOffice.UpdatingOfficeStatus += installOffice_UpdatingOfficeStatus;

                    await installOffice.RunOfficeUpdateAsync(LocalInstall.LatestVersion);

                    Dispatcher.Invoke(() =>
                    {
                        //InstallOffice.IsEnabled = true;
                        //ReInstallOffice.IsEnabled = true;
                        UpdateStatus.Content = "";
                        UpdateStatus.Visibility = Visibility.Collapsed;
                        UpdateButtonColumn.Width = new GridLength(50, GridUnitType.Pixel);
                        ImgLatestUpdate.Visibility = Visibility.Visible;
                        ImgOfficeInstalled.Visibility = Visibility.Visible;
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        ImgLatestUpdateFail.Visibility = Visibility.Visible;
                        ImgLatestUpdate.Visibility = Visibility.Collapsed;
                        UpdateStatus.Visibility = Visibility.Visible;
                        UpdateStatus.Content = "The update failed";
                        UpdateStatus.Foreground = (Brush) FindResource("ErrorBrush");
                        ErrorRow.Visibility = Visibility.Visible;
                        ErrorText.Text = ex.Message;
                    });

                    LogErrorMessage(ex);
                }
                finally
                {
                    Dispatcher.Invoke(() =>
                    {
                        WaitUpdateImage.Visibility = Visibility.Collapsed;
                    });
                }
            });
        }

        private void installOffice_UpdatingOfficeStatus(object sender, Microsoft.OfficeProPlus.InstallGenerator.Events.Events.UpdatingOfficeArgs e)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    UpdateStatus.Content = e.Status;
                });
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }



        public void Reset()
        {
            //ProductVersion.Text = "";
           
        }

        public void LoadXml()
        {
            Reset();

            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Add != null)
            {

            }
        }

        public void UpdateXml()
        {
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Add == null)
            {
                configXml.Add = new ODTAdd();
            }


            if (configXml.Add.Products == null)
            {
                configXml.Add.Products = new List<ODTProduct>();   
            }

            var versionText = "";
            //if (ProductVersion.SelectedIndex > -1)
            //{
            //    var version = (Build) ProductVersion.SelectedValue;
            //    versionText = version.Version;
            //}
            //else
            //{
            //    versionText = ProductVersion.Text;
            //}

            try
            {
                if (!string.IsNullOrEmpty(versionText))
                {
                    Version productVersion = null;
                    Version.TryParse(versionText, out productVersion);
                    configXml.Add.Version = productVersion;
                }
                else
                {
                    configXml.Add.Version = null;
                }
            }
            catch { }

     
        }
        
        private async Task GetBranchVersion(OfficeBranch branch, OfficeEdition officeEdition)
        {
                if (branch.Updated) return;
                var ppDownload = new ProPlusDownloader();
                var latestVersion = await ppDownload.GetLatestVersionAsync(branch.Branch.ToString(), officeEdition);

                var modelBranch = GlobalObjects.ViewModel.Branches.FirstOrDefault(b =>
                    b.Branch.ToString().ToLower() == branch.Branch.ToString().ToLower());
                if (modelBranch == null) return;
                if (modelBranch.Versions.Any(v => v.Version == latestVersion)) return;
                modelBranch.Versions.Insert(0, new Build() { Version = latestVersion });
                modelBranch.CurrentVersion = latestVersion;

                //ProductVersion.ItemsSource = modelBranch.Versions;
                //ProductVersion.SetValue(TextBoxHelper.WatermarkProperty, modelBranch.CurrentVersion);

                modelBranch.Updated = true;
        }

        private bool TransitionProductTabs(TransitionTabDirection direction)
        {
            if (direction == TransitionTabDirection.Forward)
            {
                if (MainTabControl.SelectedIndex < MainTabControl.Items.Count - 1)
                {
                    MainTabControl.SelectedIndex++;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                if (MainTabControl.SelectedIndex > 0)
                {
                    MainTabControl.SelectedIndex--;
                }
                else
                {
                    return true;
                }
            }

            return false;
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

        private void SetTabStatus(bool enabled)
        {
            Dispatcher.Invoke(() =>
            {
                OptionalTab.IsEnabled = enabled;
            });
        }

        private void LogAnaylytics(string path, string pageName)
        {
            try
            {
                GoogleAnalytics.Log(path, pageName);
            }
            catch { }
        }
        

        #region "Events"

        private async void UpdateOffice_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                await RunUpdateOffice();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void ProductBranch_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                ChangeChannel.IsEnabled = ProductBranch.SelectedItem != LocalInstall.Channel;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void InstallOffice_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        InstallOffice.IsEnabled = false;
                        ReInstallOffice.IsEnabled = false;
                    });

                    var installGenerator = new OfficeInstallExecutableGenerator();
                    installGenerator.InstallOffice(GlobalObjects.ViewModel.ConfigXmlParser.Xml);

                    Dispatcher.Invoke(() =>
                    {
                        InstallOffice.IsEnabled = true;
                        ReInstallOffice.IsEnabled = true;
                    });
                });
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void UpdatePath_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                //var dlg1 = new Ionic.Utils.FolderBrowserDialogEx
                //{
                //    Description = "Select a folder:",
                //    ShowNewFolderButton = true,
                //    ShowEditBox = true,
                //    SelectedPath = ProductUpdateSource.Text,
                //    ShowFullPathInEditBox = true,
                //    RootFolder = System.Environment.SpecialFolder.MyComputer
                //};
                ////dlg1.NewStyle = false;

                //// Show the FolderBrowserDialog.
                //var result = dlg1.ShowDialog();
                //if (result == DialogResult.OK)
                //{
                //    ProductUpdateSource.Text = dlg1.SelectedPath;
                //}
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void MainTabControl_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (GlobalObjects.ViewModel.BlockNavigation)
                {
                    MainTabControl.SelectedIndex = _cachedIndex;
                    return;
                }

                switch (MainTabControl.SelectedIndex)
                {
                    case 0:
                        LogAnaylytics("/ProductView", "Products");
                        break;
                    case 1:
                        LogAnaylytics("/ProductView", "Languages");
                        break;
                    case 2:
                        LogAnaylytics("/ProductView", "Optional");
                        break;
                    case 3:
                        LogAnaylytics("/ProductView", "Excluded");
                        break;
                }

                _cachedIndex = MainTabControl.SelectedIndex;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateXml();

                if (TransitionProductTabs(TransitionTabDirection.Forward))
                {
                    this.TransitionTab(this, new TransitionTabEventArgs()
                    {
                        Direction = TransitionTabDirection.Forward,
                        Index = 1
                    });
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void PreviousButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateXml();

                if (TransitionProductTabs(TransitionTabDirection.Back))
                {
                    this.TransitionTab(this, new TransitionTabEventArgs()
                    {
                        Direction = TransitionTabDirection.Back,
                        Index = 1
                    });
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }
        
        public BranchChangedEventHandler BranchChanged { get; set; }

        #endregion

        #region "Info"

        private void ProductInfo_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic) sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private InformationDialog informationDialog = null;

        private void LaunchInformationDialog(string sourceName)
        {
            try
            {
                if (informationDialog == null)
                {

                    informationDialog = new InformationDialog
                    {
                        Height = 500,
                        Width = 400
                    };
                    informationDialog.Closed += (o, args) =>
                    {
                        informationDialog = null;
                    };
                    informationDialog.Closing += (o, args) =>
                    {

                    };
                }
                
                informationDialog.Height = 500;
                informationDialog.Width = 400;

                var filePath = AppDomain.CurrentDomain.BaseDirectory + @"HelpFiles\" + sourceName + ".html";
                var helpFile = File.ReadAllText(filePath);

                informationDialog.HelpInfo.NavigateToString(helpFile);
                informationDialog.Launch();

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

 

        #endregion


    }




}

