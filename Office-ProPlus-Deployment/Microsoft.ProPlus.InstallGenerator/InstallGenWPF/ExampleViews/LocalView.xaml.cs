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
using Micorosft.OfficeProPlus.ConfigurationXml.Enums;
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

                if (GlobalObjects.ViewModel.RunLocalConfigs)
                {
                    LoadViewState().ConfigureAwait(false);
                }

                LoadCurrentXml();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        public void LoadCurrentXml()
        {
            if (xmlBrowser == null) return;

            if (GlobalObjects.ViewModel == null) return;
            if (GlobalObjects.ViewModel.ConfigXmlParser != null)
            {
                var configXml = GlobalObjects.ViewModel.ConfigXmlParser;

                if (!string.IsNullOrEmpty(configXml.Xml))
                {
                    xmlBrowser.XmlDoc = configXml.Xml;
                }
            }
        }


        private async Task LoadViewState()
        {
            try
            {
                await Retry.BlockAsync(10, 1, async () => {
                    ErrorRow.Visibility = Visibility.Collapsed;
                    SetItemState(LocalViewItem.Install, LocalViewState.Default);

                    var installGenerator = new OfficeLocalInstallManager();
                    LocalInstall = await installGenerator.CheckForOfficeLocalInstallAsync();

                    if (LocalInstall.Installed)
                    {
                        SetItemState(LocalViewItem.Install, LocalViewState.Success);

                        VersionLabel.Content = LocalInstall.Version;
                        
                        var selectIndex = 0;
                        for (var i = 0; i < ProductBranch.Items.Count; i++)
                        {
                            var item = (OfficeBranch) ProductBranch.Items[i];
                            if (item == null) continue;
                            if (item.NewName.ToLower() != LocalInstall.Channel.ToLower()) continue;
                            selectIndex = i;
                            break;
                        }

                        BranchChanged(this, new BranchChangedEventArgs()
                        {
                            BranchName = LocalInstall.Channel
                        });

                        ProductBranch.SelectedIndex = selectIndex;

                        var installOffice = new InstallOffice();
                        if (installOffice.IsUpdateRunning())
                        {
                            await RunUpdateOffice();
                        }
                        else
                        {
                            if (LocalInstall.LatestVersionInstalled)
                            {
                                SetItemState(LocalViewItem.Update, LocalViewState.Success);
                            }
                            else
                            {
                                SetItemState(LocalViewItem.Update, LocalViewState.Action);
                                UpdateStatus.Content = "New version available  (" + LocalInstall.LatestVersion + ")";
                            }
                        }
                    }
                    else
                    {
                        SetItemState(LocalViewItem.Install, LocalViewState.Action);
                    }
                });
            }
            catch (Exception ex)
            {
                SetItemState(LocalViewItem.Install, LocalViewState.Fail);
                Dispatcher.Invoke(() =>
                {
                    ErrorText.Text = ex.Message;
                });
                LogErrorMessage(ex);
            }
        }

        private void SetItemState(LocalViewItem item, LocalViewState state)
        {
            Dispatcher.Invoke(() =>
            {
                ErrorRow.Visibility = Visibility.Collapsed;
                switch (item)
                {
                    case LocalViewItem.Install:
                        var installedRows = Visibility.Visible;

                        InstallOffice.Visibility = Visibility.Collapsed;
                        ImgOfficeInstalled.Visibility = Visibility.Collapsed;
                        WaitInstallImage.Visibility = Visibility.Collapsed;
                        ImgLatestInstallFail.Visibility = Visibility.Collapsed;
                        UpdateButtonColumn.Width = new GridLength(45, GridUnitType.Pixel);
                        ReInstallOffice.IsEnabled = true;
                        UpdateOffice.IsEnabled = true;
                        RetryUpdateOffice.IsEnabled = true;
                        ProductBranch.IsEnabled = true;
                        switch (state)
                        {
                            case LocalViewState.Default:
                                WaitInstallImage.Visibility = Visibility.Visible;
                                installedRows = Visibility.Collapsed;
                                break;
                            case LocalViewState.Action:
                                InstallOffice.Visibility = Visibility.Visible;
                                installedRows = Visibility.Collapsed;
                                UpdateButtonColumn.Width = new GridLength(90, GridUnitType.Pixel);
                                break;
                            case LocalViewState.Fail:
                                ImgLatestInstallFail.Visibility = Visibility.Visible;
                                ErrorRow.Visibility = Visibility.Visible;
                                installedRows = UpdateRow.Visibility;
                                break;
                            case LocalViewState.Success:
                                ImgOfficeInstalled.Visibility = Visibility.Visible;
                                break;
                            case LocalViewState.Wait:
                                WaitInstallImage.Visibility = Visibility.Visible;
                                ReInstallOffice.IsEnabled = false;
                                UpdateOffice.IsEnabled = false;
                                RetryUpdateOffice.IsEnabled = false;
                                ProductBranch.IsEnabled = false;
                                break;
                        }
                        UpdateRow.Visibility = installedRows;
                        VersionRow.Visibility = installedRows;
                        ChannelRow.Visibility = installedRows;
                        ModifyInstallRow.Visibility = installedRows;
                        break;
                    case LocalViewItem.Update:
                        UpdateOffice.Visibility = Visibility.Collapsed;
                        UpdateStatus.Foreground = (Brush)FindResource("MessageBrush");
                        ImgLatestUpdate.Visibility = Visibility.Collapsed;
                        ImgLatestUpdateFail.Visibility = Visibility.Collapsed;
                        WaitUpdateImage.Visibility = Visibility.Collapsed;
                        ReInstallOffice.IsEnabled = true;
                        UpdateOffice.IsEnabled = true;
                        RetryUpdateOffice.IsEnabled = true;
                        ProductBranch.IsEnabled = true;
                        RetryButtonColumn.Width = new GridLength(0, GridUnitType.Pixel);
                        UpdateStatus.Visibility = Visibility.Collapsed;
                        switch (state)
                        {
                            case LocalViewState.Action:
                                UpdateOffice.Visibility = Visibility.Visible;
                                UpdateButtonColumn.Width = new GridLength(90, GridUnitType.Pixel);
                                UpdateStatus.Visibility = Visibility.Visible;
                                break;
                            case LocalViewState.Fail:
                                ImgLatestUpdateFail.Visibility = Visibility.Visible;
                                UpdateStatus.Foreground = (Brush)FindResource("ErrorBrush");
                                ErrorRow.Visibility = Visibility.Visible;
                                UpdateButtonColumn.Width = new GridLength(45, GridUnitType.Pixel);
                                RetryButtonColumn.Width = new GridLength(90, GridUnitType.Pixel);
                                UpdateStatus.Visibility = Visibility.Visible;
                                break;
                            case LocalViewState.Success:
                                ImgLatestUpdate.Visibility = Visibility.Visible;
                                UpdateButtonColumn.Width = new GridLength(45, GridUnitType.Pixel);
                                break;
                            case LocalViewState.Wait:
                                WaitUpdateImage.Visibility = Visibility.Visible;
                                UpdateStatus.Visibility = Visibility.Visible;
                                UpdateButtonColumn.Width = new GridLength(45, GridUnitType.Pixel);
                                ReInstallOffice.IsEnabled = false;
                                UpdateOffice.IsEnabled = false;
                                RetryUpdateOffice.IsEnabled = false;
                                ProductBranch.IsEnabled = false;
                                break;
                        }
                        break;
                   
                }
            });
        }

        public async Task RunUpdateOffice()
        {
            await Task.Run(async () =>
            {
                try
                {
                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus.Content = "Updating...";
                    });

                    SetItemState(LocalViewItem.Update, LocalViewState.Wait);

                    var installOffice = new InstallOffice();
                    installOffice.UpdatingOfficeStatus += installOffice_UpdatingOfficeStatus;

                    var currentChannel = LocalInstall.Channel;
                    if (!installOffice.IsUpdateRunning())
                    {
                        var ppDownloader = new ProPlusDownloader();
                        var baseUrl =
                            await ppDownloader.GetChannelBaseUrlAsync(currentChannel, OfficeEdition.Office32Bit);
                        if (string.IsNullOrEmpty(baseUrl))
                            throw (new Exception(string.Format("Cannot find BaseUrl for Channel: {0}", currentChannel)));

                        installOffice.ChangeUpdateSource(baseUrl);
                    }

                    await installOffice.RunOfficeUpdateAsync(LocalInstall.LatestVersion);
                    
                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus.Content = "";
                    });

                    var installGenerator = new OfficeLocalInstallManager();
                    LocalInstall = await installGenerator.CheckForOfficeLocalInstallAsync();
                    if (LocalInstall.Installed)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            VersionLabel.Content = LocalInstall.Version;
                            ProductBranch.SelectedItem = LocalInstall.Channel;
                        });

                        if (LocalInstall.LatestVersionInstalled)
                        {
                            SetItemState(LocalViewItem.Update, LocalViewState.Success);
                        }
                        else
                        {
                            SetItemState(LocalViewItem.Update, LocalViewState.Action);
                            Dispatcher.Invoke(() =>
                            {
                                UpdateStatus.Content = "New version available  (" + LocalInstall.LatestVersion + ")";
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    SetItemState(LocalViewItem.Update, LocalViewState.Fail);
                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus.Content = "The update failed";
                        ErrorText.Text = ex.Message;
                    });

                    LogErrorMessage(ex);
                }
                finally
                {
                    var installOffice = new InstallOffice();
                    installOffice.ResetUpdateSource();
                }
            });
        }

        public async Task ChangeOfficeChannel()
        {
            await Task.Run(async () =>
            {
                try
                {
                    var newChannel = "";
                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus.Content = "Updating...";
                        newChannel = ((OfficeBranch) ProductBranch.SelectedItem).NewName;
                        ChangeChannel.IsEnabled = false;
                    });

                    SetItemState(LocalViewItem.Update, LocalViewState.Wait);

                    var installOffice = new InstallOffice();
                    installOffice.UpdatingOfficeStatus += installOffice_UpdatingOfficeStatus;

                    var ppDownloader = new ProPlusDownloader();
                    var baseUrl = await ppDownloader.GetChannelBaseUrlAsync(newChannel, OfficeEdition.Office32Bit);
                    if (string.IsNullOrEmpty(baseUrl))
                        throw (new Exception(string.Format("Cannot find BaseUrl for Channel: {0}", newChannel)));

                    var latestChannelVersion = await ppDownloader.GetLatestVersionAsync(newChannel, OfficeEdition.Office32Bit);

                    installOffice.ChangeUpdateSource(baseUrl);

                    await installOffice.RunOfficeUpdateAsync(latestChannelVersion);

                    installOffice.ChangeBaseCdnUrl(baseUrl);

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus.Content = "";
                    });
                    
                    var installGenerator = new OfficeLocalInstallManager();
                    LocalInstall = await installGenerator.CheckForOfficeLocalInstallAsync();
                    if (LocalInstall.Installed)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            VersionLabel.Content = LocalInstall.Version;
                            ProductBranch.SelectedItem = LocalInstall.Channel;
                        });

                        if (LocalInstall.LatestVersionInstalled)
                        {
                            SetItemState(LocalViewItem.Update, LocalViewState.Success);
                        }
                        else
                        {
                            SetItemState(LocalViewItem.Update, LocalViewState.Action);
                            Dispatcher.Invoke(() =>
                            {
                                UpdateStatus.Content = "New version available  (" + LocalInstall.LatestVersion + ")";
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    SetItemState(LocalViewItem.Update, LocalViewState.Fail);
                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus.Content = "The update failed";
                        ErrorText.Text = ex.Message;
                    });

                    LogErrorMessage(ex);
                }
                finally
                {
                    var installOffice = new InstallOffice();
                    installOffice.ResetUpdateSource();

                    Dispatcher.Invoke(() =>
                    {
                        ChangeChannel.IsEnabled = true;
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
                    UpdateStatus.Visibility = Visibility.Visible;
                    UpdateStatus.Content = e.Status;
                });
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }


        #region Page Functions

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

        #endregion

        #region "Events"

        private void xmlBrowser_Loaded(object sender, RoutedEventArgs e)
        {

        }


        private async void ChangeChannel_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                GlobalObjects.ViewModel.BlockNavigation = true;
                await ChangeOfficeChannel();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
            finally
            {
                GlobalObjects.ViewModel.BlockNavigation = false;
            }
        }

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
                var selectedBranch = (OfficeBranch) ProductBranch.SelectedItem;
                if (selectedBranch != null && LocalInstall != null)
                {
                    ChangeChannel.IsEnabled = selectedBranch.NewName != LocalInstall.Channel;
                }
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


    public enum LocalViewItem
    {
        Install = 0,
        Update = 1
    }

    public enum LocalViewState
    {
        Default = 0,
        Success = 1,
        Fail = 2,
        Action = 3,
        Wait = 5,
        Running = 6
    }

}

