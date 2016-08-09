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
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Enums;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Extentions;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Implementation;
using Microsoft.OfficeProPlus.InstallGenerator.Model;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using OfficeInstallGenerator;
using OfficeInstallGenerator.Model;
using File = System.IO.File;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace MetroDemo.ExampleViews
{
    /// <summary>V
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class LocalView : UserControl
    {
        #region Declarations
        private LanguagesDialog languagesDialog = null;
        private CancellationTokenSource _tokenSource = new CancellationTokenSource();

        public event TransitionTabEventHandler TransitionTab;
        public event MessageEventHandler InfoMessage;
        public event MessageEventHandler ErrorMessage;

        public MetroWindow MainWindow { get; set; }

        private Task _downloadTask = null;
        private int _cachedIndex = 0;
        private DateTime _lastUpdated;

        private List<Channel> items = null;
        private DownloadAdvanced advancedSettings = null;

        private OfficeInstallation LocalInstall { get; set; }
        private bool FirstRun = true;
        #endregion

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

                var currentIndex = ProductBranch.SelectedIndex;
                ProductBranch.ItemsSource = GlobalObjects.ViewModel.Branches;
                if (currentIndex == -1) currentIndex = 0;
                ProductBranch.SelectedIndex = currentIndex;

                GlobalObjects.ViewModel.PropertyChangeEventEnabled = false;
                LoadXml();
                GlobalObjects.ViewModel.PropertyChangeEventEnabled = true;

                if (GlobalObjects.ViewModel.ApplicationMode == ApplicationMode.ManageLocal)
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
                    Dispatcher.Invoke(() =>
                    {
                       ErrorRow.Visibility = Visibility.Collapsed;
                    });

                    SetItemState(LocalViewItem.Install, LocalViewState.Default);

                    var installGenerator = new OfficeLocalInstallManager(); 
                    LocalInstall = await installGenerator.CheckForOfficeInstallAsync();

                    if (LocalInstall.Installed)
                    {
                        SetItemState(LocalViewItem.Install, LocalViewState.Success);
                        SetItemState(LocalViewItem.Uninstall, LocalViewState.Action);

                        Dispatcher.Invoke(() =>
                        {
                            VersionLabel.Content = "(" + LocalInstall.Version + ")";
                            ChannelLabel.Content = LocalInstall.Channel;

                           // ChannelLabel.Content = "First Release Deferred";

                            if (ChannelLabel.Content.ToString().Length < 10)
                            {
                                ChannelName.Width = new GridLength(100);
                            }
                            else if (ChannelLabel.Content.ToString().Length < 20)
                            {
                                ChannelName.Width = new GridLength(150);
                            }
                            else 
                            {
                                ChannelName.Width = new GridLength(200);
                            }

                            var selectIndex = 0;
                            if (LocalInstall.Channel != null)
                            {
                                for (var i = 0; i < ProductBranch.Items.Count; i++)
                                {
                                    var item = (OfficeBranch) ProductBranch.Items[i];
                                    if (item == null) continue;
                                    if (item.NewName.ToLower() != LocalInstall.Channel.ToLower()) continue;
                                    selectIndex = i;
                                    break;
                                }
                            }

                            BranchChanged(this, new BranchChangedEventArgs()
                            {
                                BranchName = LocalInstall.Channel
                            });

                            ProductBranch.SelectedIndex = selectIndex;
                        });

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
                                Dispatcher.Invoke(() =>
                                {
                                    UpdateStatus.Content = "New version available  (" + LocalInstall.LatestVersion + ")";
                                });
                            }
                        }

                        Dispatcher.Invoke(() =>
                        {
                            ChangeChannel.IsEnabled = false;
                        });
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
                var latestInstalled = false;
                if (LocalInstall != null)
                {
                    latestInstalled = LocalInstall.LatestVersionInstalled;
                }
                if (!latestInstalled)
                {
                    if (RetryButtonColumn.Width.Value > 0)
                    {
                        latestInstalled = true;
                    }
                }

                ErrorRow.Visibility = Visibility.Collapsed;
                switch (item)
                {
                    case LocalViewItem.Install:
                        var installedRows = Visibility.Visible;
                        bool isNotfifteen = true;
                        InstallOffice.Visibility = Visibility.Collapsed;
                        ImgOfficeInstalled.Visibility = Visibility.Collapsed;
                        WaitInstallImage.Visibility = Visibility.Collapsed;
                        ImgLatestInstallFail.Visibility = Visibility.Collapsed;                        

                        UpdateButtonColumn.Width = latestInstalled ? new GridLength(45, GridUnitType.Pixel) : new GridLength(90, GridUnitType.Pixel);

                        ReInstallOffice.IsEnabled = true;
                        UpdateOffice.IsEnabled = true;
                        RetryUpdateOffice.IsEnabled = true;
                        ProductBranch.IsEnabled = true;
                        UnInstallOffice.IsEnabled = true;
                        switch (state)
                        {
                            case LocalViewState.InstallingOffice:
                                WaitInstallImage.Visibility = Visibility.Visible;
                                ReInstallOffice.IsEnabled = false;
                                UpdateOffice.IsEnabled = false;
                                RetryUpdateOffice.IsEnabled = false;
                                ProductBranch.IsEnabled = false;
                                installedRows = Visibility.Collapsed;
                                break;
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
                                //ErrorRow.Visibility = Visibility.Visible;
                                installedRows = UpdateRow.Visibility;
                                break;
                            case LocalViewState.Success:
                                ImgOfficeInstalled.Visibility = Visibility.Visible;
                                if (LocalInstall != null)
                                {
                                    if (LocalInstall.Version.StartsWith("15."))
                                    {
                                        ChannelRow.Visibility = Visibility.Collapsed;
                                        ModifyInstallRow.Visibility = Visibility.Collapsed;
                                        isNotfifteen = false;
                                    }
                                }
                                break;
                            case LocalViewState.Wait:
                                WaitInstallImage.Visibility = Visibility.Visible;
                                ReInstallOffice.IsEnabled = false;
                                UpdateOffice.IsEnabled = false;
                                RetryUpdateOffice.IsEnabled = false;
                                ProductBranch.IsEnabled = false;
                                UnInstallOffice.IsEnabled = false;

                                break;
                        }
                        UpdateRow.Visibility = installedRows;
                        VersionRow.Visibility = installedRows;                        
                        ModifyUninstallRow.Visibility = installedRows;
                        //maybe set bool switch here to see if rows collapsed due to version 15.x.x.x
                        if (isNotfifteen)
                        {
                            ChannelRow.Visibility = installedRows;
                            ModifyInstallRow.Visibility = installedRows;
                        }
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
                                UnInstallOffice.IsEnabled = true;
                                ChangeChannel.IsEnabled = true;
                                NewVersion.IsEnabled = true;
                                break;
                            case LocalViewState.Fail:
                                ImgLatestUpdateFail.Visibility = Visibility.Visible;
                                UpdateStatus.Foreground = (Brush)FindResource("ErrorBrush");
                                //ErrorRow.Visibility = Visibility.Visible;
                                UpdateButtonColumn.Width = new GridLength(45, GridUnitType.Pixel);
                                RetryButtonColumn.Width = new GridLength(90, GridUnitType.Pixel);
                                UpdateStatus.Visibility = Visibility.Visible;
                                UnInstallOffice.IsEnabled = true;
                                ChangeChannel.IsEnabled = true;
                                NewVersion.IsEnabled = true;
                                break;
                            case LocalViewState.Success:
                                ImgLatestUpdate.Visibility = Visibility.Visible;
                                UpdateButtonColumn.Width = new GridLength(45, GridUnitType.Pixel);
                                UnInstallOffice.IsEnabled = true;
                                ChangeChannel.IsEnabled = true;
                                NewVersion.IsEnabled = true;
                                break;
                            case LocalViewState.Wait:
                                WaitUpdateImage.Visibility = Visibility.Visible;
                                UpdateStatus.Visibility = Visibility.Visible;
                                UpdateButtonColumn.Width = new GridLength(45, GridUnitType.Pixel);
                                ReInstallOffice.IsEnabled = false;
                                UpdateOffice.IsEnabled = false;
                                RetryUpdateOffice.IsEnabled = false;
                                ProductBranch.IsEnabled = false;
                                UnInstallOffice.IsEnabled = false;
                                ChangeChannel.IsEnabled = false;
                                NewVersion.IsEnabled = false;
                                break;
                        }
                        break;
                    case LocalViewItem.Uninstall:
                        UninstallIconColumn.Width = new GridLength(0, GridUnitType.Pixel);
                        ImgRemoveFail.Visibility = Visibility.Collapsed;
                        WaitRemoveImage.Visibility = Visibility.Collapsed;
                        switch (state)
                        {
                            case LocalViewState.Success:
                                InstallOffice.Visibility = Visibility.Visible;
                                UpdateButtonColumn.Width = new GridLength(90, GridUnitType.Pixel);
                                UpdateRow.Visibility = Visibility.Collapsed;
                                VersionRow.Visibility = Visibility.Collapsed;
                                ChannelRow.Visibility = Visibility.Collapsed;
                                ModifyInstallRow.Visibility = Visibility.Collapsed;
                                ModifyUninstallRow.Visibility = Visibility.Collapsed;
                                ImgOfficeInstalled.Visibility = Visibility.Collapsed;
                                ProductBranch.IsEnabled = true;
                                break;
                            case LocalViewState.Action:
                                UnInstallOffice.Visibility = Visibility.Visible;
                                ProductBranch.IsEnabled = true;
                                break;
                            case LocalViewState.Fail:
                                ImgRemoveFail.Visibility = Visibility.Visible;
                                ProductBranch.IsEnabled = true;
                                break;
                            case LocalViewState.Wait:
                                UnInstallOffice.Visibility = Visibility.Collapsed;
                                UninstallIconColumn.Width = new GridLength(50, GridUnitType.Pixel);
                                WaitRemoveImage.Visibility = Visibility.Visible;
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
                InstallOffice installOffice = null;
                try
                {
                    installOffice = new InstallOffice();
                    installOffice.UpdatingOfficeStatus += installOffice_UpdatingOfficeStatus;

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus.Content = "Updating...";
                        ShowVersion.IsEnabled = false;
                        ChangeChannel.IsEnabled = false;
                    });

                    GlobalObjects.ViewModel.BlockNavigation = true;

                    SetItemState(LocalViewItem.Update, LocalViewState.Wait);

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

                    SetSelectedVersion();

                    await installOffice.RunOfficeUpdateAsync(LocalInstall.LatestVersion);
                    
                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus.Content = "";
                    });

                    var installGenerator = new OfficeInstallManager(); 
                    LocalInstall = await installGenerator.CheckForOfficeInstallAsync();
                    if (LocalInstall.Installed)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            VersionLabel.Content = LocalInstall.Version;
                            ChannelLabel.Content = LocalInstall.Channel;
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
                    installOffice.ResetUpdateSource();
                    Dispatcher.Invoke(() =>
                    {
                        ShowVersion.IsEnabled = true;
                        ChangeChannel.IsEnabled = true;
                    });
                    GlobalObjects.ViewModel.BlockNavigation = false;
                }
            });
        }

        public async Task RunInstallOffice()
        {
            await Task.Run(async () =>
            {
                try
                {
                    Dispatcher.Invoke(() =>
                    {
                        InstallOffice.IsEnabled = false;
                        ReInstallOffice.IsEnabled = false;
                    });
                    GlobalObjects.ViewModel.BlockNavigation = true;
                    GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml.Display.Level = DisplayLevel.Full;

                    FirstRun = false;

                    SetItemState(LocalViewItem.Install, LocalViewState.InstallingOffice);

                    var installGenerator = new OfficeInstallExecutableGenerator();
                    installGenerator.InstallOffice(GlobalObjects.ViewModel.ConfigXmlParser.Xml);

                    await LoadViewState();

                    Dispatcher.Invoke(() =>
                    {
                        InstallOffice.IsEnabled = true;
                        ReInstallOffice.IsEnabled = true;
                    });
                }
                catch (Exception ex)
                {
                    SetItemState(LocalViewItem.Install, LocalViewState.Fail);
                    LogErrorMessage(ex);
                }
                finally
                {
                    GlobalObjects.ViewModel.BlockNavigation = false;
                }
            });
        }

        private void SetSelectedVersion()
        {
            Dispatcher.Invoke(() =>
            {
                if (NewVersionRow.Visibility == Visibility.Visible)
                {
                    var versionFound = false;
                    for (var i = 0; i < NewVersion.Items.Count; i++)
                    {
                        var item = NewVersion.Items[i];
                        if (item == null) continue;

                        var version = (Build)item;
                        if (version.Version != LocalInstall.LatestVersion) continue;
                        NewVersion.SelectedIndex = i;
                        versionFound = true;
                        break;
                    }
                    if (!versionFound)
                    {
                        NewVersion.Text = LocalInstall.LatestVersion;
                    }
                }
            });
        }

        public async Task ChangeOfficeChannel()
        {
            await Task.Run(async () =>
            {
                var installOffice = new InstallOffice();
                try
                {
                    installOffice = new InstallOffice();
                    installOffice.UpdatingOfficeStatus += installOffice_UpdatingOfficeStatus;

                    var newChannel = "";
                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus.Content = "Updating...";
                        newChannel = ((OfficeBranch) ProductBranch.SelectedItem).NewName;
                        ChangeChannel.IsEnabled = false;
                        NewVersion.IsEnabled = false;
                    });

                    SetItemState(LocalViewItem.Update, LocalViewState.Wait);
                    
                    var ppDownloader = new ProPlusDownloader();
                    var baseUrl = await ppDownloader.GetChannelBaseUrlAsync(newChannel, OfficeEdition.Office32Bit);
                    if (string.IsNullOrEmpty(baseUrl))
                        throw (new Exception(string.Format("Cannot find BaseUrl for Channel: {0}", newChannel)));

                    var channelToChangeTo = "";
                    if (NewVersionRow.Visibility != Visibility.Visible)
                    {
                        channelToChangeTo =
                            await ppDownloader.GetLatestVersionAsync(newChannel, OfficeEdition.Office32Bit);
                    }
                    else
                    {
                        Dispatcher.Invoke(() =>
                        {
                            var manualVersion = NewVersion.Text;

                            if (string.IsNullOrEmpty(manualVersion) && NewVersion.SelectedItem != null)
                            {
                                manualVersion = ((Build)NewVersion.SelectedItem).Version;
                            }
                            if (!string.IsNullOrEmpty(manualVersion))
                            {
                                channelToChangeTo = manualVersion;
                            }
                        });
                    }

                    if (string.IsNullOrEmpty(channelToChangeTo))
                    {
                        throw (new Exception("Version required"));
                    }
                    else
                    {
                        if (!channelToChangeTo.IsValidVersion())
                        {
                            throw (new Exception(string.Format("Invalid Version: {0}", channelToChangeTo)));
                        }
                    }

                    await installOffice.ChangeOfficeChannel(channelToChangeTo, baseUrl);

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus.Content = "";
                    });
                    
                    var installGenerator = new OfficeInstallManager();
                    LocalInstall = await installGenerator.CheckForOfficeInstallAsync();
                    if (LocalInstall.Installed)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            VersionLabel.Content = LocalInstall.Version;
                            ChannelLabel.Content = LocalInstall.Channel;
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
                        RetryButtonColumn.Width = new GridLength(0, GridUnitType.Pixel);
                    });
                    LogErrorMessage(ex);
                }
                finally
                {
                    Dispatcher.Invoke(() =>
                    {
                        ChangeChannel.IsEnabled = true;
                        NewVersion.IsEnabled = true;
                    });
                }
            });
        }

        public async Task UninstallOffice()
        {
            await Task.Run(async () =>
            {
                try {
                    GlobalObjects.ViewModel.BlockNavigation = true;

                    Dispatcher.Invoke(() =>
                    {
                        InstallOffice.IsEnabled = false;
                        ReInstallOffice.IsEnabled = false;
                    });

                    GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml.Display.Level = DisplayLevel.Full;

                    SetItemState(LocalViewItem.Uninstall, LocalViewState.Wait);

                    var installGenerator = new OfficeInstallManager();
                    string installVer = "2016";
                    if (LocalInstall.Version.StartsWith("15."))
                    {
                        installVer = "2013";
                    }
                    installGenerator.UninstallOffice(installVer);

                    SetItemState(LocalViewItem.Uninstall, LocalViewState.Success);

                    await LoadViewState();

                    Dispatcher.Invoke(() =>
                    {
                        InstallOffice.IsEnabled = true;
                        ReInstallOffice.IsEnabled = true;
                    });
                }
                catch (Exception ex)
                {
                    SetItemState(LocalViewItem.Uninstall, LocalViewState.Fail);
                    LogErrorMessage(ex);
                }
                finally
                {
                    GlobalObjects.ViewModel.BlockNavigation = false;
                }
            });
        }

        public async Task UpdateVersions()
        {
            if (ProductBranch.SelectedItem == null) return;
            var branch = (OfficeBranch)ProductBranch.SelectedItem;
            NewVersion.ItemsSource = branch.Versions;
            NewVersion.SetValue(TextBoxHelper.WatermarkProperty, branch.CurrentVersion);

            var edition = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml.Add.OfficeClientEdition;

            var officeEdition = OfficeEdition.Office32Bit;
            if (edition == OfficeClientEdition.Office64Bit)
            {
                officeEdition = OfficeEdition.Office64Bit;
            }

            await GetBranchVersion(branch, officeEdition);
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

        private string GetSelectedVersion()
        {
            var selVersionText = "";
            if (NewVersion.SelectedItem != null)
            {
                selVersionText = ((Build)NewVersion.SelectedItem).Version;
            }
            else
            {
                selVersionText = NewVersion.Text;
            }
            return selVersionText;
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
            var modelBranch = GlobalObjects.ViewModel.Branches.FirstOrDefault(b =>
                b.Branch.ToString().ToLower() == branch.Branch.ToString().ToLower());
            if (modelBranch == null) return;

            NewVersion.SetValue(TextBoxHelper.WatermarkProperty, modelBranch.CurrentVersion);
        }

        private bool TransitionProductTabs(TransitionTabDirection direction)
        {
            var currentIndex = MainTabControl.SelectedIndex;
            var tmpIndex = currentIndex;
            if (direction == TransitionTabDirection.Forward)
            {
                if (MainTabControl.SelectedIndex < MainTabControl.Items.Count - 1)
                {
                    do
                    {
                        tmpIndex++;
                        if (tmpIndex < MainTabControl.Items.Count)
                        {
                            var item = (TabItem)MainTabControl.Items[tmpIndex];
                            if (item == null || item.IsVisible) break;
                        }
                        else
                        {
                            return true;
                        }
                    } while (true);
                    MainTabControl.SelectedIndex = tmpIndex;
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
                    do
                    {
                        tmpIndex--;
                        if (tmpIndex > 0)
                        {
                            var item = (TabItem)MainTabControl.Items[tmpIndex];
                            if (item == null || item.IsVisible) break;
                        }
                        else
                        {
                            return true;
                        }
                    } while (true);
                    MainTabControl.SelectedIndex = tmpIndex;
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

        private void NewVersion_OnKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (LocalInstall == null) return;
                var selVersionText = GetSelectedVersion();
                if (selVersionText.IsValidVersion())
                {
                    ChangeChannel.IsEnabled = LocalInstall.Version != selVersionText;
                }
                else
                {
                    ChangeChannel.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void OPPInstalled_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic)sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LatestVInstall_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic)sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void ShowVersion_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!GlobalObjects.ViewModel.BlockNavigation)
                {
                    if (NewVersionRow.Visibility == Visibility.Visible)
                    {
                        NewVersionRow.Visibility = Visibility.Collapsed;
                        ShowVersion.Content = "Show version";
                        var selectedBranch = (OfficeBranch)ProductBranch.SelectedItem;
                        if (selectedBranch != null && LocalInstall != null)
                        {
                            ChangeChannel.IsEnabled = selectedBranch.NewName != LocalInstall.Channel;
                        }
                        ProductBranch.Focus();
                    }
                    else
                    {
                        NewVersionRow.Visibility = Visibility.Visible;
                        ShowVersion.Content = "Hide version";

                        var selVersionText = GetSelectedVersion();
                        if (selVersionText.IsValidVersion())
                        {
                            ChangeChannel.IsEnabled = LocalInstall.Version != selVersionText;
                        }
                        else
                        {
                            ChangeChannel.IsEnabled = false;
                        }

                        ProductBranch.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void UnInstallOffice_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                GlobalObjects.ViewModel.BlockNavigation = true;

                var result = await MainWindow.ShowMessageAsync("Uninstall Office 365 ProPlus", "Confirm: Completely Uninstall Office 365 ProPlus from this computer?", MessageDialogStyle.AffirmativeAndNegative);
                if (result == MessageDialogResult.Affirmative)
                {
                    await UninstallOffice();
                }
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
                if (LocalInstall.Version.StartsWith("15."))
                {
                    //If 2013 rut this
                    await RunInstallOffice();
                }
                else
                {
                    //If 2016 rut this
                    await RunUpdateOffice();
                }   
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void ProductBranch_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var selectedBranch = (OfficeBranch) ProductBranch.SelectedItem;
                if (selectedBranch != null && LocalInstall != null)
                {
                    if (NewVersionRow.Visibility == Visibility.Visible)
                    {
                        var selVersionText = GetSelectedVersion();
                        if (selVersionText.IsValidVersion())
                        {
                            ChangeChannel.IsEnabled = LocalInstall.Version != selVersionText;
                        }
                        else
                        {
                            ChangeChannel.IsEnabled = false;
                        }
                    }
                    else
                    {
                        ChangeChannel.IsEnabled = selectedBranch.NewName != LocalInstall.Channel;
                    }
                }

                await UpdateVersions();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void NewVersion_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (LocalInstall == null) return;
                var selVersionText = GetSelectedVersion();
                if (selVersionText.IsValidVersion())
                {
                    ChangeChannel.IsEnabled = LocalInstall.Version != selVersionText;
                }
                else
                {
                    ChangeChannel.IsEnabled = false;
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
                await RunInstallOffice();
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

        private async void ReRunInstallOffice_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                await Task.Run(async () =>
                {
                    try
                    {
                        Dispatcher.Invoke(() =>
                        {
                            InstallOffice.IsEnabled = false;
                            ReInstallOffice.IsEnabled = false;
                            NewVersionRow.Visibility = Visibility.Collapsed;
                            ChangeChannel.IsEnabled = false;
                            ShowVersion.Content = "Show version";
                        });
                        GlobalObjects.ViewModel.BlockNavigation = true;
                        GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml.Display.Level = DisplayLevel.Full;

                        FirstRun = false;
                        
                        SetItemState(LocalViewItem.Install, LocalViewState.Wait);

                        var installGenerator = new OfficeInstallExecutableGenerator();
                        installGenerator.InstallOffice(GlobalObjects.ViewModel.ConfigXmlParser.Xml);

                        await LoadViewState();

                        Dispatcher.Invoke(() =>
                        {
                            InstallOffice.IsEnabled = true;
                            ReInstallOffice.IsEnabled = true;
                        });
                    }
                    catch (Exception ex)
                    {
                        SetItemState(LocalViewItem.Install, LocalViewState.Fail);
                        LogErrorMessage(ex);
                    }
                    finally
                    {
                        GlobalObjects.ViewModel.BlockNavigation = false;
                    }
                });
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

        private void xmlBrowser_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void InstalledVersion_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic)sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void UpdateChannel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic)sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void TargetVersionInfo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic)sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void ModifyExisting_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic)sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void OfficeInstall_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic)sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

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
        Update = 1,
        Uninstall = 2
    }

    public enum LocalViewState
    {
        Default = 0,
        Success = 1,
        Fail = 2,
        Action = 3,
        Wait = 5,
        Running = 6,
        InstallingOffice=7
    }

}

