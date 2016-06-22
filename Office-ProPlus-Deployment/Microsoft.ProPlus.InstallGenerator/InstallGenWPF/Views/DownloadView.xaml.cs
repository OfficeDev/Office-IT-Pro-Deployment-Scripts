using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using OfficeInstallGenerator.Model;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace MetroDemo.ExampleViews
{
    /// <summary>V
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class DownloadView : UserControl
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

        public DownloadView()
        {
            InitializeComponent();
        }

        private void DownloadView_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (MainTabControl == null) return;
                MainTabControl.SelectedIndex = 0;

                if (GlobalObjects.ViewModel == null) return;

                GlobalObjects.ViewModel.PropertyChangeEventEnabled = false;
                LoadXml();
                GlobalObjects.ViewModel.PropertyChangeEventEnabled = true;

                LoadChannelView();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LoadChannelView()
        {
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;

            if (ProductUpdateSource.Text.Length == 0)
            {
                if (!string.IsNullOrEmpty(configXml.Add.SourcePath))
                {
                    ProductUpdateSource.Text = configXml.Add.SourcePath;
                }
            }

            if (!string.IsNullOrEmpty(GlobalObjects.ViewModel.DownloadFolderPath))
            {
                ProductUpdateSource.Text = GlobalObjects.ViewModel.DownloadFolderPath;
            }

            var currentBranch = GlobalObjects.ViewModel.Branches.FirstOrDefault(b => b.NewName.ToLower() == "Current".ToLower());
            var deferredBranch = GlobalObjects.ViewModel.Branches.FirstOrDefault(b => b.NewName.ToLower() == "Deferred".ToLower());
            var firstReleaseDeferred = GlobalObjects.ViewModel.Branches.FirstOrDefault(b => b.NewName.ToLower() == "FirstReleaseDeferred".ToLower());
            var firstReleaseCurrent = GlobalObjects.ViewModel.Branches.FirstOrDefault(b => b.NewName.ToLower() == "FirstReleaseCurrent".ToLower());
            if (currentBranch == null) currentBranch = new OfficeBranch();
            if (deferredBranch == null) deferredBranch = new OfficeBranch();
            if (firstReleaseDeferred == null) firstReleaseDeferred = new OfficeBranch();
            if (firstReleaseCurrent == null) firstReleaseCurrent = new OfficeBranch();

            items = new List<Channel>
            {
                new Channel()
                {
                    Name = "Current",
                    ChannelName = "Current",
                    Version = "Latest",
                    Builds = currentBranch.Versions,
                    ForeGround = "Gray",
                },
                new Channel() {
                    Name = "Deferred", 
                    ChannelName = "Deferred", 
                    Version = "Latest", 
                    ForeGround = "Gray",
                    Builds = deferredBranch.Versions
                },
                new Channel()
                {
                    Name = "First Release Deferred",
                    ChannelName = "FirstReleaseDeferred",
                    Version = "Latest",
                    ForeGround = "Gray",
                    Builds = firstReleaseDeferred.Versions
                },
                new Channel()
                {
                    Name = "First Release Current",
                    ChannelName = "FirstReleaseCurrent",
                    Version = "Latest",
                    ForeGround = "Gray",
                    Builds = firstReleaseCurrent.Versions
                }
            };

            if (configXml.Add.Branch.HasValue)
            {
                var branchName = configXml.Add.Branch.Value.ToString();
                if (branchName.ToLower() == "Business".ToLower()) branchName = "Deferred";
                if (branchName.ToLower() == "FirstReleaseBusiness".ToLower()) branchName = "FirstReleaseDeferred";

                var selectedChannel = items.FirstOrDefault(c => c.ChannelName == branchName);
                if (selectedChannel != null)
                {
                    selectedChannel.Editable = true;
                    selectedChannel.Selected = true;
                }
            }

            if (configXml.Add.ODTChannel.HasValue)
            {
                var channelName = configXml.Add.ODTChannel.Value.ToString();
                if (channelName.ToLower() == "Deferred".ToLower()) channelName = "Deferred";
                if (channelName.ToLower() == "FirstReleaseDeferred".ToLower()) channelName = "FirstReleaseDeferred";

                var selectedChannel = items.FirstOrDefault(c => c.ChannelName == channelName);
                if (selectedChannel != null)
                {
                    selectedChannel.Editable = true;
                    selectedChannel.Selected = true;
                }
            }

            lvUsers.ItemsSource = items;

            if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office32Bit)
            {
                Download32Bit.IsChecked = true;
                Download64Bit.IsChecked = false;
                Download32Bit.IsEnabled = false;
                Download64Bit.IsEnabled = true;
            }
            else
            {
                Download64Bit.IsChecked = true;
                Download32Bit.IsChecked = false;
                Download32Bit.IsEnabled = true;
                Download64Bit.IsEnabled = false;
            }


        }

        private void LogAnaylytics(string path, string pageName)
        {
            try
            {
                GoogleAnalytics.Log(path, pageName);
            }
            catch
            {
            }
        }

        private async Task DownloadOfficeFiles()
        {
            try
            {
                SetTabStatus(false);
                GlobalObjects.ViewModel.BlockNavigation = true;
                _tokenSource = new CancellationTokenSource();

                UpdateXml();

                ProductUpdateSource.IsReadOnly = true;
                UpdatePath.IsEnabled = false;

                DownloadProgressBar.Maximum = 100;
                DownloadPercent.Content = "";

                var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
                var startPath = ProductUpdateSource.Text.Trim();

                if (!string.IsNullOrEmpty(startPath))
                {
                    GlobalObjects.ViewModel.DownloadFolderPath = startPath;
                }

                var channelItems = (List<Channel>) lvUsers.ItemsSource;

                var taskList = new List<Task>();
                _lastUpdated = DateTime.Now.AddDays(-10);
                var startTime = DateTime.Now;

                foreach (var channelItem in channelItems)
                {
                    if (!channelItem.Selected) continue;
                    var branch = channelItem.ChannelName;

                    var task = Task.Run(async () =>
                    {
                        try
                        {
                            var proPlusDownloader = new ProPlusDownloader();
                            proPlusDownloader.DownloadFileProgress += (senderfp, progress) =>
                            {
                                if (!_tokenSource.Token.IsCancellationRequested)
                                {
                                    DownloadFileProgress(progress, channelItems, channelItem);
                                }
                            };

                            proPlusDownloader.VersionDetected += (sender, version) =>
                            {
                                UpdateVersion(channelItems, channelItem, version.Version);
                            };

                            if (string.IsNullOrEmpty(startPath)) return;

                            var languages =
                                (from product in configXml.Add.Products
                                    from language in product.Languages
                                    select language.ID.ToLower()).Distinct().ToList();

                            var officeEdition = GetSelectedEdition();

                            var buildPath = GlobalObjects.SetBranchFolderPath(branch, startPath);

                            Directory.CreateDirectory(buildPath);

                            var setVersion = channelItem.DisplayVersion;

                            await proPlusDownloader.DownloadBranch(new DownloadBranchProperties()
                            {
                                BranchName = branch,
                                OfficeEdition = officeEdition,
                                TargetDirectory = buildPath,
                                Languages = languages,
                                Version = setVersion
                            }, _tokenSource.Token);

                            if (!_tokenSource.Token.IsCancellationRequested)
                            {
                                UpdatePercentage(channelItems, channelItem.Name);
                            }

                            LogAnaylytics("/ProductView", "Download." + branch);
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.ToLower().Contains("aborted"))
                            {
                                ex.LogException(false);
                                UpdateError(channelItems, channelItem.Name, "ERROR: " + ex.Message);
                            }
                        }
                    });

                    if (!GlobalObjects.ViewModel.AllowMultipleDownloads)
                    {
                        await task;
                    }

                    if (_tokenSource.Token.IsCancellationRequested) break;

                    var timeTaken = DateTime.Now - startTime;

                    taskList.Add(task);
                    await Task.Delay(new TimeSpan(0, 0, 5));
                }

                await Task.Delay(new TimeSpan(0, 0, 1));

                foreach (var task in taskList)
                {
                    if (task.Exception != null)
                    {

                    }
                    await task;
                }

                //MessageBox.Show("Download Complete");
            }
            finally
            {
                SetTabStatus(true);
                GlobalObjects.ViewModel.BlockNavigation = false;
                ProductUpdateSource.IsReadOnly = false;
                UpdatePath.IsEnabled = true;
                DownloadProgressBar.Value = 0;
                DownloadPercent.Content = "";

                DownloadButton.Content = "Download";
                _tokenSource = new CancellationTokenSource();
            }
        }

        private void UpdateVersion(IEnumerable<Channel> channelItems, Channel channelItem, string version)
        {
            if (channelItem.ChannelName == null) return;
            var modelBranch =
                GlobalObjects.ViewModel.Branches.FirstOrDefault(
                    b => b.NewName.ToString().ToLower() == channelItem.ChannelName.ToLower());
            if (modelBranch == null) return;

            ChangeVersion(channelItems, channelItem.Name, version);

            if (modelBranch.Versions.Any(v => v.Version == version)) return;
            modelBranch.Versions.Insert(0, new Build() {Version = version});
            modelBranch.CurrentVersion = version;
        }

        private void DownloadFileProgress(Microsoft.OfficeProPlus.Downloader.Events.DownloadFileProgress progress,
            IEnumerable<Channel> channelItems, Channel channelItem)
        {
            var percent = progress.PercentageComplete;
            if (percent > 0)
            {
                Dispatcher.Invoke(() =>
                {
                    channelItem.PercentDownload = percent;

                    var newList = channelItems.ToList();
                    var tempItem = newList.FirstOrDefault(c => c.Name == channelItem.Name);
                    if (tempItem != null)
                    {
                        tempItem.PercentDownload = percent;
                        tempItem.PercentDownloadText = percent + "%";
                    }

                    if (_lastUpdated < DateTime.Now.AddSeconds(-2))
                    {
                        lvUsers.ItemsSource = newList;
                        _lastUpdated = DateTime.Now;
                    }
                });
            }
        }

        private void ChangeVersion(IEnumerable<Channel> channelItems, string channelName, string version)
        {
            var newList = channelItems.ToList();
            var tempItem = newList.FirstOrDefault(c => c.Name == channelName);
            if (tempItem != null)
            {
                tempItem.DisplayVersion = version;
            }
            Dispatcher.Invoke(() =>
            {
                lvUsers.ItemsSource = newList;
            });
        }

        private void UpdateError(IEnumerable<Channel> channelItems, string channelName, string error)
        {
            var newTmpList = channelItems.ToList();
            var tempItem2 = newTmpList.FirstOrDefault(c => c.Name == channelName);
            if (tempItem2 != null)
            {
                tempItem2.PercentDownloadText = error;
                Dispatcher.Invoke(() =>
                {
                    lvUsers.ItemsSource = newTmpList;
                });
            }
        }

        private void UpdatePercentage(IEnumerable<Channel> channelItems, string channelName)
        {
            var newTmpList = channelItems.ToList();
            var tempItem2 = newTmpList.FirstOrDefault(c => c.Name == channelName);
            if (tempItem2 != null)
            {
                tempItem2.PercentDownload = 100.00;
                tempItem2.PercentDownloadText = "100%";
                Dispatcher.Invoke(() =>
                {
                    lvUsers.ItemsSource = newTmpList;
                });
            }
        }

        private OfficeEdition GetSelectedEdition()
        {
            var officeEdition = OfficeEdition.Office32Bit;

            Dispatcher.Invoke(() =>
            {
                if ((Download32Bit.IsChecked.HasValue && Download32Bit.IsChecked.Value) &&
                    (Download64Bit.IsChecked.HasValue && Download64Bit.IsChecked.Value))
                {
                    officeEdition = OfficeEdition.Both;
                }
                else
                {
                    if (Download32Bit.IsChecked.HasValue && Download32Bit.IsChecked.Value)
                    {
                        officeEdition = OfficeEdition.Office32Bit;
                    }
                    else
                    {
                        officeEdition = OfficeEdition.Office64Bit;
                    }
                }
            });

            return officeEdition;
        }


        public void Reset()
        {
            //ProductVersion.Text = "";
            ProductUpdateSource.Text = "";
        }

        public void LoadXml()
        {
            Reset();

            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Add != null)
            {
                //ProductVersion.Text = configXml.Add.Version != null ? configXml.Add.Version.ToString() : "";
                ProductUpdateSource.Text = configXml.Add.SourcePath != null ? configXml.Add.SourcePath.ToString() : "";
            }
            else
            {
                //ProductVersion.Text = "";
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
            catch
            {
            }

            configXml.Add.SourcePath = ProductUpdateSource.Text.Length > 0 ? ProductUpdateSource.Text : null;

        }



        private async Task GetBranchVersion(OfficeBranch branch, OfficeEdition officeEdition)
        {
            try
            {
                if (branch.Updated) return;
                var ppDownload = new ProPlusDownloader();
                var latestVersion = await ppDownload.GetLatestVersionAsync(branch.Branch.ToString(), officeEdition);

                var modelBranch = GlobalObjects.ViewModel.Branches.FirstOrDefault(b =>
                    b.Branch.ToString().ToLower() == branch.Branch.ToString().ToLower());
                if (modelBranch == null) return;
                if (modelBranch.Versions.Any(v => v.Version == latestVersion)) return;
                modelBranch.Versions.Insert(0, new Build() {Version = latestVersion});
                modelBranch.CurrentVersion = latestVersion;

                //ProductVersion.ItemsSource = modelBranch.Versions;
                //ProductVersion.SetValue(TextBoxHelper.WatermarkProperty, modelBranch.CurrentVersion);

                modelBranch.Updated = true;
            }
            catch (Exception)
            {

            }
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

        #region "Events"

        private bool allowCheck = true;
        private void ToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!allowCheck) return;
                var chkBox = (System.Windows.Controls.CheckBox) sender;
                if (GlobalObjects.ViewModel.BlockNavigation)
                {
                    allowCheck = false;
                    chkBox.IsChecked = !chkBox.IsChecked;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
            finally
            {
                allowCheck = true;
            }
        }

        private void AdvDownloadButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (advancedSettings == null)
                {
                    advancedSettings = new DownloadAdvanced();
                }

                advancedSettings.ShowDialog();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void DownloadButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_tokenSource != null)
                {
                    if (_tokenSource.IsCancellationRequested)
                    {
                        GlobalObjects.ViewModel.BlockNavigation = false;
                        SetTabStatus(true);
                        return;
                    }
                    if (_downloadTask.IsActive())
                    {
                        GlobalObjects.ViewModel.BlockNavigation = false;
                        SetTabStatus(true);
                        _tokenSource.Cancel();
                        return;
                    }
                }

                DownloadButton.Content = "Stop";

                _downloadTask = DownloadOfficeFiles();
                await _downloadTask;
            }
            catch (Exception ex)
            {
                if (ex.Message.ToLower().Contains("aborted") ||
                    ex.Message.ToLower().Contains("canceled"))
                {
                    GlobalObjects.ViewModel.BlockNavigation = false;
                    SetTabStatus(true);
                }
                else
                {
                    LogErrorMessage(ex);
                }
            }
        }

        private async void OpenFolderButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var folderPath = ProductUpdateSource.Text.Trim();
                if (string.IsNullOrEmpty(folderPath)) return;

                if (await GlobalObjects.DirectoryExists(folderPath))
                {
                    Process.Start("explorer", folderPath);
                }
                else
                {
                    MessageBox.Show("Directory path does not exist.");
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }


        private async void BuildFilePath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                var enabled = false;
                var openFolderEnabled = false;
                if (ProductUpdateSource.Text.Trim().Length > 0)
                {
                    var match = Regex.Match(ProductUpdateSource.Text, @"^\w:\\|\\\\.*\\..*");
                    if (match.Success)
                    {
                        enabled = true;
                        var folderExists = await GlobalObjects.DirectoryExists(ProductUpdateSource.Text);
                        if (!folderExists)
                        {
                            folderExists = await GlobalObjects.DirectoryExists(ProductUpdateSource.Text);
                        }

                        openFolderEnabled = folderExists;
                    }
                }

                OpenFolderButton.IsEnabled = openFolderEnabled;
                DownloadButton.IsEnabled = enabled;
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
                var dlg1 = new Ionic.Utils.FolderBrowserDialogEx
                {
                    Description = "Select a folder:",
                    ShowNewFolderButton = true,
                    ShowEditBox = true,
                    SelectedPath = ProductUpdateSource.Text,
                    ShowFullPathInEditBox = true,
                    RootFolder = System.Environment.SpecialFolder.MyComputer
                };
                //dlg1.NewStyle = false;

                // Show the FolderBrowserDialog.
                var result = dlg1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    ProductUpdateSource.Text = dlg1.SelectedPath;
                }
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
