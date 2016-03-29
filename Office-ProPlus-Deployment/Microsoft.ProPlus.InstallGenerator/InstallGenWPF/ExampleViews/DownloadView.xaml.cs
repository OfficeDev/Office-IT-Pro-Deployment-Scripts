using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using MetroDemo.Models;
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

        private List<Channel> items = null; 

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

            items = new List<Channel>
            {
                new Channel() {Name = "Current", ChannelName = "Current" },
                new Channel() {Name = "Deferred", ChannelName = "Deferred" },
                new Channel() {Name = "First Release Deferred", ChannelName = "FirstReleaseDeferred" },
                new Channel() {Name = "First Release Current", ChannelName = "FirstReleaseCurrent" }
            };

            if (configXml.Add.Branch.HasValue)
            {
                var selectedChannel = items.FirstOrDefault(c => c.ChannelName == configXml.Add.Branch.Value.ToString());
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
            catch { }
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

                var channelItems = (List<Channel>)lvUsers.ItemsSource;

                foreach (var channelItem in channelItems)
                {
                    if (!channelItem.Selected) continue;
                    var branch = channelItem.ChannelName;

                    var proPlusDownloader = new ProPlusDownloader();
                    proPlusDownloader.DownloadFileProgress += async (senderfp, progress) =>
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
                                lvUsers.ItemsSource = newList;

                                //DownloadPercent.Content = percent + "%";
                                //DownloadProgressBar.Value = Convert.ToInt32(Math.Round(percent, 0));
                            });
                        }
                    };
                    
                    proPlusDownloader.VersionDetected += (sender, version) =>
                    {
                        if (branch == null) return;
                        var modelBranch =
                            GlobalObjects.ViewModel.Branches.FirstOrDefault(
                                b => b.Branch.ToString().ToLower() == branch.ToLower());
                        if (modelBranch == null) return;
                        if (modelBranch.Versions.Any(v => v.Version == version.Version)) return;
                        modelBranch.Versions.Insert(0, new Build() {Version = version.Version});
                        modelBranch.CurrentVersion = version.Version;

                        //ProductVersion.ItemsSource = modelBranch.Versions;
                        //ProductVersion.SetValue(TextBoxHelper.WatermarkProperty, modelBranch.CurrentVersion);
                    };


                    if (string.IsNullOrEmpty(startPath)) return;

                    var languages =
                        (from product in configXml.Add.Products
                            from language in product.Languages
                            select language.ID.ToLower()).Distinct().ToList();

                    var officeEdition = OfficeEdition.Office32Bit;

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

                    var buildPath = GlobalObjects.SetBranchFolderPath(branch, startPath);
                    Directory.CreateDirectory(buildPath);

                    await proPlusDownloader.DownloadBranch(new DownloadBranchProperties()
                    {
                        BranchName = branch,
                        OfficeEdition = officeEdition,
                        TargetDirectory = buildPath,
                        Languages = languages
                    }, _tokenSource.Token);


                    var newTmpList = channelItems.ToList();
                    var tempItem2 = newTmpList.FirstOrDefault(c => c.Name == channelItem.Name);
                    if (tempItem2 != null)
                    {
                        tempItem2.PercentDownload = 100.00;
                    }
                    lvUsers.ItemsSource = newTmpList;

                    LogAnaylytics("/ProductView", "Download." + branch);
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


        public void Reset()
        {
            //ProductVersion.Text = "";
            ProductUpdateSource.Text = "";
            GlobalObjects.ViewModel.ClearLanguages();
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
            catch { }

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
                modelBranch.Versions.Insert(0, new Build() { Version = latestVersion });
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

        #region "Events"

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

    public class Channel
    {
        public string Name { get; set; }

        public string ChannelName { get; set; }

        public string Version { get; set; }

        public bool Selected { get; set; }

        public bool Editable { get; set; }

        public double PercentDownload { get; set; }

        public string PercentDownloadText { get; set; }

    }

    public class NegateConverter : IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool)
            {
                return !(bool)value;
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool)
            {
                return !(bool)value;
            }
            return value;
        }

    }
}

