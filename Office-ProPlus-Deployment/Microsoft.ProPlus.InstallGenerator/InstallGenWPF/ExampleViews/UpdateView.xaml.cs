﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
using MahApps.Metro.Converters;
using MetroDemo.Annotations;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using Microsoft.VisualBasic;
using File = System.IO.File;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class UpdateView : UserControl
    {

        private bool _updatePathChanged = false;
        private CancellationTokenSource _tokenSource = new CancellationTokenSource();
        private Task _downloadTask = null;

        public event TransitionTabEventHandler TransitionTab;
        public event MessageEventHandler InfoMessage;
        public event MessageEventHandler ErrorMessage;

        private string ddTimeHour = "00";
        private string ddTimeMinute = "00";


        public UpdateView()
        {
            InitializeComponent();
        }

        private void UpdateView_OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                GlobalObjects.ViewModel.PropertyChanged += ViewModel_PropertyChanged;

                LogAnaylytics("/UpdateView", "Load");
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        public void LoadXml()
        {
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Updates != null)
            {
                EnabledSwitch.IsChecked = configXml.Updates.Enabled;
                UpdateUpdatePath.Text = configXml.Updates.UpdatePath;

                if (configXml.Updates.TargetVersion != null)
                {
                    var targetVersion = configXml.Updates.TargetVersion.ToString();

                    var targetVersionIndex = -1;
                    for (var i = 0; i < UpdateTargetVersion.Items.Count; i++)
                    {
                        var item = (Build) UpdateTargetVersion.Items[i];
                        if (item.Version.ToLower() == targetVersion.ToLower())
                        {
                            targetVersionIndex = i;
                        }
                    }

                    UpdateTargetVersion.SelectedIndex = targetVersionIndex;
                    if (targetVersionIndex == -1)
                    {
                        UpdateTargetVersion.Text = targetVersion;
                    }
                }

                if (configXml.Updates.Deadline.HasValue)
                {
                    UpdateDeadline.SelectedDate = configXml.Updates.Deadline.Value;
                    DeadlineTimeHour.Text = configXml.Updates.Deadline.Value.Hour.ToString();
                    DeadlineTimeMinute.Text = configXml.Updates.Deadline.Value.Minute.ToString();
                }
            }
        }

        public void UpdateXml()
        {
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Updates == null)
            {
                configXml.Updates = new ODTUpdates();
            }

            var updatesEnabled = false;
            if (EnabledSwitch.IsChecked.HasValue)
            {
               updatesEnabled = EnabledSwitch.IsChecked.Value;
            }
           
            var updateBranch = (OfficeBranch) UpdateBranch.SelectedItem;
            var txtTargetVersion = UpdateTargetVersion.Text;
            Version targetVersion = null;

            if (updateBranch != null)
            {
                configXml.Updates.Branch = updateBranch.Branch;
            }

            configXml.Updates.Enabled = updatesEnabled;
            configXml.Updates.UpdatePath = UpdateUpdatePath.Text;

            configXml.Updates.TargetVersion = null;
            if (Version.TryParse(txtTargetVersion, out targetVersion))
            {
                configXml.Updates.TargetVersion = targetVersion;
            }

            configXml.Updates.Deadline = null;
            if (UpdateDeadline.SelectedDate.HasValue)
            {
                DateTime? deadLine = null;
                UpdateDeadline.SelectedDateFormat = DatePickerFormat.Short;

                var dl = UpdateDeadline.SelectedDate.Value;
                var hour = 0;
                var minute = 0;

                if (!string.IsNullOrEmpty(DeadlineTimeHour.Text) && 
                    !string.IsNullOrEmpty(DeadlineTimeMinute.Text))
                {
                    hour = Convert.ToInt32(DeadlineTimeHour.Text);
                    minute = Convert.ToInt32(DeadlineTimeMinute.Text);
                }

                deadLine = new DateTime(dl.Year, dl.Month, dl.Day, hour, minute, 0);

                configXml.Updates.Deadline = deadLine;
            }
        }

        public void Reset()
        {
            _updatePathChanged = false;
            UpdateBranch.SelectedIndex = 0;
            EnabledSwitch.IsChecked = false;
            UpdateTargetVersion.Text = "";
            UpdateUpdatePath.Text = "";
            UpdateDeadline.Text = "";
            DeadlineTimeHour.Text = "";
            DeadlineTimeMinute.Text = "";
        }


        private async Task DownloadOfficeFiles()
        {
            try
            {
                _tokenSource = new CancellationTokenSource();

                UpdateXml();

                UpdateUpdatePath.IsReadOnly = true;
                UpdatePath.IsEnabled = false;

                DownloadProgressBar.Maximum = 100;
                DownloadPercent.Content = "";

                var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;

                string branch = null;
                if (configXml.Updates.Branch.HasValue)
                {
                    branch = configXml.Updates.Branch.Value.ToString();
                }

                var proPlusDownloader = new ProPlusDownloader();
                proPlusDownloader.DownloadFileProgress += async (senderfp, progress) =>
                {
                    var percent = progress.PercentageComplete;
                    if (percent > 0)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            DownloadPercent.Content = percent + "%";
                            DownloadProgressBar.Value = Convert.ToInt32(Math.Round(percent, 0));
                        });
                    }
                };
                proPlusDownloader.VersionDetected += (sender, version) =>
                {
                    if (branch == null) return;
                    var modelBranch = GlobalObjects.ViewModel.Branches.FirstOrDefault(b => b.Branch.ToString().ToLower() == branch.ToLower());
                    if (modelBranch == null) return;
                    if (modelBranch.Versions.Any(v => v.Version == version.Version)) return;
                    modelBranch.Versions.Insert(0, new Build() { Version = version.Version });
                    modelBranch.CurrentVersion = version.Version;

                    UpdateTargetVersion.ItemsSource = modelBranch.Versions;
                    UpdateTargetVersion.SetValue(TextBoxHelper.WatermarkProperty, modelBranch.CurrentVersion);
                };

                var buildPath = UpdateUpdatePath.Text.Trim();
                if (string.IsNullOrEmpty(buildPath)) return;

                var languages =
                    (from product in configXml.Add.Products
                     from language in product.Languages
                     select language.ID.ToLower()).Distinct().ToList();

                var officeEdition = OfficeEdition.Office32Bit;
                if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office64Bit)
                {
                    officeEdition = OfficeEdition.Office64Bit;
                }

                buildPath = GlobalObjects.SetBranchFolderPath(branch, buildPath);
                Directory.CreateDirectory(buildPath);

                UpdateUpdatePath.Text = buildPath;

                await proPlusDownloader.DownloadBranch(new DownloadBranchProperties()
                {
                    BranchName = branch,
                    OfficeEdition = officeEdition,
                    TargetDirectory = buildPath,
                    Languages = languages
                }, _tokenSource.Token);

                MessageBox.Show("Download Complete");
            }
            finally
            {
                UpdateUpdatePath.IsReadOnly = false;
                UpdatePath.IsEnabled = true;
                DownloadProgressBar.Value = 0;
                DownloadPercent.Content = "";

                DownloadButton.Content = "Download";
                _tokenSource = new CancellationTokenSource();
            }
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

                UpdateTargetVersion.ItemsSource = modelBranch.Versions;
                UpdateTargetVersion.SetValue(TextBoxHelper.WatermarkProperty, modelBranch.CurrentVersion);

                modelBranch.Updated = true;
            }
            catch (Exception ex)
            {
                var strError = ex.Message;
                if (strError != null)
                {
                    
                }
            }
        }

        private async Task UpdateVersions()
        {
            var branch = (OfficeBranch)UpdateBranch.SelectedItem;
            if (branch == null) return;

            UpdateTargetVersion.ItemsSource = branch.Versions;
            UpdateTargetVersion.SetValue(TextBoxHelper.WatermarkProperty, branch.CurrentVersion);

            var officeEdition = OfficeEdition.Office32Bit;

            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Add != null)
            {
                if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office64Bit)
                {
                    officeEdition = OfficeEdition.Office64Bit;
                }
            }

            if (UpdateUpdatePath.Text.Length > 0)
            {
                var otherFolder = GlobalObjects.SetBranchFolderPath(branch.Branch.ToString(), UpdateUpdatePath.Text);
                if (await GlobalObjects.DirectoryExists(otherFolder))
                {
                    if (!string.IsNullOrEmpty(UpdateUpdatePath.Text))
                    {
                        UpdateUpdatePath.Text = GlobalObjects.SetBranchFolderPath(branch.Branch.ToString(),
                            UpdateUpdatePath.Text);
                    }
                }
            }

            await GetBranchVersion(branch, officeEdition);
        }

        private string AllowHours(string text)
        {
            var newHour = "";

            if (string.IsNullOrEmpty(text)) return newHour;
            if (!Microsoft.VisualBasic.Information.IsNumeric(text)) return ddTimeHour;
            var numHours = Convert.ToInt32(text);
            if (numHours > 23) newHour = "23";
            if (numHours < 0) newHour = "00";

            if (newHour.Length == 1)
            {
                newHour = "0" + newHour;
            }
            return newHour;
        }

        private string AllowMinute(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            if (!Microsoft.VisualBasic.Information.IsNumeric(text)) return ddTimeMinute;
            var numHours = Convert.ToInt32(text);
            if (numHours > 59) return "59";
            if (numHours < 0) return "00";
            return "";
        }

        private void UpdatePath_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateXml();

                var dlg1 = new Ionic.Utils.FolderBrowserDialogEx
                {
                    Description = "Select a folder:",
                    ShowNewFolderButton = true,
                    ShowEditBox = true,
                    SelectedPath = UpdateUpdatePath.Text,
                    ShowFullPathInEditBox = true,
                    RootFolder = System.Environment.SpecialFolder.MyComputer
                };
                //dlg1.NewStyle = false;

                // Show the FolderBrowserDialog.
                var result = dlg1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    UpdateUpdatePath.Text = dlg1.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void PreviousButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateXml();

                this.TransitionTab(this, new TransitionTabEventArgs()
                {
                    Direction = TransitionTabDirection.Back
                });
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void NextButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateXml();

                this.TransitionTab(this, new TransitionTabEventArgs()
                {
                    Direction = TransitionTabDirection.Forward
                });
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
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

        private void LogAnaylytics(string path, string pageName)
        {
            try
            {
                GoogleAnalytics.Log(path, pageName);
            }
            catch { }
        }

        #region "Events"

        private void UpdateDeadline_OnPreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (UpdateDeadline.Text.Trim().Length == 0)
                {
                    DeadlineTimeHour.Text = "";
                    DeadlineTimeMinute.Text = "";
                }

            }
            catch { }
        }

        private void UpdateDeadline_OnSelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (UpdateDeadline.Text.Trim().Length > 0)
                {
                    if (DeadlineTimeHour.Text.Trim().Length == 0)
                    {
                        DeadlineTimeHour.Text = "00";    
                    }
                    if (DeadlineTimeMinute.Text.Trim().Length == 0)
                    {
                        DeadlineTimeMinute.Text = "00";
                    }
                }
            }
            catch { }
        }

        private void DeadlineTimeMinute_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                var ddMinute = DeadlineTimeMinute.Text;
                if (string.IsNullOrEmpty(ddMinute)) ddMinute = "00";
                if (Information.IsNumeric(ddMinute))
                {
                    var minute = Convert.ToInt32(ddMinute);
                    if (e.Key == Key.Up)
                    {
                        if (minute < 59)
                        {
                            var newMinute = (minute + 1).ToString();
                            if (newMinute.Length == 1)
                            {
                                newMinute = "0" + newMinute;
                            }

                            DeadlineTimeMinute.Text = newMinute;
                        }
                    }
                    if (e.Key == Key.Down)
                    {
                        if (minute > 0)
                        {
                            var newMinute = (minute - 1).ToString();
                            if (newMinute.Length == 1)
                            {
                                newMinute = "0" + newMinute;
                            }

                            DeadlineTimeMinute.Text = newMinute;
                        }
                    }
                }
            }
            catch { }
        }

        private void DeadlineTimeHour_OnKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                var ddHour = DeadlineTimeHour.Text;
                if (string.IsNullOrEmpty(ddHour)) ddHour = "00";
                if (Information.IsNumeric(ddHour))
                {
                    var hour = Convert.ToInt32(ddHour);
                    if (e.Key == Key.Up)
                    {
                        if (hour < 23)
                        {
                            var newHour = (hour + 1).ToString();
                            if (newHour.Length == 1)
                            {
                                newHour = "0" + newHour;
                            }

                            DeadlineTimeHour.Text = newHour;
                        }
                    }
                    if (e.Key == Key.Down)
                    {
                        if (hour > 0)
                        {
                            var newHour = (hour - 1).ToString();
                            if (newHour.Length == 1)
                            {
                                newHour = "0" + newHour;
                            }

                            DeadlineTimeHour.Text = newHour;
                        }
                    }
                }
            }
            catch { }
        }
        
        private void DeadlineTimeHour_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (DeadlineTimeHour.Text.Contains("."))
                {
                    DeadlineTimeHour.Text = DeadlineTimeHour.Text.Replace(".", "");
                }

                var hourConvert = AllowHours(DeadlineTimeHour.Text);
                if (string.IsNullOrEmpty(hourConvert)) return;

                var newHour = hourConvert.Replace(".", "");
                DeadlineTimeHour.Text = newHour;
                ddTimeHour = DeadlineTimeHour.Text;
            }
            catch { }
        }

        private void DeadlineTimeMinute_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (DeadlineTimeMinute.Text.Contains("."))
                {
                    DeadlineTimeMinute.Text = DeadlineTimeMinute.Text.Replace(".", "");
                }

                var minConvert = AllowMinute(DeadlineTimeMinute.Text);
                if (string.IsNullOrEmpty(minConvert)) return;

                DeadlineTimeMinute.Text = minConvert.Replace(".", "");
                ddTimeMinute = DeadlineTimeHour.Text;
            }
            catch { }
        }

        private async void DownloadButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_tokenSource != null)
                {
                    if (_tokenSource.IsCancellationRequested)
                    {
                        return;
                    }
                    if (_downloadTask.IsActive())
                    {
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
                var folderPath = UpdateUpdatePath.Text.Trim();
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

        private async void UpdateUpdatePath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                _updatePathChanged = true;

                var enabled = false;
                var openFolderEnabled = false;
                if (UpdateUpdatePath.Text.Trim().Length > 0)
                {
                    var match = Regex.Match(UpdateUpdatePath.Text, @"^\w:\\|\\\\.*\\..*");
                    if (match.Success)
                    {
                        enabled = true;
                        var folderExists = await GlobalObjects.DirectoryExists(UpdateUpdatePath.Text);
                        if (!folderExists)
                        {
                            folderExists = await GlobalObjects.DirectoryExists(UpdateUpdatePath.Text);
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

        private void ViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            try
            {
                switch (e.PropertyName.ToLower())
                {
                    case "updatepath":
                        if (_updatePathChanged) return;
                        UpdateUpdatePath.TextChanged -= UpdateUpdatePath_OnTextChanged;
                        UpdateUpdatePath.Text = GlobalObjects.ViewModel.UpdatePath;
                        UpdateUpdatePath.TextChanged += UpdateUpdatePath_OnTextChanged;
                        break;
                    case "selectedbranch":
                        for (var i = 0; i < UpdateBranch.Items.Count; i++)
                        {
                            var branch = (OfficeBranch) UpdateBranch.Items[i];
                            if (branch.Branch.ToString().ToLower() == GlobalObjects.ViewModel.SelectedBranch.ToLower())
                            {
                                UpdateBranch.SelectedIndex = i;
                                break;
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void UpdateBranch_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                await UpdateVersions();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void EnabledSwitch_OnIsCheckedChanged(object sender, EventArgs e)
        {
            try
            {
                UpdateBranch.IsEnabled = EnabledSwitch.IsChecked.HasValue && EnabledSwitch.IsChecked.Value;
                UpdateUpdatePath.IsEnabled = EnabledSwitch.IsChecked.HasValue && EnabledSwitch.IsChecked.Value;
                UpdateTargetVersion.IsEnabled = EnabledSwitch.IsChecked.HasValue && EnabledSwitch.IsChecked.Value;
                UpdateDeadline.IsEnabled = EnabledSwitch.IsChecked.HasValue && EnabledSwitch.IsChecked.Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message, "ERROR", MessageBoxButton.OK);
            }
        }

        #endregion

        #region "Info"

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
        
        private void ProductInfo_OnClick(object sender, RoutedEventArgs e)
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


    }
}
