using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class UpdateView : UserControl
    {

        public UpdateView()
        {
            InitializeComponent();
        }

        private void ToggleControls(bool enabled)
        {

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

            if (Version.TryParse(txtTargetVersion, out targetVersion))
            {
                configXml.Updates.TargetVersion = targetVersion;
            }

            var xml = GlobalObjects.ViewModel.ConfigXmlParser.Xml;
            if (xml != null)
            {

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

            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Add == null) return;

            var officeEdition = OfficeEdition.Office32Bit;
            if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office64Bit)
            {
                officeEdition = OfficeEdition.Office64Bit;
            }

            await GetBranchVersion(branch, officeEdition);
        }

        public event TransitionTabEventHandler TransitionTab;

        private void UpdatePath_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var openDialog = new OpenFileDialog
                {
                    Filter = "v32.cab File|v32.cab|v64.cab File|v64.cab",
                    Multiselect = false
                };

                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    var filePath = openDialog.FileName;
                    filePath = Regex.Replace(filePath, @"\\Office\\Data\\v32.cab", "", RegexOptions.IgnoreCase);
                    filePath = Regex.Replace(filePath, @"\\Office\\Data\\v64.cab", "", RegexOptions.IgnoreCase);

                    UpdateUpdatePath.Text = filePath;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
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
                MessageBox.Show("ERROR: " + ex.Message);
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
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        #region "Events"

        private async void UpdateBranch_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                await UpdateVersions();
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
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

                var filePath = AppDomain.CurrentDomain.BaseDirectory + @"HelpFiles\" + sourceName + ".html";
                var helpFile = File.ReadAllText(filePath);

                informationDialog.HelpInfo.NavigateToString(helpFile);
                informationDialog.Launch();

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
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
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        #endregion




    }
}
