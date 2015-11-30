using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MetroDemo.Events;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
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


    }
}
