using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace MetroDemo.ExampleWindows
{
    /// <summary>
    /// Interaction logic for RemoteChannelVersionDialog.xaml
    /// </summary>
    public partial class RemoteChannelVersionDialog : IDisposable
    {
        public RemoteChannelVersionDialog()
        {
            InitializeComponent();
        }
        public DialogResult Result = System.Windows.Forms.DialogResult.Cancel;
        private List<Channel> items;

        public void Launch()
        {
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

            Owner = System.Windows.Application.Current.MainWindow;
            ChannelSelection.Items.Add("Deferred");
            ChannelSelection.Items.Add("Current");
            ChannelSelection.Items.Add("FirstReleaseDeferred");
            ChannelSelection.Items.Add("FirstReleaseCurrent");

            // only for this window, because we allow minimizing
            if (WindowState == WindowState.Minimized)
            {
                WindowState = WindowState.Normal;
            }
            Show();
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
               
                Result = System.Windows.Forms.DialogResult.OK;
                if(ChannelSelection.SelectedItem != null  && VersionSelection.SelectedItem != null)
                {
                    GlobalObjects.ViewModel.newChannel = ChannelSelection.SelectedItem.ToString();
                    GlobalObjects.ViewModel.newVersion = VersionSelection.SelectedItem.ToString();
                }
               
                this.Close();                
            }
            catch (Exception ex)
            {
                ex.LogException();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Result = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void ChannelSelection_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedBranch = (sender as System.Windows.Controls.ComboBox).SelectedValue as string;
            var newVersions = new List<String>();
            var branches = GlobalObjects.ViewModel.Branches;

            foreach (var branch in branches)
            {
                if (branch.NewName.ToString() == selectedBranch)
                {
                    newVersions = getVersions(branch, newVersions, "");
                    break;
                }
            }



            VersionSelection.ItemsSource = newVersions;
            VersionSelection.SelectedItem = newVersions[0];
            VersionSelection.Items.Refresh();
        }

        private List<String> getVersions(OfficeBranch currentChannel, List<String> versions, string currentVersion)
        {

            foreach (var version in currentChannel.Versions)
            {
                if (version.Version.ToString() != currentVersion)
                {
                    versions.Add(version.Version.ToString());
                }
            }

            return versions;
        }

        //private void ChannelSelection_DropDownClosed(object sender, EventArgs e)
        //{
        //    switch (ChannelSelection.SelectedValue.ToString())
        //    {
        //        case "Current":                    
        //            VersionSelection.Items.Clear();
        //            VersionSelection.Items.Add(items[0].Version);
        //            foreach (var build in items[0].Builds)
        //                VersionSelection.Items.Add(build.Version);
        //            VersionSelection.SelectedValue = items[0].Version;
        //            break;
        //        case "Deferred":
        //            VersionSelection.Items.Clear();
        //            VersionSelection.Items.Add(items[1].Version);
        //            foreach (var build in items[1].Builds)
        //                VersionSelection.Items.Add(build.Version);
        //            VersionSelection.SelectedValue = items[1].Version;
        //            break;
        //        case "FirstReleaseDeferred":
        //            VersionSelection.Items.Clear();
        //            VersionSelection.Items.Add(items[2].Version);
        //            foreach (var build in items[2].Builds)
        //                VersionSelection.Items.Add(build.Version);
        //            VersionSelection.SelectedValue = items[2].Version;
        //            break;
        //        case "FirstReleaseCurrent":
        //            VersionSelection.Items.Clear();
        //            VersionSelection.Items.Add(items[3].Version);
        //            foreach (var build in items[3].Builds)
        //                VersionSelection.Items.Add(build.Version);
        //            VersionSelection.SelectedValue = items[3].Version;
        //            break;
        //    }
        //}

    }
}
