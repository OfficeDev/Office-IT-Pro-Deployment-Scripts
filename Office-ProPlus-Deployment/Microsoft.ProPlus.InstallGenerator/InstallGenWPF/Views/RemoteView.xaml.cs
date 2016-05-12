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
using System.Data;

namespace MetroDemo.ExampleViews
{
    /// <summary>V
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class RemoteView : UserControl
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

        private OfficeInstall RemoteInstall { get; set; }
        private bool FirstRun = true;
        private List<RemoteMachine> remoteClients = new List<RemoteMachine>();
        #endregion

        public RemoteView()
        {            
            InitializeComponent();            
        }

        private void RemoteView_Loaded(object sender, RoutedEventArgs e)             
        {
            
            

            RemoteMachineList.ItemsSource = remoteClients;
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

        private async void AddComputersButton_Click(object sender, RoutedEventArgs e)
        {
            //placeholder text for data entry Username\Password\IP\Domain


            if (txtBxAddMachines.Text != "")
            {
                //parse text 

                WaitImage.Visibility = Visibility.Visible;


                var connectionInfo = txtBxAddMachines.Text.Split('\\');
                var installGenerator = new OfficeInstallManager(connectionInfo);


                await Task.Run(async () => { await installGenerator.initConnections(); });

                var officeInstall = await installGenerator.CheckForOfficeInstallAsync();

                List<string> versions = new List<String>();
                string channels = "";
                

                //set UI channel/version 

                var info = new RemoteMachine();
                if(officeInstall.Channel != null)
                {
                    var branches = GlobalObjects.ViewModel.Branches;

                    versions.Add(officeInstall.Version);
                    channels = channels + officeInstall.Channel;

                    foreach (var branch in branches)
                    {
                        if (branch.Branch.ToString() != officeInstall.Channel)
                        {
                            channels= channels + branch.Branch.ToString();
                        }

                        if (branch.Versions.ToString() != officeInstall.Version)
                        {
                            versions.Add(branch.CurrentVersion);
                        }
                    }


                    Channel.ItemsSource = channels;

                    info = new RemoteMachine { include = false, Machine = txtBxAddMachines.Text, Status = "Found", Channel = channels, Version = versions };


                }
                else
                {
                    info = new RemoteMachine { include = false, Machine = txtBxAddMachines.Text, Status = "Not Found", Channel = channels, Version = versions };
                }
                remoteClients.Add(info);
                txtBxAddMachines.Clear();                
                RemoteMachineList.Items.Refresh();

                WaitImage.Visibility = Visibility.Hidden;

            }
        }

        public class RemoteMachine
        {
            public bool include { get; set; }
            public string Machine { get; set; }
            public string Status { get; set; }
            public string Channel { get; set; }
            public List<string> Version { get; set; }
        }

        private RemoteChannelVersionDialog remoteUpdateDialog = null;
        private void btnChangeChannelOrVersion_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                remoteUpdateDialog = new RemoteChannelVersionDialog();



                remoteUpdateDialog.Closing += RemoteUpdateDialog_Closing;
                remoteUpdateDialog.Launch();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void RemoteUpdateDialog_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                var dialog = (RemoteChannelVersionDialog)sender;
                if (dialog.Result == DialogResult.OK)
                {
                    //Alex implementation here
                    
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void ImportComputersButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".png",
                Filter = "Text Files (.txt)|*.txt|CSV Files (.csv)|*.csv"
            };

            var result = dlg.ShowDialog();
            if (result == true)
            {

                List<string> versions = new List<String>();
                string channels = "";
                string line;
                StreamReader file = new StreamReader(dlg.FileName);
                while ((line = file.ReadLine()) != null)
                {
                    if (!line.Contains(","))
                    {
                        var info = new RemoteMachine { include = false, Machine = line, Status = "Found", Channel = channels, Version = versions };
                        remoteClients.Add(info);
                    }
                    else
                    {
                        string[] tempStrArray = line.Split(',');
                        foreach (string tempStr in tempStrArray)
                        {
                            var info = new RemoteMachine { include = false, Machine = tempStr, Status = "Found", Channel = channels, Version = versions };
                            remoteClients.Add(info);
                        }
                    }
                    
                }
                RemoteMachineList.Items.Refresh();
                
            }
        }

        private void chkAll_Click(object sender, RoutedEventArgs e)
        {
            foreach(var client in remoteClients)
            {
                client.include = chkAll.IsChecked.Value;
            }
            RemoteMachineList.Items.Refresh();
        }

        private async void btnUpdateRemote_Click(object sender, RoutedEventArgs e)
        {


            WaitImage.Visibility = Visibility.Visible;

            var connectionInfo = new string[4] { "Molly Clark", "pass@word1", "10.10.8.225", "WORKGROUP" };
            var installGenerator = new OfficeInstallManager(connectionInfo);



            await Task.Run(async () => { await installGenerator.initConnections(); });

            await Task.Run(() => { installGenerator.UpdateOffice();});

            //var officeInstall = await installGenerator.CheckForOfficeInstallAsync();

            



            WaitImage.Visibility = Visibility.Hidden;

        }

        //#endregion

    }


    public enum RemoteViewItem
    {
        Install = 0,
        Update = 1,
        Uninstall = 2
    }

    public enum RemoteViewState
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

