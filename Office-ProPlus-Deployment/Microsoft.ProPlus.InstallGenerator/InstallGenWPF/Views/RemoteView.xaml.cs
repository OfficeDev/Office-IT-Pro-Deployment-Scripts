using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using MahApps.Metro.Controls;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Implementation;
using UserControl = System.Windows.Controls.UserControl;

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

                GlobalObjects.ViewModel.RemoteMachines = new List<RemoteMachine>();

                var connectionInfo = txtBxAddMachines.Text.Split('\\');
                var installGenerator = new OfficeInstallManager(connectionInfo);


                await Task.Run(async () => { await installGenerator.initConnections(); });

                var officeInstall = await installGenerator.CheckForOfficeInstallAsync();

                var versions = new List<String>();
                var channels = "";
                

                //set UI channel/version 

                var info = new RemoteMachine();

                if(officeInstall.Channel != null)
                {
                    var branches = GlobalObjects.ViewModel.Branches;

                    versions.Add(officeInstall.Version);
                    channels = channels + officeInstall.Channel;

                    //versionsComboBox.Width = 30;
                    //channelsComboBox.Width = 100;

                    //versionsComboBox.Height = 20;
                    //channelsComboBox.Height = 20;

                   
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

                    info = new RemoteMachine { include = false, Machine = txtBxAddMachines.Text, Status = "Found",
                    Channels = new List<Channel>()
                    {
                       new Channel()
                       {
                           Name = "Current"
                       }    
                    } 
                    , Channel = new Channel()
                    {
                        Name = "Current"
                    }, Version = versions };
                }
                else
                {
                    info = new RemoteMachine { include = false, Machine = txtBxAddMachines.Text, Status = "Not Found",
                    Channels = new List<Channel>()
                    {
                       new Channel()
                       {
                           Name = "Current"
                       }
                    }
                    ,Channel = new Channel()
                    {
                        Name = "Current"
                    }, Version = versions };
                }

                GlobalObjects.ViewModel.RemoteMachines.Add(info);

                txtBxAddMachines.Clear();

                RemoteMachineList.ItemsSource = GlobalObjects.ViewModel.RemoteMachines;

                WaitImage.Visibility = Visibility.Hidden;

            }
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
                        var info = new RemoteMachine { include = false, Machine = line, Status = "Found", Channel = null, Version = versions };
                        remoteClients.Add(info);
                    }
                    else
                    {
                        string[] tempStrArray = line.Split(',');
                        foreach (string tempStr in tempStrArray)
                        {
                            var info = new RemoteMachine { include = false, Machine = tempStr, Status = "Found", Channel = null, Version = versions };
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
            var officeInstall = await installGenerator.CheckForOfficeInstallAsync();



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

