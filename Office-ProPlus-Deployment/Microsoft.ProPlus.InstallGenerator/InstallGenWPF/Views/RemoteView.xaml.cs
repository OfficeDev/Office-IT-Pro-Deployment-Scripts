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
using Microsoft.OfficeProPlus.InstallGen.Presentation.Extentions;
using Microsoft.OfficeProPlus.InstallGenerator.Model;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.Downloader;
using System.Windows.Media;
using UserControl = System.Windows.Controls.UserControl;
using System.Diagnostics;

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

            GlobalObjects.ViewModel.RemoteMachines = new List<RemoteMachine>();
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

        private async void addMachines(string[] connectionInfo)
        {
           
               var installGenerator = new OfficeInstallManager(connectionInfo);


               await installGenerator.initConnections();


                var officeInstall = await installGenerator.CheckForOfficeInstallAsync();

                var channels = new List<Channel>();

                var versions = new List<officeVersion>();


                //set UI channel/version 

                var info = new RemoteMachine();

                if (officeInstall.Channel != null)
                {
                    var branches = GlobalObjects.ViewModel.Branches;

                    var currentChannel = new Channel()
                    {
                        Name = officeInstall.Channel.Trim()
                    };

                    var currentVersion = new officeVersion()
                    {
                        Number = officeInstall.Version.Trim()
                    };

                    versions.Add(currentVersion);
                    channels.Add(currentChannel);




                    foreach (var branch in branches)
                    {
                        if (branch.Branch.ToString() != officeInstall.Channel)
                        {
                            var tempChannel = new Channel()
                            {
                                Name = branch.NewName.ToString()
                            };
                            channels.Add(tempChannel);
                        }
                        else
                        {

                            versions = getVersions(branch, versions, currentVersion.Number);
                        }


                    }

                    info = new RemoteMachine
                    {
                        include = false,
                        Machine = connectionInfo[2],
                        UserName = connectionInfo[0],
                        Password = connectionInfo[1],
                        WorkGroup = connectionInfo[3],
                        Status = "Found",
                        Channels = channels,
                        Channel = currentChannel,
                        Versions = versions,
                        Version = currentVersion

                    };
                }
                else
                {
                    info = new RemoteMachine
                    {
                        include = false,
                        Machine = connectionInfo[2],
                        UserName = connectionInfo[0],
                        Password = connectionInfo[1],
                        WorkGroup = connectionInfo[3],
                        Status = "Not Found",
                        Channels = null,
                        Channel = null,
                        Versions = null,
                        Version = null
                    };
                }

                GlobalObjects.ViewModel.RemoteMachines.Add(info);

                remoteClients.Add(info);
                //txtBxAddMachines.Clear();

            Dispatcher.Invoke(() =>
            {
                RemoteMachineList.ItemsSource = remoteClients;
                RemoteMachineList.Items.Refresh();
                WaitImage.Visibility = Visibility.Hidden;
            });

        }

        private async void AddComputersButton_Click(object sender, RoutedEventArgs e)
        {
            //placeholder text for data entry Username\Password\IP\Domain


            if (txtBxAddMachines.Text != "")
            {
                //parse text 

                WaitImage.Visibility = Visibility.Visible;

                //GlobalObjects.ViewModel.RemoteMachines = new List<RemoteMachine>();

                var connectionInfo = txtBxAddMachines.Text.Split('\\');

                //addMachines 
                await Task.Run(() => { addMachines(connectionInfo); }); 
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
                Filter = "CSV Files (.csv)|*.csv"
            };

            var result = dlg.ShowDialog();
            if (result == true)
            {

                List<string> versions = new List<String>();
                string line;

                try
                {
                    StreamReader file = new StreamReader(dlg.FileName);

                    while ((line = file.ReadLine()) != null)
                    {
                            //get client info 

                            
                        
                        string[] tempStrArray = line.Split(',');

                        addMachines(tempStrArray);
                          
                            //var info = new RemoteMachine
                            //{
                            //    include = false,
                            //    Machine = tempStrArray[0],
                            //    Status = "Not Found",
                            //    Channels = null,
                            //    Channel = null,
                            //    Versions = null,
                            //    Version = null
                            //};
                            //remoteClients.Add(info);
                            
                        

                    }

                    //RemoteMachineList.Items.Refresh();

                }
                catch (Exception)
                {

                }





            }
        }

        private void chkAll_Click(object sender, RoutedEventArgs e)
        {
            var handler = sender as System.Windows.Controls.CheckBox;
            foreach (var client in remoteClients)
            {
                client.include = handler.IsChecked.Value;
            }
            RemoteMachineList.Items.Refresh();
        }

        private async void btnUpdateRemote_Click(object sender, RoutedEventArgs e)
        {
            //need to have iterate over ALL entries in datagrid and grab their info once updating work



            foreach (var client in remoteClients)
            {

                if (client.include)
                {
                    GlobalObjects.ViewModel.BlockNavigation = true;
                    WaitImage.Visibility = Visibility.Visible;
                    try
                    {
                        var connectionInfo = new string[4] { client.UserName, client.Password, client.Machine, client.WorkGroup };
                        var installGenerator = new OfficeInstallManager(connectionInfo);

                        var newVersion = client.Version;
                        var newChannel = client.Channel;

                        await Task.Run(async () => { await installGenerator.initConnections(); });
                        var officeInstall = await installGenerator.CheckForOfficeInstallAsync();
                        var updateInfo = new List<string> { client.UserName, client.Password, client.Machine, client.WorkGroup, client.Channel.Name, client.Version.Number };

                        await ChangeOfficeChannelWmi(updateInfo, officeInstall); 

                    }
                    catch (Exception)
                    {
                        //powershell
                    }




                }
            }
            GlobalObjects.ViewModel.BlockNavigation = false;
            WaitImage.Visibility = Visibility.Hidden;

        }


        public async Task ChangeOfficeChannelWmi(List<string> updateinfo, OfficeInstallation LocalInstall)
        {
            var newChannel = updateinfo[4];
             
            await Task.Run(async () =>
            {
                var installOffice = new InstallOfficeWmi();

                installOffice.remoteUser = updateinfo[0];
                installOffice.remoteComputerName = updateinfo[2];
                installOffice.remoteDomain = updateinfo[3];
                installOffice.remotePass = updateinfo[1];
                installOffice.newChannel = updateinfo[4];
                installOffice.newVersion = updateinfo[5];
                installOffice.connectionNamespace = "\\root\\cimv2";

                try
                {

                    //UI Stuff

                    //installOffice = new InstallOffice();
                    //installOffice.UpdatingOfficeStatus += installOffice_UpdatingOfficeStatus;

                    //var newChannel = "";
                    //Dispatcher.Invoke(() =>
                    //{
                    //    UpdateStatus.Content = "Updating...";
                    //    newChannel = ((OfficeBranch)ProductBranch.SelectedItem).NewName;
                    //    ChangeChannel.IsEnabled = false;
                    //    NewVersion.IsEnabled = false;
                    //});

                    //SetItemState(LocalViewItem.Update, LocalViewState.Wait);

                    var ppDownloader = new ProPlusDownloader();
                    var baseUrl = await ppDownloader.GetChannelBaseUrlAsync(newChannel, OfficeEdition.Office32Bit);
                    if (string.IsNullOrEmpty(baseUrl))
                        throw (new Exception(string.Format("Cannot find BaseUrl for Channel: {0}", newChannel)));


                    //More UI 

                    var channelToChangeTo = updateinfo[5];
                    //if (NewVersionRow.Visibility != Visibility.Visible)
                    //{
                        //channelToChangeTo =
                        //    await ppDownloader.GetLatestVersionAsync(newChannel, OfficeEdition.Office32Bit);
                    //}
                    //else
                    //{
                    //    Dispatcher.Invoke(() =>
                    //    {
                    //        var manualVersion = NewVersion.Text;

                    //        if (string.IsNullOrEmpty(manualVersion) && NewVersion.SelectedItem != null)
                    //        {
                    //            manualVersion = ((Build)NewVersion.SelectedItem).Version;
                    //        }
                    //        if (!string.IsNullOrEmpty(manualVersion))
                    //        {
                    //            channelToChangeTo = manualVersion;
                    //        }
                    //    });
                    //}

                    if (string.IsNullOrEmpty(channelToChangeTo))
                    {
                        throw (new Exception("Version required"));
                    }
                    //else
                    //{
                    //    if (!channelToChangeTo.IsValidVersion())
                    //    {
                    //        throw (new Exception(string.Format("Invalid Version: {0}", channelToChangeTo)));
                    //    }
                    //}

                    //implement this in WMI ****************


                    await installOffice.ChangeOfficeChannel(channelToChangeTo, baseUrl);


              

                //Dispatcher.Invoke(() =>
                //{
                //    UpdateStatus.Content = "";
                //});

                var installGenerator = new OfficeInstallManager();
                    //if (LocalInstall.Installed)
                    //{
                    //    Dispatcher.Invoke(() =>
                    //    {
                    //        VersionLabel.Content = LocalInstall.Version;
                    //        ProductBranch.SelectedItem = LocalInstall.Channel;
                    //    });

                    //    if (LocalInstall.LatestVersionInstalled)
                    //    {
                    //        SetItemState(LocalViewItem.Update, LocalViewState.Success);
                    //    }
                    //    else
                    //    {
                    //        SetItemState(LocalViewItem.Update, LocalViewState.Action);
                    //        Dispatcher.Invoke(() =>
                    //        {
                    //            UpdateStatus.Content = "New version available  (" + LocalInstall.LatestVersion + ")";
                    //        });
                    //    }
                    //}
                }
                catch (Exception ex)
                {
                    //SetItemState(LocalViewItem.Update, LocalViewState.Fail);
                    //Dispatcher.Invoke(() =>
                    //{
                    //    UpdateStatus.Content = "The update failed";
                    //    ErrorText.Text = ex.Message;
                    //    RetryButtonColumn.Width = new GridLength(0, GridUnitType.Pixel);
                    //});
                    LogErrorMessage(ex);
                }
                finally
                {
                    //Dispatcher.Invoke(() =>
                    //{
                    //    ChangeChannel.IsEnabled = true;
                    //    NewVersion.IsEnabled = true;
                    //});
                }
            });
        }



        //#endregion
        private List<officeVersion> getVersions(OfficeBranch currentChannel, List<officeVersion> versions, string currentVersion)
        {

            foreach (var version in currentChannel.Versions)
            {
                if (version.Version.ToString() != currentVersion)
                {
                    var tempVersion = new officeVersion()
                    {
                        Number = version.Version.ToString()
                    };

                    versions.Add(tempVersion);
                }
            }

            return versions;
        }

        

        private void ProductChannel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //TODO add a channel change listener to get correct versions for that channel
            var selectedBranch = (sender as System.Windows.Controls.ComboBox).SelectedValue as string;
            var newVersions = new List<officeVersion>();
            var branches = GlobalObjects.ViewModel.Branches;
            var row = GetAncestorOfType<DataGridRow>(sender as System.Windows.Controls.ComboBox);
            var versionCB = row.FindChild<System.Windows.Controls.ComboBox>("ProductVersion");

            foreach (var branch in branches)
            {
                if (branch.NewName.ToString() == selectedBranch)
                {
                    newVersions = getVersions(branch, newVersions, "");
                    break;
                }
            }

          

            versionCB.ItemsSource = newVersions;
            versionCB.SelectedItem = newVersions[0];
            versionCB.Items.Refresh();
        }

        public T GetAncestorOfType<T>(FrameworkElement child) where T : FrameworkElement
        {
            var parent = VisualTreeHelper.GetParent(child);
            if (parent != null && !(parent is T))
                return (T)GetAncestorOfType<T>((FrameworkElement)parent);
            return (T)parent;
        }

        //private void setOfficeChannel()

        //private void setOfficeVersion(object sender, SelectionChangedEventArgs e)
        //{
        //    var row = GetAncestorOfType<DataGridRow>(sender as System.Windows.Controls.ComboBox);

        //    var handler = sender as System.Windows.Controls.ComboBox;
        //    var currentClient = remoteClients[row.GetIndex()];
        //    var tempVersion = new officeVersion()
        //    {
        //        Number = handler.SelectedValue.ToString()
        //    };

        //    currentClient.Version = tempVersion;
        //    RemoteMachineList.Items.Refresh();
        //}
    }
}

