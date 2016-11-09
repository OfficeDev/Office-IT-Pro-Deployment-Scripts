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
using System.Windows.Threading;
using UserControl = System.Windows.Controls.UserControl;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Windows.Data;
using Microsoft.OfficeProPlus.InstallGenerator.Models;

namespace MetroDemo.ExampleViews
{

    public partial class RemoteView : UserControl
    {
        #region Declarations
        private CancellationTokenSource _tokenSource = new CancellationTokenSource();
        private RemoteChannelVersionDialog remoteUpdateDialog = null;
        private RemoteClientInfoDialog remoteClientDialog = null;
        private InformationDialog informationDialog = null;

        public event MessageEventHandler ErrorMessage;

        public MetroWindow MainWindow { get; set; }

        private int _cachedIndex = 0;



        private OfficeInstall RemoteInstall { get; set; }
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

        private async Task CollectMachineData(RemoteMachine remoteMachine)
        {
            var installGenerator = new OfficeInstallManager(remoteMachine.Machine, remoteMachine.WorkGroup, remoteMachine.UserName, remoteMachine.Password);
            try
            {
                remoteMachine.Status = "Checking...";

                RemoteMachineList.ItemsSource = null;
                RemoteMachineList.ItemsSource = remoteClients;

                await installGenerator.InitConnections();

                var officeInstall = await installGenerator.CheckForOfficeInstallAsync();

                var channels = new List<Channel>();

                var versions = new List<officeVersion>();

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
                        if (branch.NewName.ToString() != officeInstall.Channel)
                        {
                            var tempChannel = new Channel()
                            {
                                Name = branch.NewName.ToString()
                            };
                            channels.Add(tempChannel);
                        }
                        else
                        {
                            versions = GetVersions(branch, versions, currentVersion.Number);
                        }
                    }

                    remoteMachine.include = false;
                    remoteMachine.Machine = remoteMachine.Machine;
                    remoteMachine.UserName = remoteMachine.UserName;
                    remoteMachine.Password = remoteMachine.Password;
                    remoteMachine.WorkGroup = remoteMachine.WorkGroup;
                    remoteMachine.Status = "Found";
                    remoteMachine.Channels = channels;
                    remoteMachine.Channel = currentChannel;
                    remoteMachine.OriginalChannel = currentChannel;
                    remoteMachine.Versions = versions;
                    remoteMachine.Version = currentVersion;
                    remoteMachine.OriginalVersion = currentVersion;
                 
                    RemoteMachineList.ItemsSource = null;
                    RemoteMachineList.ItemsSource = remoteClients;
                }

            }
            catch (Exception ex)
            {
                remoteMachine.Status = ex.Message;
                RemoteMachineList.ItemsSource = null;
                RemoteMachineList.ItemsSource = remoteClients;

                LogWmiErrorMessage(ex, new RemoteComputer()
                {
                    Name = remoteMachine.Machine,
                    Domain = remoteMachine.WorkGroup,
                    UserName = remoteMachine.UserName,
                    Password = remoteMachine.Password
                });
            }
            if (remoteMachine.Status == "Checking...")
            {
                remoteMachine.Status = "Not Found";
            }
            RemoteMachineList.ItemsSource = null;
            RemoteMachineList.ItemsSource = remoteClients;
        }

        private Task UpdateMachineMultiThread(RemoteMachine mech, int i)
        {
            //allows for updates to be run at the same time.  otherwise, each task stays in queue until al awaits have finished
            return Task.Run(() =>
            {
                UpdateMachine(mech, i);
            });
        }


        private async Task UpdateMachine(RemoteMachine client, int i)
        {
                var connectionInfo = new string[4];
                RemoteMachineList.Dispatcher.Invoke(new Action(() => {
                                                                         RemoteMachineList.UpdateLayout();
                }));

                var row = (DataGridRow) RemoteMachineList.ItemContainerGenerator.ContainerFromIndex(i);

                System.Windows.Controls.TextBlock statusText = null;


                if (row != null)
                {
                    row.Dispatcher.Invoke(new Action(() =>
                    {
                        statusText = row.FindChild<System.Windows.Controls.TextBlock>("TxtStatus");
                    }));
                }
                else
                {
                    return;
                }


                RemoteMachineList.Dispatcher.Invoke(new Action(() =>
                {
                    RemoteMachineList.Items.Refresh();
                }));


                try
                {

                    client.Status = "Updating";

                    statusText.Dispatcher.Invoke(new Action(() =>
                    {
                        statusText.Text = "Updating";
                    }));

                    RemoteMachineList.Dispatcher.Invoke(new Action(() =>
                    {
                        RemoteMachineList.Items.Refresh();
                    }));

                    //throw (new Exception(""));

                    connectionInfo = new string[4] {client.UserName, client.Password, client.Machine, client.WorkGroup};
                    var installGenerator = new OfficeInstallManager(client.Machine, client.WorkGroup, client.UserName,
                        client.Password);

                    var newVersion = client.Version;
                    var newChannel = client.Channel;

                    await Task.Run(async () => { await installGenerator.InitConnections(); });
                    var officeInstall = await Task.Run(() => { return installGenerator.CheckForOfficeInstallAsync(); });

                    await Task.Run(async () => { await ChangeOfficeChannelWmi(client, officeInstall); });

                    client.Status = "Success";

                    statusText.Dispatcher.Invoke(new Action(() =>
                    {
                        statusText.Text = "Success";
                    }));

                    RemoteMachineList.Dispatcher.Invoke(new Action(() =>
                    {
                        RemoteMachineList.Items.Refresh();
                    }));

                }
                catch (Exception ex) // if fails via WMI, try via powershell
                {


                    try
                    {
                        LogWmiErrorMessage(ex, new RemoteComputer()
                        {
                            Name = client.Machine,
                            Domain = client.WorkGroup,
                            Password = client.Password,
                            UserName = client.UserName
                        });

                        string PSPath = System.IO.Path.GetTempPath() + client.Machine + "PowershellAttempt.txt";
                        System.IO.File.Delete(PSPath);
                    
                        var powerShellInstance = System.Management.Automation.PowerShell.Create();
                        if (!String.IsNullOrEmpty(client.OriginalVersion.Number) ||
                            client.Version.Number != client.OriginalVersion.Number)
                        {

                            powerShellInstance.AddScript(System.IO.Directory.GetCurrentDirectory() +
                                                         "\\Resources\\UpdateScriptLaunch.ps1 -Channel " +
                                                         client.Channel.Name + " -DisplayLevel $false -machineToRun " +
                                                         client.Machine + " -UpdateToVersion " + client.Version.Number);
                            var asyncResult = powerShellInstance.BeginInvoke();
                            //possible make async so toolkit doesn't freeze the console
                            client.Status = "Updating...";
                        //statusText.Dispatcher.Invoke(new Action(() =>
                        //{
                        //    statusText.Text = "Updating...";
                        //}));

                        //RemoteMachineList.Dispatcher.Invoke(new Action(() =>
                        //{
                        //    RemoteMachineList.Items.Refresh();
                        //    RemoteMachineList.ItemsSource = null;
                        //    RemoteMachineList.ItemsSource = remoteClients;
                        //}));
                        
                        powerShellInstance.EndInvoke(asyncResult);
                        }
                        else
                        {
                            //statusText.Dispatcher.Invoke(new Action(() =>
                            //{
                            //    statusText.Text = "Success";
                            //}));
                        client.Status = "Success";
                        //RemoteMachineList.Dispatcher.Invoke(new Action(() =>
                        //    {
                        //        RemoteMachineList.Items.Refresh();
                        //    }));
                        RemoteMachineList.Items.Refresh();
                    }




                        PsUpdateExited(PSPath, statusText, client);


                    }
                    catch (Exception ex1)
                    {


                        client.Status = "Error: " + ex.Message;
                        using (System.IO.StreamWriter file =
                            new System.IO.StreamWriter(
                                System.IO.Path.GetTempPath() + client.Machine + "PowershellError.txt", true))
                        {
                            file.WriteLine(ex1.Message);
                            file.WriteLine(ex1.StackTrace);
                        }
                        statusText.Dispatcher.Invoke(new Action(() =>
                        {
                            statusText.Text = "Error: " + ex.Message;
                        }));

                        RemoteMachineList.Dispatcher.Invoke(new Action(() =>
                        {
                            RemoteMachineList.Items.Refresh();
                        }));
                    }

                }
        }

        private void PsUpdateExited(string psPath, TextBlock statusText, RemoteMachine client)
        {

            string readtext = System.IO.File.ReadAllText(psPath);
            if (readtext.Contains("Update Completed") && !readtext.Contains("Update Not Running"))
            {

                client.Status = "Success";
                statusText.Dispatcher.Invoke(new Action(() =>
                {
                    statusText.Text = "Success";
                }));

                RemoteMachineList.Dispatcher.Invoke(new Action(() =>
                {
                    RemoteMachineList.Items.Refresh();
                }));


            }
            else
            {

                client.Status = "Failed";
                //statusText.Text = "Failed";
                //RemoteMachineList.Items.Refresh();
                statusText.Dispatcher.Invoke(new Action(() =>
                {
                    statusText.Text = "Failed";
                }));

                RemoteMachineList.Dispatcher.Invoke(new Action(() =>
                {
                    RemoteMachineList.Items.Refresh();
                }));
            }

        }

        public async Task ChangeOfficeChannelWmi(RemoteMachine client, OfficeInstallation localInstall)
        {
            var newChannel = client.Channel.ToString();
            string version = null;
            if (client.Version != null)
            {
                version = client.Version.ToString();
            }

            await Task.Run(async () =>
            {
                var installOffice = new InstallOfficeWmi
                {
                    remoteUser = client.UserName,
                    remoteComputerName = client.Machine,
                    remoteDomain = client.WorkGroup,
                    remotePass = client.Password,
                    newChannel = client.Channel.ToString(),
                    newVersion = version,
                    connectionNamespace = "\\root\\cimv2"
                };

                try
                {
                    var ppDownloader = new ProPlusDownloader();
                    var baseUrl = await ppDownloader.GetChannelBaseUrlAsync(newChannel, OfficeEdition.Office32Bit);
                    if (string.IsNullOrEmpty(baseUrl))
                        throw (new Exception(string.Format("Cannot find BaseUrl for Channel: {0}", newChannel)));

                    var channelToChangeTo = newChannel;

                    if (string.IsNullOrEmpty(channelToChangeTo))
                    {
                        throw (new Exception("Version required"));
                    }

                    await installOffice.ChangeOfficeChannel(channelToChangeTo, baseUrl);
                    var installGenerator = new OfficeInstallManager();
                }
                catch (Exception ex)
                {
                    LogWmiErrorMessage(ex, new RemoteComputer()
                    {
                        Name = client.Machine,
                        Domain = client.WorkGroup,
                        UserName = client.UserName,
                        Password = client.Password
                    });
                    throw (new Exception("Update Failed"));
                }

            });
        }

        private static List<officeVersion> GetVersions(OfficeBranch currentChannel, List<officeVersion> versions, string currentVersion)
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

            var selectedVersion = new officeVersion()
            {
                Number = currentVersion
            };

            versions.Insert(0, selectedVersion);

            return versions;
        }

        private void ToggleControls(bool enabled)
        {
            AddComputersButton.IsEnabled = enabled;
            btnChangeChannelOrVersion.IsEnabled = enabled;
            btnUpdateRemote.IsEnabled = enabled;
            btnRemoveComputers.IsEnabled = enabled;

        }

        private void ToggleControlsMulti(bool enabled)
        {

            AddComputersButton.Dispatcher.Invoke(new Action(() => { this.IsEnabled = enabled;}));
            btnChangeChannelOrVersion.Dispatcher.Invoke(new Action(() => { this.IsEnabled = enabled; }));
            btnUpdateRemote.Dispatcher.Invoke(new Action(() => { this.IsEnabled = enabled; }));
            btnRemoveComputers.Dispatcher.Invoke(new Action(() => { this.IsEnabled = enabled; }));

        }
        
        public T GetAncestorOfType<T>(FrameworkElement child) where T : FrameworkElement
        {
            var parent = VisualTreeHelper.GetParent(child);
            if (parent != null && !(parent is T))
                return (T)GetAncestorOfType<T>((FrameworkElement)parent);
            return (T)parent;
        }
        
        private void LaunchInformationDialog(string sourceName)
        {
            try
            {
                if (informationDialog == null)
                {

                    informationDialog = new InformationDialog
                    {
                        Height = 700,
                        Width = 600
                    };
                    informationDialog.Closed += (o, args) =>
                    {
                        informationDialog = null;
                    };
                    informationDialog.Closing += (o, args) =>
                    {

                    };
                }

                informationDialog.Height = 700;
                informationDialog.Width = 600;

                var filePath = AppDomain.CurrentDomain.BaseDirectory + @"HelpFiles\" + sourceName + ".html";
                var helpFile = System.IO.File.ReadAllText(filePath);

                informationDialog.HelpInfo.NavigateToString(helpFile);
                informationDialog.Launch();

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        #region Logging

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

        private void LogWmiErrorMessage(Exception ex, RemoteComputer machine = null)
        {
            var computerName = "";

            if (machine != null)
            {
                computerName = machine.Name;
            }

            var logPath = Path.GetTempPath() + "\\" + computerName + "wmiLog.txt";
            var stackTrace = new StackTrace(ex, true);
            var frame = stackTrace.GetFrame(0);
            var lineNumber = frame.GetFileColumnNumber();

            if (System.IO.File.Exists(logPath))
            {
                using (TextWriter sw = System.IO.File.AppendText(logPath))
                {
                    if (machine == null)
                    {
                        sw.WriteLine(ex.Message);
                    }
                    else
                    {
                        sw.WriteLine("Client Error: " + ex.Message + "," + computerName + "," + stackTrace.ToString());
                    }
                }
            }
            else
            {
                using (TextWriter sw = new StreamWriter(logPath))
                {
                    if (machine == null)
                    {
                        sw.WriteLine(ex.Message);
                    }
                    else
                    {
                        sw.WriteLine("Client Error: " + ex.Message + "," + computerName + "," + stackTrace.ToString());
                    }
                }
            }

        }

        private void ClearLogFile()
        {
            var logPath = Path.GetTempPath() + "\\wmiLog.txt";

            if (System.IO.File.Exists(logPath))
            {
                System.IO.File.WriteAllText(logPath, String.Empty);
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

        #endregion

        #region Events

        private async void btnUpdateRemote_Click(object sender, RoutedEventArgs e)
        {
            List<Task> updateTasks = new List<Task>();


            try
            {
                ClearLogFile();
                GlobalObjects.ViewModel.BlockNavigation = true;
                ToggleControlsMulti(false);
                WaitImage.Visibility = Visibility.Visible;
                RemoteMachineList.Items.Refresh();

                for (var i = 0; i < remoteClients.Count; i++)
                {
                    var client = remoteClients[i];

                    //use common username/password for each client
                    if (!string.IsNullOrEmpty(GlobalObjects.ViewModel.GetUsername()) &&
                        !string.IsNullOrEmpty(GlobalObjects.ViewModel.GetPassword()))
                    {
                        client.Password = GlobalObjects.ViewModel.GetPassword();
                        client.UserName = GlobalObjects.ViewModel.GetUsername();
                        client.WorkGroup = GlobalObjects.ViewModel.GetDomain();
                    }

                    if (client.include && client.Status.Trim() != "Not Found")
                    {
                        int tempInt = i;
                        Task tempTask = UpdateMachineMultiThread(client, tempInt);// UpdateMachine(client, i);
                        updateTasks.Add(tempTask);
                    }
                }

                await Task.Factory.ContinueWhenAll(updateTasks.ToArray(), t =>
                {
                    RemoteMachineList.Dispatcher.Invoke(new Action(() =>
                    {
                        RemoteMachineList.Items.Refresh();
                    }));
                });
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);

                RemoteMachineList.Dispatcher.Invoke(new Action(() =>
                {
                    RemoteMachineList.Items.Refresh();
                }));
            }

        }

        private void ProductChannel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {


                var selectedBranch = (sender as System.Windows.Controls.ComboBox).SelectedValue as string;
                var newVersions = new List<officeVersion>();
                var branches = GlobalObjects.ViewModel.Branches;
                var row = GetAncestorOfType<DataGridRow>(sender as System.Windows.Controls.ComboBox);
                var versionCB = row.FindChild<System.Windows.Controls.ComboBox>("ProductVersion");


                var branch = branches.Find(a => a.NewName == GlobalObjects.ViewModel.newChannel);

                if (String.IsNullOrEmpty(GlobalObjects.ViewModel.newChannel))
                {
                    branch = branches.Find(a => a.NewName == selectedBranch);
                }


                foreach (var version in branch.Versions)
                {
                    var tempVersion = new officeVersion()
                    {
                        Number = version.Version
                    };

                    newVersions.Add(tempVersion);
                }

                if (row != null)
                {
                    var client = remoteClients[row.GetIndex()];
                    versionCB.ItemsSource = newVersions;
                    versionCB.Items.Refresh();

                    if (String.IsNullOrEmpty(GlobalObjects.ViewModel.newVersion))
                    {
                        versionCB.SelectedValue = client.Version.Number;
                    }
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
        
        private async void RemoteClientInfoDialog_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                var dialog = (RemoteClientInfoDialog)sender;

                if (GlobalObjects.ViewModel.RemoteConnectionInfo() == null || dialog.Result != DialogResult.OK) return;
                var connectionInfo = GlobalObjects.ViewModel.RemoteConnectionInfo();

                foreach (var client in connectionInfo)
                {
                    var info = new RemoteMachine
                    {
                        include = false,
                        Machine = client.Name,
                        UserName = client.UserName,
                        Password = client.Password,
                        WorkGroup = client.Domain,
                        Status = "",
                        Channels = null,
                        Channel = null,
                        OriginalChannel = null,
                        Versions = null,
                        Version = null,
                        OriginalVersion = null
                    };

                    GlobalObjects.ViewModel.RemoteMachines.Add(info);
                    remoteClients.Add(info);
                }

                RemoteMachineList.ItemsSource = null;
                RemoteMachineList.ItemsSource = remoteClients;

                var taskList = new List<Task>();
                foreach (var client in remoteClients)
                {
                    taskList.Add(CollectMachineData(client));
                    await Task.Delay(1000);
                }

                foreach (var task in taskList)
                {
                    await task;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
            finally
            {
                ToggleControlsMulti(true);
                WaitImage.Dispatcher.Invoke(new Action(() =>
                {
                    WaitImage.Visibility = Visibility.Hidden;
                }));
                GlobalObjects.ViewModel.BlockNavigation = false;
            }
        }

        private void AddComputersButton_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                remoteClientDialog = new RemoteClientInfoDialog();
                remoteClientDialog.Closing += RemoteClientInfoDialog_Closing;

                //ToggleControlsMulti(false);
                //WaitImage.Visibility = Visibility.Visible;
                GlobalObjects.ViewModel.BlockNavigation = true;

                remoteClientDialog.Launch();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }

        }

        private void btnChangeChannelOrVersion_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                remoteUpdateDialog = new RemoteChannelVersionDialog();
                remoteUpdateDialog.Closing += RemoteUpdateDialog_Closing;


                ToggleControlsMulti(false);
                WaitImage.Visibility = Visibility.Visible;
                GlobalObjects.ViewModel.BlockNavigation = true;

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
                var newVersion = new officeVersion

                {
                    Number = GlobalObjects.ViewModel.newVersion
                };

                var newChannel = new Channel
                {
                    Name = GlobalObjects.ViewModel.newChannel
                };

                if (dialog.Result == DialogResult.OK && !String.IsNullOrEmpty(newVersion.Number) && !String.IsNullOrEmpty(newChannel.Name))
                {
                    for (var i = 0; i < remoteClients.Count; i++)
                    {
                        if (remoteClients[i].include)
                        {
                            var row = (DataGridRow)RemoteMachineList.ItemContainerGenerator.ContainerFromIndex(i);

                            var channelCB = row.FindChild<System.Windows.Controls.ComboBox>("ProductChannel");
                            var versionCB = row.FindChild<System.Windows.Controls.ComboBox>("ProductVersion");

                            channelCB.SelectedValue = GlobalObjects.ViewModel.newChannel;
                            versionCB.SelectedValue = GlobalObjects.ViewModel.newVersion;

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
            finally
            {
                GlobalObjects.ViewModel.newChannel = null;
                GlobalObjects.ViewModel.newVersion = null;

                ToggleControlsMulti(true);
                WaitImage.Dispatcher.Invoke(new Action(() =>
                {
                    WaitImage.Visibility = Visibility.Hidden;
                }));
                GlobalObjects.ViewModel.BlockNavigation = false;

                RemoteMachineList.Items.Refresh();
            }
        }
        private void chkAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var handler = sender as System.Windows.Controls.CheckBox;
                foreach (var client in remoteClients)
                {
                    client.include = handler.IsChecked.Value;
                }
            }
            catch (Exception) { }
            RemoteMachineList.Items.Refresh();
        }
        
        private void btnRemoveComputers_Click(object sender, RoutedEventArgs e)
        {
            remoteClients.RemoveAll(a => a.include == true);

            RemoteMachineList.ItemsSource = remoteClients;
            RemoteMachineList.Items.Refresh();
        }

        private void RemoteConfiguration_Click(object sender, RoutedEventArgs e)
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

        
        private void TxtStatus_OnTargetUpdated(object sender, DataTransferEventArgs e)
        {
            //check text in status, block nav if it's working on something, unblock if not
            bool inProgress = false;
            foreach (var client in remoteClients)
            {
                if (client.Status.ToLower().Contains("checking") || client.Status.ToLower().Contains("updat"))
                {
                    inProgress = true;
                }
            }
            if (inProgress)
            {
                GlobalObjects.ViewModel.BlockNavigation = true;
                ToggleControlsMulti(false);
                WaitImage.Visibility = Visibility.Visible;
            }
            else
            {
                ToggleControlsMulti(true);
                WaitImage.Visibility = Visibility.Hidden;
                GlobalObjects.ViewModel.BlockNavigation = false;
            }
        }
    }

}