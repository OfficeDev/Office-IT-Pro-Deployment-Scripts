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

        private void LogWmiErrorMessage(Exception ex, string[] connectionInfo = null)
        {
            var logPath = Path.GetTempPath() + "\\wmiLog.txt";

            if (System.IO.File.Exists(logPath))
            {
                using (TextWriter sw = System.IO.File.AppendText(logPath))
                {
                    if (connectionInfo == null)
                    {
                        sw.WriteLine(ex.Message);
                    }
                    else
                    {
                        sw.WriteLine("Client Error: " + ex.Message + "," + connectionInfo[2]);
                    }
                }
            }
            else
            {
                using (TextWriter sw = new StreamWriter(logPath))
                {
                    if (connectionInfo == null)
                    {
                        sw.WriteLine(ex.Message);
                    }
                    else
                    {
                        sw.WriteLine("Client Error: " + ex.Message + "," + connectionInfo[2]);
                    }
                }
            }

        }

        private void clearLogFile()
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

            var info = new RemoteMachine
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

            try
            {


                await installGenerator.initConnections();


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


                //txtBxAddMachines.Clear();        

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
                LogWmiErrorMessage(ex, connectionInfo);

            }
            finally
            {
                GlobalObjects.ViewModel.RemoteMachines.Add(info);
                remoteClients.Add(info);

                Dispatcher.Invoke(() =>
                {
                    RemoteMachineList.ItemsSource = remoteClients;
                    RemoteMachineList.Items.Refresh();
                    toggleControls(true);
                    WaitImage.Visibility = Visibility.Hidden;
                });
            }
        }
        private async void AddComputersButton_Click(object sender, RoutedEventArgs e)
        {
            //placeholder text for data entry Username\Password\IP\Domain


            if (txtBxAddMachines.Text != "")
            {
                //parse text 

                GlobalObjects.ViewModel.BlockNavigation = true;
                toggleControls(false);
                WaitImage.Visibility = Visibility.Visible;

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

            clearLogFile();
            GlobalObjects.ViewModel.BlockNavigation = true;
            toggleControls(false);
            WaitImage.Visibility = Visibility.Visible;
            var connectionInfo = new string[4];

            foreach (var client in remoteClients)
            {

                if (client.include)
                {
                    try
                    {
                        connectionInfo = new string[4] { client.UserName, client.Password, client.Machine, client.WorkGroup };
                        var installGenerator = new OfficeInstallManager(connectionInfo);

                        var newVersion = client.Version;
                        var newChannel = client.Channel;

                        await Task.Run(async () => { await installGenerator.initConnections(); });
                        var officeInstall = await installGenerator.CheckForOfficeInstallAsync();
                        var updateInfo = new List<string> { client.UserName, client.Password, client.Machine, client.WorkGroup, client.Channel.Name, client.Version.Number };

                        client.Status = "Updating";
                        RemoteMachineList.Items.Refresh();
                        await Task.Run(async () => { await ChangeOfficeChannelWmi(updateInfo, officeInstall); });
                        client.Status = "Success";
                        RemoteMachineList.Items.Refresh();


                    }
                    catch (Exception)// if fails via WMI, try via powershell
                    {
                        client.Status = "Error";
                        try
                        {
                            string PSPath = System.IO.Directory.GetCurrentDirectory() + "\\Resources\\PowershellAttempt.txt";
                            System.IO.File.Delete(PSPath);
                            Process p = new Process();
                            p.StartInfo.FileName = "Powershell.exe";                                //replace path to use local path                            switch out arguments so your program throws in the necessary args
                            p.StartInfo.Arguments = @"-ExecutionPolicy Bypass -NoExit -Command ""& {& '" + System.IO.Directory.GetCurrentDirectory() + "\\Resources\\UpdateScriptLaunch.ps1' -Channel " + client.Channel.Name + " -DisplayLevel $false -machineToRun " + client.Machine + " -UpdateToVersion " + client.Version.Number + "}\"";
                            p.Start();
                            p.WaitForExit();
                            p.Close();
                            string readtext = System.IO.File.ReadAllText(PSPath);
                            if (readtext.Contains("Update Completed") && !readtext.Contains("Update Not Running"))
                            {
                                client.Status = "Succeeded";
                            }
                            else
                            {
                                client.Status = "Failed";
                            }
                        }
                        catch (Exception)
                        {
                            client.Status = "Error";
                        }
                    }


                }
            }
            RemoteMachineList.Items.Refresh();
            GlobalObjects.ViewModel.BlockNavigation = false;
            WaitImage.Visibility = Visibility.Hidden;
            toggleControls(true);

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

                    var ppDownloader = new ProPlusDownloader();
                    var baseUrl = await ppDownloader.GetChannelBaseUrlAsync(newChannel, OfficeEdition.Office32Bit);
                    if (string.IsNullOrEmpty(baseUrl))
                        throw (new Exception(string.Format("Cannot find BaseUrl for Channel: {0}", newChannel)));



                    var channelToChangeTo = updateinfo[5];

                    if (string.IsNullOrEmpty(channelToChangeTo))
                    {
                        throw (new Exception("Version required"));
                    }

                    await installOffice.ChangeOfficeChannel(channelToChangeTo, baseUrl);
                    var installGenerator = new OfficeInstallManager();
                }
                catch (Exception ex)
                {
                    LogErrorMessage(ex);
                    LogWmiErrorMessage(ex, updateinfo.ToArray());
                    throw (new Exception("Update Failed"));
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

        private void toggleControls(bool enabled)
        {
            txtBxAddMachines.IsEnabled = enabled;
            AddComputersButton.IsEnabled = enabled;
            ImportComputersButton.IsEnabled = enabled;
            btnChangeChannelOrVersion.IsEnabled = enabled;
            btnUpdateRemote.IsEnabled = enabled;

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

