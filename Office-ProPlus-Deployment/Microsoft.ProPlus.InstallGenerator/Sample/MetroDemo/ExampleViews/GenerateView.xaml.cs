using System;
using System.Collections.Generic;
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
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;
using OfficeInstallGenerator;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using UserControl = System.Windows.Controls.UserControl;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class GenerateView : UserControl
    {

        public event TransitionTabEventHandler TransitionTab;
        private CancellationTokenSource _tokenSource = new CancellationTokenSource();
        private Task _downloadTask = null;

        public GenerateView()
        {
            InitializeComponent();
        }

        private void GenerateView_OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                LoadCurrentXml();

                if (xmlBrowser.InstallOffice == null)
                {
                    xmlBrowser.InstallOffice += InstallOffice;
                }

                InstallExecutable.IsChecked = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void InstallOffice(object sender, InstallOfficeEventArgs args)
        {
            try
            {
                var installGenerator = new OfficeInstallExecutableGenerator();
                installGenerator.InstallOffice(args.Xml);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private async Task DownloadOfficeFiles()
        {
            try
            {
                _tokenSource = new CancellationTokenSource();

                BuildFilePath.IsReadOnly = true;
                BrowseSourcePathButton.IsEnabled = false;

                DownloadProgressBar.Maximum = 100;
                DownloadPercent.Content = "";

                var proPlusDownloader = new ProPlusDownloader();
                proPlusDownloader.DownloadFileProgress += async (senderfp, progress) =>
                {
                    var percent = progress.PercentageComplete;
                    if (percent > 0)
                    {
                        Dispatcher.Invoke(() => { 
                            DownloadPercent.Content = percent + "%";
                            DownloadProgressBar.Value = Convert.ToInt32(Math.Round(percent, 0));
                        });
                    }
                };

                var buildPath = BuildFilePath.Text;
                if (string.IsNullOrEmpty(buildPath)) return;

                var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
                var languages =
                    (from product in configXml.Add.Products
                        from language in product.Languages
                        select language.ID.ToLower()).Distinct().ToList();

                string branch = null;
                if (configXml.Add.Branch.HasValue)
                {
                    branch = configXml.Add.Branch.Value.ToString();
                }

                var officeEdition = OfficeEdition.Office32Bit;
                if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office64Bit)
                {
                    officeEdition = OfficeEdition.Office64Bit;
                }

                buildPath = GlobalObjects.SetBranchFolderPath(branch, buildPath);
                Directory.CreateDirectory(buildPath);

                BuildFilePath.Text = buildPath;

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
                BuildFilePath.IsReadOnly = false;
                BrowseSourcePathButton.IsEnabled = true;
                DownloadProgressBar.Value = 0;
                DownloadPercent.Content = "";

                DownloadButton.Content = "Download";
                _tokenSource = new CancellationTokenSource();
            }
        }

        public void LoadCurrentXml()
        {
            if (GlobalObjects.ViewModel.ConfigXmlParser != null)
            {
                if (!string.IsNullOrEmpty(GlobalObjects.ViewModel.ConfigXmlParser.Xml))
                {
                    xmlBrowser.XmlDoc = GlobalObjects.ViewModel.ConfigXmlParser.Xml;
                }
            }
        }

        #region "Events"

        private void displayNext_Click(object sender, RoutedEventArgs e)
        {
            try
            {
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

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openDialog = new SaveFileDialog()
                {
                    Filter = "Executable|*.exe"
                };

                if (InstallMsi.IsChecked.HasValue && InstallMsi.IsChecked.Value)
                {
                    openDialog.Filter = "MSI|*.msi";
                }

                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    var filePath = openDialog.FileName;
                    FileSavePath.Text = filePath;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
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
                    MessageBox.Show("ERROR: " + ex.Message);
                }
            }
        }
        
        private void BuildFilePath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                var enabled = false;
                if (BuildFilePath.Text.Length > 0)
                {
                    var match = Regex.Match(BuildFilePath.Text, @"^\w:\\|\\\\.*\\..*");
                    if (match.Success)
                    {
                        enabled = true;
                    } 
                }

                DownloadButton.IsEnabled = enabled;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void BrowseSourcePathButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg1 = new Ionic.Utils.FolderBrowserDialogEx
                {
                    Description = "Select a folder to download files to:",
                    ShowNewFolderButton = true,
                    ShowEditBox = true,
                    SelectedPath = BuildFilePath.Text,
                    ShowFullPathInEditBox = true,
                    RootFolder = System.Environment.SpecialFolder.MyComputer
                };
                //dlg1.NewStyle = false;

                // Show the FolderBrowserDialog.
                DialogResult result = dlg1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    BuildFilePath.Text = dlg1.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }
        

        #endregion

    }
}
