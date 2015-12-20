using System;
using System.Collections.Generic;
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
using MahApps.Metro.Controls.Dialogs;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Enums;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGenerator.Implementation;
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
        public event MessageEventHandler InfoMessage;
        public event MessageEventHandler ErrorMessage;

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

                InstallMsi.IsChecked = true;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
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
                LogErrorMessage(ex);
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

                var buildPath = BuildFilePath.Text.Trim();
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

        private async Task GenerateInstall()
        {
            await Task.Run(async () =>
            {
                try
                {
                    FixFileExtension();

                    var executablePath = "";

                    for (var i = 1; i <= 2; i++)
                    {
                        await Dispatcher.InvokeAsync(() =>
                        {
                            executablePath = FileSavePath.Text.Trim();
                            WaitImage.Visibility = Visibility.Visible;
                            GenerateButton.IsEnabled = false;
                            GenerateButton.Content = "";

                            if (string.IsNullOrEmpty(executablePath))
                            {
                                if (i == 1)
                                {
                                    GetSaveFilePath();
                                }
                            }
                        });

                        if (!string.IsNullOrEmpty(executablePath))
                        {
                            if (executablePath.ToLower().EndsWith(".exe") ||
                                executablePath.ToLower().EndsWith(".msi"))
                            {
                                break;
                            }
                            else
                            {
                                await Dispatcher.InvokeAsync(GetSaveFilePath);
                            }
                        }
                    }

                    if (string.IsNullOrEmpty(executablePath))
                    {
                        //throw (new Exception("File Path Required"));
                        return;
                    }

                    var directoryPath = System.IO.Path.GetDirectoryName(executablePath);
                    if (directoryPath != null)
                    {
                        if (!Directory.Exists(directoryPath))
                        {

                            var result = MessageBox.Show("The directory '" + directoryPath + "' does not exist." +
                                Environment.NewLine + Environment.NewLine + "Create Directory?", "Create Directory", MessageBoxButton.YesNo,
                                MessageBoxImage.Question, MessageBoxResult.Yes);
                            Directory.CreateDirectory(directoryPath);

                            await Dispatcher.InvokeAsync(() =>
                            {
                                OpenExeFolderButton.IsEnabled = true;
                            });
                        }
                    }

                    var configFilePath =
                        Environment.ExpandEnvironmentVariables(@"%temp%\OfficeProPlus\" + Guid.NewGuid().ToString() +
                                                               ".xml");
                    Directory.CreateDirectory(Environment.ExpandEnvironmentVariables(@"%temp%\OfficeProPlus"));

                    GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml.Logging = new ODTLogging
                    {
                        Level = LoggingLevel.Standard,
                        Path = @"%temp%"
                    };

                    System.IO.File.WriteAllText(configFilePath, GlobalObjects.ViewModel.ConfigXmlParser.Xml);

                    string sourceFilePath = null;
                    await Dispatcher.InvokeAsync(() =>
                    {
                        if (IncludeBuild.IsChecked.HasValue && IncludeBuild.IsChecked.Value)
                        {
                            sourceFilePath = BuildFilePath.Text.Trim();
                            if (string.IsNullOrEmpty(sourceFilePath)) sourceFilePath = null;
                        }
                    });

                    var isInstallExe = false;
                    await Dispatcher.InvokeAsync(() =>
                    {
                        isInstallExe = InstallExecutable.IsChecked.HasValue && InstallExecutable.IsChecked.Value;
                    });

                    if (isInstallExe)
                    {
                        var generateExe = new OfficeInstallExecutableGenerator();
                        generateExe.Generate(new OfficeInstallProperties()
                        {
                            ConfigurationXmlPath = configFilePath,
                            OfficeVersion = OfficeVersion.Office2016,
                            ExecutablePath = executablePath,
                            SourceFilePath = sourceFilePath
                        });
                    }
                    else
                    {
                        var generateMsi = new OfficeInstallMsiGenerator();
                        generateMsi.Generate(new OfficeInstallProperties()
                        {
                            ConfigurationXmlPath = configFilePath,
                            OfficeVersion = OfficeVersion.Office2016,
                            ExecutablePath = executablePath,
                            SourceFilePath = sourceFilePath
                        });
                    }

                    if (InfoMessage != null)
                    {
                        if (isInstallExe)
                        {
                            InfoMessage(this, new MessageEventArgs()
                            {
                                Title = "Generate Executable",
                                Message = "File Generation Complete"
                            });
                        }
                        else
                        {
                            InfoMessage(this, new MessageEventArgs()
                            {
                                Title = "Generate MSI",
                                Message = "File Generation Complete"
                            });
                        }

                    }

                    await Task.Delay(500);
                }
                catch (Exception ex)
                {
                    LogErrorMessage(ex);
                }
                finally
                {
                    Dispatcher.Invoke(() =>
                    {
                        WaitImage.Visibility = Visibility.Hidden;
                        GenerateButton.IsEnabled = true;
                        GenerateButton.Content = "Generate";
                    });
                }
            });
        }

        public void LoadCurrentXml()
        {
            if (FileSavePath == null) return;
            if (xmlBrowser == null) return;
            if (BuildFilePath == null) return;

            if (GlobalObjects.ViewModel == null) return;
            if (GlobalObjects.ViewModel.ConfigXmlParser != null)
            {
                var configXml = GlobalObjects.ViewModel.ConfigXmlParser;

                if (!string.IsNullOrEmpty(GlobalObjects.ViewModel.ImportFile))
                {
                    FileSavePath.Text = GlobalObjects.ViewModel.ImportFile;
                }

                if (!string.IsNullOrEmpty(configXml.Xml))
                {
                    xmlBrowser.XmlDoc = configXml.Xml;
                }

                if (configXml.ConfigurationXml.Add != null)
                {
                    var currentBuildFilePath = BuildFilePath.Text;
                    if (string.IsNullOrEmpty(currentBuildFilePath))
                    {
                        BuildFilePath.Text = configXml.ConfigurationXml.Add.SourcePath;
                    }
                }

                var silentInstall = false;
                if (configXml.ConfigurationXml.Display != null)
                {
                    if (configXml.ConfigurationXml.Display.Level.HasValue &&
                        configXml.ConfigurationXml.Display.Level == DisplayLevel.None)
                    {
                        if (configXml.ConfigurationXml.Display.AcceptEULA.HasValue &&
                            configXml.ConfigurationXml.Display.AcceptEULA == true)
                        {
                            silentInstall = true;
                        }
                    }
                }

                SilentInstall.IsChecked = silentInstall;
            }
        }

        private void EnableGenerateButton()
        {
            var saveFileExists = false;
            var buildFolderExists = false;

            var filePath = FileSavePath.Text.Trim();
            if (!string.IsNullOrEmpty(filePath))
            {
                saveFileExists = true;
            }

            if (IncludeBuild.IsChecked.HasValue && IncludeBuild.IsChecked.Value)
            {
                var buildFolder = BuildFilePath.Text.Trim();
                if (!string.IsNullOrEmpty(buildFolder))
                {
                    if (Directory.Exists(buildFolder))
                    {
                        buildFolderExists = true;
                    }
                }
            }
            else
            {
                buildFolderExists = true;
            }


            if (buildFolderExists)
            {
                GenerateButton.IsEnabled = true;
            }
            else
            {
                GenerateButton.IsEnabled = false;
            }

        }

        private async void FixFileExtension()
        {
            if (FileSavePath == null) return;
            await Dispatcher.InvokeAsync(() =>
            {
                var currentPath = FileSavePath.Text;

                if (InstallExecutable.IsChecked.HasValue && InstallExecutable.IsChecked.Value)
                {
                    var match = Regex.Match(currentPath, ".exe$", RegexOptions.IgnoreCase);
                    if (!match.Success)
                    {
                        FileSavePath.Text = Regex.Replace(currentPath, @"\.\w{3}$", ".exe");
                    }

                    FileSavePath.SetValue(TextBoxHelper.WatermarkProperty, "Office365ProPlus.exe");
                }
                else
                {
                    var match = Regex.Match(currentPath, ".msi$", RegexOptions.IgnoreCase);
                    if (!match.Success)
                    {
                        FileSavePath.Text = Regex.Replace(currentPath, @"\.\w{3}$", ".msi");
                    }

                    FileSavePath.SetValue(TextBoxHelper.WatermarkProperty, "Office365ProPlus.msi");
                }
            });
        }

        private void GetSaveFilePath()
        {
            var openDialog = new SaveFileDialog()
            {
                Filter = "Executable|*.exe"
            };

            var fileName = "";
            if (InstallMsi.IsChecked.HasValue && InstallMsi.IsChecked.Value)
            {
                fileName = "OfficeProPlus.msi";
            }
            else
            {
                fileName = "OfficeProPlus.exe";
            }

            var currentFilePath = FileSavePath.Text;
            if (string.IsNullOrEmpty(currentFilePath))
            {
                openDialog.FileName = fileName;
            }
            else
            {
                var directoryPath = currentFilePath;
                if (currentFilePath.ToLower().EndsWith(".exe") || currentFilePath.ToLower().EndsWith(".msi"))
                {
                    directoryPath = System.IO.Path.GetDirectoryName(currentFilePath);
                    fileName = System.IO.Path.GetFileName(currentFilePath);
                }
                if (!string.IsNullOrEmpty(directoryPath))
                {
                    openDialog.InitialDirectory = directoryPath;
                    openDialog.FileName = fileName;
                }
            }

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

        #region "Events"

        private void InstallExecutable_OnChecked(object sender, RoutedEventArgs e)
        {
            try
            {
                FixFileExtension();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void SilentInstall_OnChecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (GlobalObjects.ViewModel.ConfigXmlParser != null)
                {
                    var configXml = GlobalObjects.ViewModel.ConfigXmlParser;

                    var silentInstall = SilentInstall.IsChecked.HasValue && SilentInstall.IsChecked.Value;

                    if (silentInstall)
                    {
                        configXml.ConfigurationXml.Display.AcceptEULA = true;
                        configXml.ConfigurationXml.Display.Level = DisplayLevel.None;
                    }
                    else
                    {
                        configXml.ConfigurationXml.Display.AcceptEULA = false;
                        configXml.ConfigurationXml.Display.Level = DisplayLevel.Full;
                    }

                    GlobalObjects.ViewModel.SilentInstall = silentInstall;

                    LoadCurrentXml();

                    if (!string.IsNullOrEmpty(configXml.Xml))
                    {
                        xmlBrowser.XmlDoc = configXml.Xml;
                    }
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
            finally
            {
                EnableGenerateButton();
            }
        }

        private void FileSavePath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                var openEnabled = false;

                var filePath = FileSavePath.Text.Trim();
                if (!string.IsNullOrEmpty(filePath))
                {
                    var folderPath = System.IO.Path.GetDirectoryName(filePath);
                    openEnabled = Directory.Exists(folderPath);
                }

                OpenExeFolderButton.IsEnabled = openEnabled;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
            finally
            {
                EnableGenerateButton();
            }
        }

        private async void OpenExeFolderButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var filePath = FileSavePath.Text.Trim();
                if (string.IsNullOrEmpty(filePath)) return;

                var folderPath = System.IO.Path.GetDirectoryName(filePath);
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

        private async void GenerateButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                await GenerateInstall();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }
        
        private async void IncludeBuild_OnChecked(object sender, RoutedEventArgs e)
        {
            try
            {
                var enabled = IncludeBuild.IsChecked.HasValue && IncludeBuild.IsChecked.Value;

                SourcePathLabel.IsEnabled = enabled;
                BuildFilePath.IsEnabled = enabled;
                BrowseSourcePathButton.IsEnabled = enabled;

                OpenFolderButton.IsEnabled = enabled;
                DownloadButton.IsEnabled = enabled;

                var buildFilePath = BuildFilePath.Text.Trim();
                if (enabled && !string.IsNullOrEmpty(buildFilePath))
                {
                    DownloadButton.IsEnabled = true;
                    if (await GlobalObjects.DirectoryExists(buildFilePath))
                    {
                        OpenFolderButton.IsEnabled = true;
                    }
                    else
                    {
                        OpenFolderButton.IsEnabled = false;
                    }
                }
                else
                {
                    OpenFolderButton.IsEnabled = false;
                    DownloadButton.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
            finally
            {
                EnableGenerateButton();
            }
        }

        private async void OpenFolderButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var folderPath = BuildFilePath.Text.Trim();
                if (!string.IsNullOrEmpty(folderPath))
                {
                    if (await GlobalObjects.DirectoryExists(folderPath))
                    {
                        Process.Start("explorer", folderPath);
                    }
                    else
                    {
                        MessageBox.Show("Directory path does not exist.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

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
                LogErrorMessage(ex);
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                GetSaveFilePath();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
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
                    LogErrorMessage(ex);
                }
            }
        }
        
        private async void BuildFilePath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                var enabled = false;
                var openFolderEnabled = false;
                if (BuildFilePath.Text.Length > 0)
                {
                    var match = Regex.Match(BuildFilePath.Text, @"^\w:\\|\\\\.*\\..*");
                    if (match.Success)
                    {
                        enabled = true;

                        var folderExists = await GlobalObjects.DirectoryExists(BuildFilePath.Text);
                        if (!folderExists)
                        {
                            folderExists = await GlobalObjects.DirectoryExists(BuildFilePath.Text);
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
            finally
            {
                EnableGenerateButton();
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
                LogErrorMessage(ex);
            }
        }
        
        #endregion

 
    }
}
