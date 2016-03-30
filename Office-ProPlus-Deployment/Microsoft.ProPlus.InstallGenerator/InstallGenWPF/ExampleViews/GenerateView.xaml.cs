using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
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

                MainTabControl.SelectedIndex = 0;

                InstallMsi.IsChecked = true;

                LogAnaylytics("/GenerateView", "Load");
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

        private async Task InstallerSign(string path)
        {


            var signPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "signtool.exe");
            System.IO.File.WriteAllBytes(signPath, Microsoft.OfficeProPlus.InstallGen.Presentation.Properties.Resources.signtool);

            var thumbprint = GlobalObjects.ViewModel.SelectedCertificate.ThumbPrint;

            path = path.Replace("\\\\", "\\");

            Process signProcess = new Process
            {
                StartInfo = new ProcessStartInfo()
                {
                    FileName = signPath,
                    Arguments = " sign /sha1 " + thumbprint + " " + path,
                    CreateNoWindow = true,
                    UseShellExecute = false
                }
            };


            signProcess.Start();

        }

        private async Task GenerateInstall(bool sign)
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
                            PreviousButton.IsEnabled = false;
                            GenerateButton.Content = "";
                            PreviousButton.Content = "";

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



                        LogAnaylytics("/GenerateView", "GenerateExe");
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


                        

                        LogAnaylytics("/GenerateView", "GenerateMSI");
                    }


                    await Task.Delay(500);

                    if (!String.IsNullOrEmpty(GlobalObjects.ViewModel.SelectedCertificate.ThumbPrint) && sign)
                    {
                        await InstallerSign(executablePath);

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
                    Console.WriteLine(ex.StackTrace);
                }
                finally
               { 
                    Dispatcher.Invoke(() =>
                    {
                        WaitImage.Visibility = Visibility.Hidden;
                        GenerateButton.IsEnabled = true;
                        PreviousButton.IsEnabled = true;

                        GenerateButton.Content = "Generate";
                        PreviousButton.Content = "Previous";
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

        private void LogAnaylytics(string path, string pageName)
        {
            try
            {
                GoogleAnalytics.Log(path, pageName);
            }
            catch { }
        }



        #region "Events"

        private void PreviousButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                
                this.TransitionTab(this, new TransitionTabEventArgs()
                {
                    Direction = TransitionTabDirection.Back
                });
           
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }
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
                var sign = SignInstaller.IsChecked.Value; 
                await GenerateInstall(sign);


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
                SourceFilePath.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;

                SourcePathLabel.IsEnabled = enabled;
                BuildFilePath.IsEnabled = enabled;
                BrowseSourcePathButton.IsEnabled = enabled;

                OpenFolderButton.IsEnabled = enabled;

                var buildFilePath = BuildFilePath.Text.Trim();
                if (enabled && !string.IsNullOrEmpty(buildFilePath))
                {
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

        private void SignWithCert_OnCheck(object sender, RoutedEventArgs e)
        {
            try
            {
                if (SignInstaller.IsChecked.Value)
                {
                    OpenCertificateBrowser.IsEnabled = true;
                    OpenCertGenerator.IsEnabled = true;

                }
                else
                {
                    OpenCertificateBrowser.IsEnabled = false;
                    OpenCertGenerator.IsEnabled = false;



                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
          
        }

        private CertificatesDialog certificatesDialog = null;
        private GenerateCertificate generateCertificateDialog = null;

        private void OpenCertificateBrowser_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                certificatesDialog = new CertificatesDialog();

                GlobalObjects.ViewModel.SetCertificates();
                if (GlobalObjects.ViewModel.Certificates != null)
                {
                    var certificateList = GlobalObjects.ViewModel.Certificates;
                    certificatesDialog.CertificateList.ItemsSource = certificateList;
                }

               
                certificatesDialog.Launch();
            }
             catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void OpenCertGenerator_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                generateCertificateDialog = new GenerateCertificate();


                generateCertificateDialog.Launch();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }
        
        #endregion

        #region "Info"

        private void ProductInfo_OnClick(object sender, RoutedEventArgs e)
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

                informationDialog.Height = 500;
                informationDialog.Width = 400;

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


        #endregion

      
    }
}
