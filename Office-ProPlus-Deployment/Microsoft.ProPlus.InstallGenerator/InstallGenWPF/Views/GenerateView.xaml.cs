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
using System.Windows.Forms;
using MahApps.Metro.Controls;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Enums;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGenerator;
using Microsoft.OfficeProPlus.InstallGenerator.Extensions;
using Microsoft.OfficeProPlus.InstallGenerator.Implementation;
using OfficeInstallGenerator;
using MessageBox = System.Windows.MessageBox;
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

                LoadFolder();

                MajorVersion.Value = 1;
                MinorVersion.Value = 0;
                ReleaseVersion.Value = 0;

                LogAnaylytics("/GenerateView", "Load");
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LoadFolder()
        {
            if (string.IsNullOrEmpty(GlobalObjects.ViewModel.DownloadFolderPath)) return;
            if (!Directory.Exists(GlobalObjects.ViewModel.DownloadFolderPath)) return;

            var mainFolderPath = GlobalObjects.ViewModel.DownloadFolderPath;
            var folderPath = mainFolderPath;
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Add?.Branch != null)
            {
                switch (configXml.Add.Branch)
                {
                    case Branch.Business:
                        if (Directory.Exists(folderPath + @"\DC"))
                        {
                            folderPath = mainFolderPath + @"\DC";
                        }
                        else if (Directory.Exists(folderPath + @"\Deferred"))
                        {
                            folderPath = mainFolderPath + @"\Deferred";
                        }
                        else if (Directory.Exists(folderPath + @"\Business"))
                        {
                            folderPath = mainFolderPath + @"\Business";
                        }
                        break;
                    case Branch.Current:
                        if (Directory.Exists(folderPath + @"\CC"))
                        {
                            folderPath = mainFolderPath + @"\CC";
                        }
                        else if (Directory.Exists(folderPath + @"\Current"))
                        {
                            folderPath = mainFolderPath + @"\Current";
                        }
                        break;
                    case Branch.FirstReleaseCurrent:
                        if (Directory.Exists(folderPath + @"\FRCC"))
                        {
                            folderPath = mainFolderPath + @"\FRCC";
                        }
                        else if (Directory.Exists(folderPath + @"\FirstReleaseCurrent"))
                        {
                            folderPath = mainFolderPath + @"\FirstReleaseCurrent";
                        }
                        break;
                    case Branch.Validation:
                        if (Directory.Exists(folderPath + @"\FRDC"))
                        {
                            folderPath = mainFolderPath + @"\FRDC";
                        }
                        else if (Directory.Exists(folderPath + @"\FirstReleaseDeferred"))
                        {
                            folderPath = mainFolderPath + @"\FirstReleaseDeferred";
                        }
                        else if (Directory.Exists(folderPath + @"\FirstReleaseBusiness"))
                        {
                            folderPath = mainFolderPath + @"\FirstReleaseBusiness";
                        }
                        break;
                    case Branch.FirstReleaseBusiness:
                        if (Directory.Exists(folderPath + @"\FRDC"))
                        {
                            folderPath = mainFolderPath + @"\FRDC";
                        }
                        else if (Directory.Exists(folderPath + @"\FirstReleaseDeferred"))
                        {
                            folderPath = mainFolderPath + @"\FirstReleaseDeferred";
                        }
                        else if (Directory.Exists(folderPath + @"\FirstReleaseBusiness"))
                        {
                            folderPath = mainFolderPath + @"\FirstReleaseBusiness";
                        }
                        break;
                }
            }

            if (configXml.Add?.ODTChannel != null)
            {
                switch (configXml.Add.ODTChannel)
                {
                    case ODTChannel.Deferred:
                        if (Directory.Exists(folderPath + @"\DC"))
                        {
                            folderPath = mainFolderPath + @"\DC";
                        }
                        else if (Directory.Exists(folderPath + @"\Deferred"))
                        {
                            folderPath = mainFolderPath + @"\Deferred";
                        }
                        else if (Directory.Exists(folderPath + @"\Business"))
                        {
                            folderPath = mainFolderPath + @"\Business";
                        }
                        break;
                    case ODTChannel.Current:
                        if (Directory.Exists(folderPath + @"\CC"))
                        {
                            folderPath = mainFolderPath + @"\CC";
                        }
                        else if (Directory.Exists(folderPath + @"\Current"))
                        {
                            folderPath = mainFolderPath + @"\Current";
                        }
                        break;
                    case ODTChannel.FirstReleaseCurrent:
                        if (Directory.Exists(folderPath + @"\FRCC"))
                        {
                            folderPath = mainFolderPath + @"\FRCC";
                        }
                        else if (Directory.Exists(folderPath + @"\FirstReleaseCurrent"))
                        {
                            folderPath = mainFolderPath + @"\FirstReleaseCurrent";
                        }
                        break;
                    case ODTChannel.Validation:
                        if (Directory.Exists(folderPath + @"\FRDC"))
                        {
                            folderPath = mainFolderPath + @"\FRDC";
                        }
                        else if (Directory.Exists(folderPath + @"\FirstReleaseDeferred"))
                        {
                            folderPath = mainFolderPath + @"\FirstReleaseDeferred";
                        }
                        else if (Directory.Exists(folderPath + @"\FirstReleaseBusiness"))
                        {
                            folderPath = mainFolderPath + @"\FirstReleaseBusiness";
                        }
                        break;
                    case ODTChannel.FirstReleaseDeferred:
                        if (Directory.Exists(folderPath + @"\FRDC"))
                        {
                            folderPath = mainFolderPath + @"\FRDC";
                        }
                        else if (Directory.Exists(folderPath + @"\FirstReleaseDeferred"))
                        {
                            folderPath = mainFolderPath + @"\FirstReleaseDeferred";
                        }
                        else if (Directory.Exists(folderPath + @"\FirstReleaseBusiness"))
                        {
                            folderPath = mainFolderPath + @"\FirstReleaseBusiness";
                        }
                        break;
                }
            }

            if (Directory.Exists(folderPath + @"\Office\Data"))
            {
                BuildFilePath.Text = folderPath;
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
            await Task.Run(() =>
            {
                var signPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "signtool.exe");
                System.IO.File.WriteAllBytes(signPath, Microsoft.OfficeProPlus.InstallGen.Presentation.Properties.Resources.signtool);

                var thumbprint = GlobalObjects.ViewModel.SelectedCertificate.ThumbPrint;

                path = path.Replace("\\\\", "\\");

                var signProcess = new Process
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = signPath, Arguments = " sign /sha1 " + thumbprint + " " + path, CreateNoWindow = true, UseShellExecute = false
                    }
                };

                signProcess.Start();
            });
        }

        private async Task GenerateInstall(bool sign)
        {
            await Task.Run(async () =>
            {
                try
                {
                    var remoteLogPath = "";
                    if (GlobalObjects.ViewModel.RemoteLoggingPath != null &&
                        !string.IsNullOrEmpty(GlobalObjects.ViewModel.RemoteLoggingPath))
                    {
                        remoteLogPath = GlobalObjects.ViewModel.RemoteLoggingPath;
                    }
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
                            if (executablePath.ToLower().EndsWith(".exe") || executablePath.ToLower().EndsWith(".msi"))
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
                            var result = MessageBox.Show("The directory '" + directoryPath + "' does not exist." + Environment.NewLine + Environment.NewLine + "Create Directory?", "Create Directory", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes);
                            Directory.CreateDirectory(directoryPath);

                            await Dispatcher.InvokeAsync(() => { OpenExeFolderButton.IsEnabled = true; });
                        }
                    }

                    var configFilePath = Environment.ExpandEnvironmentVariables(@"%temp%\OfficeProPlus\" + Guid.NewGuid().ToString() + ".xml");
                    Directory.CreateDirectory(Environment.ExpandEnvironmentVariables(@"%temp%\OfficeProPlus"));

                    GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml.Logging = new ODTLogging
                    {
                        Level = LoggingLevel.Standard, Path = @"%temp%"
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
                    await Dispatcher.InvokeAsync(() => { isInstallExe = InstallExecutable.IsChecked.HasValue && InstallExecutable.IsChecked.Value; });

                    var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
                    string version = null;
                    if (configXml.Add.Version != null)
                    {
                        version = configXml.Add.Version.ToString();
                    }


                    if (!string.IsNullOrEmpty(sourceFilePath))
                    {
                        var branchName = "";
                        if (configXml.Add?.Branch != null)
                        {
                            branchName = configXml.Add.Branch.ToString();
                        }
                        if (configXml.Add?.ODTChannel != null)
                        {
                            branchName = configXml.Add.ODTChannel.ToString();
                        }

                        var languages = new List<string>();
                        foreach (var product in configXml.Add.Products)
                        {
                            foreach (var languageItem in product.Languages)
                            {
                                languages.Add(languageItem.ID);
                            }
                        }

                        var edition = OfficeEdition.Office32Bit;
                        if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office64Bit)
                        {
                            edition = OfficeEdition.Office64Bit;
                        }

                        var ppDownload = new ProPlusDownloader();
                        var validFiles = await ppDownload.ValidateSourceFiles(new DownloadBranchProperties()
                        {
                            TargetDirectory = sourceFilePath,                            
                            BranchName = branchName,
                            Languages = languages,
                            OfficeEdition = edition,
                            Version = version
                        });

                        var cabFilePath = sourceFilePath + @"\Office\Data\v32.cab";
                        if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office64Bit)
                        {
                            cabFilePath = sourceFilePath + @"\Office\Data\v64.cab";
                        }

                        if (string.IsNullOrEmpty(version) && System.IO.File.Exists(cabFilePath))
                        {
                            var fInfo = new FileInfo(cabFilePath);
                            var cabExtractor = new CabExtractor(cabFilePath);
                            cabExtractor.ExtractCabFiles();
                            cabExtractor.Dispose();

                            var vdPathDir = fInfo.Directory?.FullName + @"\ExtractedFiles";
                            var vdPath = vdPathDir + @"\VersionDescriptor.xml";
                            if (System.IO.File.Exists(vdPath))
                            {
                                var latestVersion = ppDownload.GetCabVersion(vdPath);
                                if (latestVersion != null)
                                {
                                    version = latestVersion;
                                }
                                if (Directory.Exists(vdPathDir))
                                {
                                    try
                                    {
                                        Directory.Delete(vdPathDir);
                                    }
                                    catch (Exception ex)
                                    {
                                        var strError = ex.Message;
                                    }
                                }
                            }
                        }

                        if (!validFiles)
                        {
                            throw (new Exception(
                                "The Office Source Files are invalid. Please verify that all of the files have been downloaded."));
                        }
                    }

                    var productName = "Microsoft Office 365 ProPlus Installer";
                    var productId = Guid.NewGuid().ToString(); //"8AA11E8A-A882-45CC-B52C-80149B4CF47A";
                    var upgradeCode = "AC89246F-38A8-4C32-9110-FF73533F417C";

                    var productVersion = new Version("1.0.0");

                    await Dispatcher.InvokeAsync(() =>
                    {
                        if (MajorVersion.Value.HasValue && MinorVersion.Value.HasValue && ReleaseVersion.Value.HasValue)
                        {
                            productVersion =
                                new Version(MajorVersion.Value.Value + "." + MinorVersion.Value.Value + "." +
                                            ReleaseVersion.Value.Value);
                        }
                    });

                    var installProperties = new List<OfficeInstallProperties>();

                    if (GlobalObjects.ViewModel.ApplicationMode == ApplicationMode.LanguagePack)
                    {
                        productName = "Microsoft Office 365 ProPlus Language Pack";

                        var languages = configXml?.Add?.Products?.FirstOrDefault()?.Languages;
                        foreach (var language in languages)
                        {
                            var configLangXml = new ConfigXmlParser(GlobalObjects.ViewModel.ConfigXmlParser.Xml);
                            configLangXml.ConfigurationXml.Add.ODTChannel = null;
                            var tmpProducts = configLangXml?.ConfigurationXml?.Add?.Products;
                            tmpProducts.FirstOrDefault().Languages = new List<ODTLanguage>()
                            {
                                new ODTLanguage()
                                {
                                    ID = language.ID
                                }
                            };

                            var tmpXmlFilePath = Environment.ExpandEnvironmentVariables(@"%temp%\" + Guid.NewGuid().ToString() + ".xml");
                            System.IO.File.WriteAllText(tmpXmlFilePath, configLangXml.Xml);

                            var tmpSourceFilePath = executablePath;

                            if (Regex.Match(executablePath, ".msi$", RegexOptions.IgnoreCase).Success)
                            {
                                tmpSourceFilePath = Regex.Replace(executablePath, ".msi$", "(" + language.ID + ").msi",
                                    RegexOptions.IgnoreCase);
                            }

                            if (Regex.Match(executablePath, ".exe", RegexOptions.IgnoreCase).Success)
                            {
                                tmpSourceFilePath = Regex.Replace(executablePath, ".exe$", "(" + language.ID + ").exe",
                                    RegexOptions.IgnoreCase);
                            }

                            var programFilesPath = @"%ProgramFiles%\Microsoft Office 365 ProPlus Installer\" + language.ID + @"\" + productVersion;

                            var langProductName = productName + " (" + language.ID + ")";

                            installProperties.Add(new OfficeInstallProperties()
                            {
                                ProductName = langProductName,
                                ProductId = langProductName.GenerateGuid(),
                                ConfigurationXmlPath = tmpXmlFilePath,
                                OfficeVersion = OfficeVersion.Office2016,
                                ExecutablePath = tmpSourceFilePath,
                                SourceFilePath = sourceFilePath,
                                BuildVersion = version,
                                UpgradeCode = language.ID.GenerateGuid(),
                                Version = productVersion,
                                Language = "en-us",
                                ProgramFilesPath = programFilesPath,
                                OfficeClientEdition = configXml.Add.OfficeClientEdition
                            });
                        }
                    }
                    else
                    {
                        installProperties.Add(new OfficeInstallProperties()
                        {
                            ProductName = productName,
                            ProductId = productId,
                            ConfigurationXmlPath = configFilePath,
                            OfficeVersion = OfficeVersion.Office2016,
                            ExecutablePath = executablePath,
                            SourceFilePath = sourceFilePath,
                            BuildVersion = version,
                            UpgradeCode = upgradeCode,
                            Version = productVersion,
                            Language = "en-us",
                            ProgramFilesPath = @"%ProgramFiles%\Microsoft Office 365 ProPlus Installer",
                            OfficeClientEdition = configXml.Add.OfficeClientEdition
                        });
                    }

                    foreach (var installProperty in installProperties)
                    {
                        IOfficeInstallGenerator installer = null;
                        if (isInstallExe)
                        {
                            installer = new OfficeInstallExecutableGenerator();
                            LogAnaylytics("/GenerateView", "GenerateExe");
                        }
                        else
                        {
                            installer = new OfficeInstallMsiGenerator();
                            LogAnaylytics("/GenerateView", "GenerateMSI");
                        }
                        installer.Generate(installProperty, remoteLogPath);
                    }


                    await Task.Delay(500);

                    if (!string.IsNullOrEmpty(GlobalObjects.ViewModel.SelectedCertificate.ThumbPrint) && sign)
                    {
                        await InstallerSign(executablePath);
                    }

                    if (InfoMessage != null)
                    {
                        if (isInstallExe)
                        {
                            InfoMessage(this, new MessageEventArgs()
                            {
                                Title = "Generate Executable", Message = "File Generation Complete"
                            });
                        }
                        else
                        {
                            InfoMessage(this, new MessageEventArgs()
                            {
                                Title = "Generate MSI", Message = "File Generation Complete"
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
                    if (configXml.ConfigurationXml.Display.Level.HasValue && configXml.ConfigurationXml.Display.Level == DisplayLevel.None)
                    {
                        if (configXml.ConfigurationXml.Display.AcceptEULA.HasValue && configXml.ConfigurationXml.Display.AcceptEULA == true)
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
                if (GlobalObjects.ViewModel.ApplicationMode == ApplicationMode.LanguagePack)
                {
                    fileName = "OfficeProPlusLanuagePack.msi";
                }
            }
            else
            {
                fileName = "OfficeProPlus.exe";
                if (GlobalObjects.ViewModel.ApplicationMode == ApplicationMode.LanguagePack)
                {
                    fileName = "OfficeProPlusLanuagePack.exe";
                }
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
                    Title = "Error", Message = ex.Message
                });
            }
        }

        private void LogAnaylytics(string path, string pageName)
        {
            try
            {
                GoogleAnalytics.Log(path, pageName);
            }
            catch
            {
            }
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
                var sign = SignInstaller.IsChecked.HasValue && SignInstaller.IsChecked.Value;
                if (sign)
                {
                    if (GlobalObjects.ViewModel.SelectedCertificate == null)
                    {
                        throw (new Exception("Select a Signing Certificate"));
                    }
                }

                await GenerateInstall(sign);
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
                if (SignInstaller.IsChecked.HasValue && SignInstaller.IsChecked.Value)
                {
                    OpenCertificateBrowser.IsEnabled = true;
                    OpenCertGenerator.IsEnabled = true;
                    PublisherRow.Height = new GridLength(50, GridUnitType.Pixel);
                    SpacerRow.Height = new GridLength(64, GridUnitType.Pixel);
                }
                else
                {
                    OpenCertificateBrowser.IsEnabled = false;
                    OpenCertGenerator.IsEnabled = false;
                    PublisherRow.Height = new GridLength(0, GridUnitType.Pixel);
                    SpacerRow.Height = new GridLength(114, GridUnitType.Pixel);
                }
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

                if (SignInstaller.IsChecked.HasValue && SignInstaller.IsChecked.Value)
                {
                    SpacerRow.Height = new GridLength(64, GridUnitType.Pixel);
                }
                else
                {
                    SpacerRow.Height = new GridLength(114, GridUnitType.Pixel);
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
                    Description = "Select a folder to download files to:", ShowNewFolderButton = true, ShowEditBox = true, SelectedPath = BuildFilePath.Text, ShowFullPathInEditBox = true, RootFolder = System.Environment.SpecialFolder.MyComputer
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

                certificatesDialog.Closing += CertificatesDialog_Closing;
                certificatesDialog.Launch();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void CertificatesGenDialog_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                var dialog = (GenerateCertificate) sender;
                if (dialog.Result == DialogResult.OK)
                {
                    Publisher.Text = GlobalObjects.ViewModel.SelectedCertificate.FriendlyName;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void CertificatesDialog_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                var dialog = (CertificatesDialog) sender;
                if (dialog.Result == DialogResult.OK)
                {
                    Publisher.Text = GlobalObjects.ViewModel.SelectedCertificate.FriendlyName;
                }
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
                generateCertificateDialog.Closing += CertificatesGenDialog_Closing;
                generateCertificateDialog.Launch();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }
        private void SilentInstallInfo_Click(object sender, RoutedEventArgs e)
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

        private void SignInstallerInfo_Click(object sender, RoutedEventArgs e)
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

        private void VersionInfoGen_Click(object sender, RoutedEventArgs e)
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

        private void FilePathGen_Click(object sender, RoutedEventArgs e)
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

        private void SourceFilePathGen_Click(object sender, RoutedEventArgs e)
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

        #region "Info"

        private void ProductInfo_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic) sender).Name;
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
                        Height = 500, Width = 400
                    };
                    informationDialog.Closed += (o, args) => { informationDialog = null; };
                    informationDialog.Closing += (o, args) => { };
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

        private void xmlBrowser_Loaded(object sender, RoutedEventArgs e)
        {
        }

        
    }
}
