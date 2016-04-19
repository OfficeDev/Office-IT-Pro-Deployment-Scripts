using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.VisualStyles;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MetroDemo.Events;
using MetroDemo.Models;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGenerator.Implementation;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class StartView : UserControl
    {

        public event MessageEventHandler InfoMessage;
        public event MessageEventHandler ErrorMessage;
        public event TransitionTabEventHandler TransitionTab;

        private bool _running = false;

        public StartView()
        {
            InitializeComponent();
        }

        private void StartView_OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (cbActions.Items.Count < 1)
                {
                    cbActions.Items.Add("Create new Office 365 ProPlus installer");
                    cbActions.Items.Add("Import an existing Office 365 ProPlus installer");
                    cbActions.Items.Add("Manage your local Office 365 ProPlus installation");
                    cbActions.Items.Add("Create Office 365 ProPlus language pack");
                    cbActions.SelectedIndex = 0;
                }
                
                LogAnaylytics("/", "StartView");
            }
            catch (Exception ex)
            {
                ex.LogException();
            }
        }

        private void StartNew()
        {
            try
            {
                if (_running) return;
                GlobalObjects.ViewModel.LocalConfig = false;
                GlobalObjects.ViewModel.RunLocalConfigs = false;

                GlobalObjects.ViewModel.ConfigXmlParser.LoadXml(GlobalObjects.DefaultXml);
                GlobalObjects.ViewModel.ResetXml = true;
                GlobalObjects.ViewModel.ImportFile = null;

                if (RestartWorkflow != null)
                {
                    this.RestartWorkflow(this, new EventArgs());
                }

                this.TransitionTab(this, new TransitionTabEventArgs()
                {
                    Direction = TransitionTabDirection.Forward,
                    Index = 0
                });

                LogAnaylytics("/StartView", "StartNew");

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void ImportExisting()
        {
            try
            {
                if (_running) return;
                GlobalObjects.ViewModel.LocalConfig = false;
                GlobalObjects.ViewModel.RunLocalConfigs = false;

                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    DefaultExt = ".png",
                    Filter =
                        "Generated Files (*.xml,*.exe,*.msi)|*.xml;*.exe;*.msi|XML Configuation (*.xml)|*.xml|Executable (*.exe)|*.exe|MSI Installer (*.msi)|*.msi"
                };

                var result = dlg.ShowDialog();

                if (result == true)
                {
                    var filename = dlg.FileName;

                    GlobalObjects.ViewModel.ResetXml = true;

                    var configExtractor = new OfficeConfigXmlExtractor();
                    GlobalObjects.ViewModel.ImportFile = filename;

                    filename = configExtractor.ExtractXml(filename);

                    if (RestartWorkflow != null)
                    {
                        this.RestartWorkflow(this, new EventArgs());
                    }

                    GlobalObjects.ViewModel.ConfigXmlParser.LoadXml(filename);

                    if (this.XmlImported != null)
                    {
                        this.XmlImported(this, new EventArgs());
                    }

                    this.TransitionTab(this, new TransitionTabEventArgs()
                    {
                        Direction = TransitionTabDirection.Forward,
                        Index = 0
                    });

                    LogAnaylytics("/StartView", "ImportExisting");
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void ManageLocal()
        {
            try
            {
                if (_running) return;

                GlobalObjects.ViewModel.LocalConfig = true;
                GlobalObjects.ViewModel.BlockNavigation = true;
                _running = true;
                var localXml = "";

                await Task.Run(async () => {
                    Dispatcher.Invoke(() =>
                    {
                        WaitManageLocal.Visibility = Visibility.Visible;
                        //ImgManageLocal.Visibility = Visibility.Collapsed;
                    });

                    GlobalObjects.ViewModel.RunLocalConfigs = true;

                    var officeInstallManager = new OfficeLocalInstallManager();
                    localXml = await officeInstallManager.GenerateLocalConfigXml();
                });

                GlobalObjects.ViewModel.ConfigXmlParser.LoadXml(localXml);
                GlobalObjects.ViewModel.ResetXml = true;
                GlobalObjects.ViewModel.ImportFile = null;

                GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml.Add.Version = null;

                if (RestartWorkflow != null)
                {
                    this.RestartWorkflow(this, new EventArgs());
                }

                GlobalObjects.ViewModel.BlockNavigation = false;

                var installOffice = new InstallOffice();
                if (installOffice.IsUpdateRunning())
                {
                    this.TransitionTab(this, new TransitionTabEventArgs()
                    {
                        Direction = TransitionTabDirection.Forward,
                        Index = 6,
                        UseIndex = true
                    });
                }
                else
                {
                    this.TransitionTab(this, new TransitionTabEventArgs()
                    {
                        Direction = TransitionTabDirection.Forward,
                        Index = 0
                    });
                }

                LogAnaylytics("/StartView", "StartNew");
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
            finally
            {
                Dispatcher.Invoke(() =>
                {
                    WaitManageLocal.Visibility = Visibility.Collapsed;
                    //ImgManageLocal.Visibility = Visibility.Visible;
                });
                _running = false;
            }
        }

        private void CreateLanguagePack()
        {
            try
            {
                if (_running) return;
                GlobalObjects.ViewModel.LocalConfig = false;
                GlobalObjects.ViewModel.RunLocalConfigs = false;

                GlobalObjects.ViewModel.ConfigXmlParser.LoadXml(GlobalObjects.DefaultXml);
                GlobalObjects.ViewModel.ResetXml = true;
                GlobalObjects.ViewModel.ImportFile = null;

                if (RestartWorkflow != null)
                {
                    this.RestartWorkflow(this, new EventArgs());
                }

                this.TransitionTab(this, new TransitionTabEventArgs()
                {
                    Direction = TransitionTabDirection.Forward,
                    Index = 7
                });
                LogAnaylytics("/StartView", "LanguagePack");


            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
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

        public RestartEventHandler RestartWorkflow  { get; set; }

        public XmlImportedEventHandler XmlImported { get; set; }

        private void LogAnaylytics(string path, string pageName)
        {
            try
            {
                GoogleAnalytics.Log(path, pageName);
            }
            catch { }
        }

        private void strtButton_Click(object sender, RoutedEventArgs e)
        {
            switch (cbActions.SelectedIndex)
            {
                case 0:
                    StartNew();
                    break;
                case 1:
                    ImportExisting();
                    break;
                case 2:
                    ManageLocal();
                    break;
                case 3:
                    CreateLanguagePack();
                    break;
                default:
                    LogErrorMessage(new Exception("invalid selection"));
                    break;
            }
        }

        private void cbActions_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (cbActions.SelectedIndex)
            {
                case 0:
                    txtBlock.Text = "Select this option if you would like to start a new or reset an installation.";
                    break;
                case 1:
                    txtBlock.Text = "Select this option if you have an existing Configuration XML or an Executable or MSI that was generated by this application";
                    break;
                case 2:
                    txtBlock.Text = "Select this option if ou would like to install, modify or manage the local installation of Office 365 ProPlus.";
                    break;
                case 3:
                    txtBlock.Text = "Select this option if you would like to create a language pack.";
                    break;
                default:
                    txtBlock.Text = "";
                    break;
            }
        }    
    }
}
