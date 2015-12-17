using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MetroDemo.Events;
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

        public StartView()
        {
            InitializeComponent();
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

        private void StartNew_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                GlobalObjects.ViewModel.ConfigXmlParser.LoadXml(GlobalObjects.ViewModel.DefaultXml);
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
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void ImportExisting_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
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
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        public RestartEventHandler RestartWorkflow  { get; set; }

        public XmlImportedEventHandler XmlImported { get; set; }

    }
}
