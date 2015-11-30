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

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class StartView : UserControl
    {

        public StartView()
        {
            InitializeComponent();
        }
        
        public event TransitionTabEventHandler TransitionTab;

        private void StartNew_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                GlobalObjects.ViewModel.ConfigXmlParser.LoadXml(GlobalObjects.ViewModel.DefaultXml);
                GlobalObjects.ViewModel.ResetXml = true;

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
                MessageBox.Show("ERROR: " + ex.Message);
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
                        "XML Configuation (*.xml)|*.xml|Executable (*.exe)|*.exe|MSI Installer (*.msi)|*.msi"
                };

                var result = dlg.ShowDialog();

                if (result == true)
                {
                    var filename = dlg.FileName;

                    GlobalObjects.ViewModel.ResetXml = true;

                    if (filename.ToLower().EndsWith(".exe"))
                    {
                        filename = ExtractXmlFromExecutable(filename);
                    }

                    GlobalObjects.ViewModel.ConfigXmlParser.LoadXml(filename);

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
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private string ExtractXmlFromExecutable(string fileName)
        {
            var tmpDir = Environment.ExpandEnvironmentVariables("%temp%");

            var p = new Process
            {
                StartInfo = new ProcessStartInfo()
                {
                    FileName = fileName,
                    Arguments = "/extractxml=" + tmpDir + @"\configuration.xml",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                },
            };
            p.Start();
            p.WaitForExit();

            var xml = File.ReadAllText(tmpDir + @"\configuration.xml");
            return xml;
        }

        public RestartEventHandler RestartWorkflow  { get; set; }

    }
}
