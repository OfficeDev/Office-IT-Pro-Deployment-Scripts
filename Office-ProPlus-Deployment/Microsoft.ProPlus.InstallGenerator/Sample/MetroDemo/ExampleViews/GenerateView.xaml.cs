using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
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

        public event TransitionTabEventHandler TransitionTab;

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



    }
}
