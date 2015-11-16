using System;
using System.Collections.Generic;
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
        public MainWindowViewModel ViewModel { get; set; }

        public StartView()
        {
            InitializeComponent();
        }
        
        public event TransitionTabEventHandler TransitionTab;

        private void StartNew_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ViewModel.ConfigXmlParser.LoadXml("<Configuration></Configuration>");

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

                    ViewModel.ConfigXmlParser.LoadXml(filename);

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

    }
}
