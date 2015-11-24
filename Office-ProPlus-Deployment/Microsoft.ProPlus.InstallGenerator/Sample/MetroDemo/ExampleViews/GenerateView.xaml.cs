using System;
using System.Collections.Generic;
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
using MetroDemo.ExampleWindows;
using Microsoft.Win32;

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
            }
            catch (Exception ex)
            {
                
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



    }
}
