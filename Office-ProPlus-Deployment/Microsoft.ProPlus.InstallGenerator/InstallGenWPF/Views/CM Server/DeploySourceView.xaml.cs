using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
using MetroDemo;
using MetroDemo.Events;
using MetroDemo.ExampleViews;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.Win32;
using Button = System.Windows.Controls.Button;
using RadioButton = System.Windows.Controls.RadioButton;
using TextBox = System.Windows.Controls.TextBox;
using UserControl = System.Windows.Controls.UserControl;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config
{
    /// <summary>
    /// Interaction logic for DeploySource.xaml
    /// </summary>
    public partial class DeploySourceView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;

        public DeploySourceView()
        {
            InitializeComponent();
        }

        private void DeploySourceView_OnLoaded(object sender, RoutedEventArgs e)
        {

        }

        private void rdb_SourceChanged(object sender, RoutedEventArgs e)
        {
            var source = (RadioButton) sender;
            var parents = (MainWindow)Window.GetWindow(this);
            

            if (source.Content.ToString() == "CDN")
            {
                GlobalObjects.ViewModel.CmPackage.DeploymentSource = DeploymentSource.CDN;

                if (BrowseButton != null)
                {
                    BrowseButton.IsEnabled = false;
                    FilePath.IsEnabled = false;
                }

                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = true
                });
            }
            else if (source.Content.ToString() == "Distribution Point")
            {
                GlobalObjects.ViewModel.CmPackage.DeploymentSource = DeploymentSource.DistributionPoint;
                BrowseButton.IsEnabled = false;
                FilePath.IsEnabled = false;

                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = true
                });
            }
            else
            {
                GlobalObjects.ViewModel.CmPackage.DeploymentSource = DeploymentSource.Local;
                BrowseButton.IsEnabled = true;
                FilePath.IsEnabled = true;

                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = false
                });

            }

        }

        private void FilePath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var nextButton = (Button)this.FindName("NextButton");
            var textBox = (TextBox) sender;
            var filePath = textBox.Text;

            if (Directory.Exists(filePath))
            {
                GlobalObjects.ViewModel.CmPackage.DeploymentDirectory = filePath;

                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = true
                });
            }
            else
            {
                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = false
                });
            }
        }

        private void BrowseButton_OnClick(object sender, RoutedEventArgs e)
        {
            var filepath = FilePath.Text;
            var fileBrowser = new FolderBrowserDialog();

            if (filepath != null && Directory.Exists(filepath))
            {
                fileBrowser.SelectedPath = filepath;
            }
            fileBrowser.ShowDialog();
            FilePath.Text = fileBrowser.SelectedPath;
        }

        private void DownloadPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var grid = (Grid) sender;

            if (grid.Visibility == Visibility.Visible)
            {
                if (FilePath.IsEnabled &&
                    !Directory.Exists(GlobalObjects.ViewModel.CmPackage.DeploymentDirectory))
                {
                    ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                    {
                        Enabled = false
                    });
                }
                else
                {
                    ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                    {
                        Enabled = true
                    });
                }
            }
        }
    }
}