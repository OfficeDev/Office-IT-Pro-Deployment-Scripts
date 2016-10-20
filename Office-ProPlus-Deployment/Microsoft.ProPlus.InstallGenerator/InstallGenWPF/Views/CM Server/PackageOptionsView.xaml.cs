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
using MahApps.Metro.Controls;
using MetroDemo;
using MetroDemo.Events;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using TextBox = System.Windows.Controls.TextBox;
using UserControl = System.Windows.Controls.UserControl;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config
{
    /// <summary>
    /// Interaction logic for DeployOtherView.xaml
    /// </summary>
    public partial class PackageOptionsView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;
        private CmProgram CurrentCmProgram = GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1]; 


        public PackageOptionsView()
        {
            InitializeComponent();
        }

      

        private void PackageOptionsView_OnLoaded(object sender, RoutedEventArgs e)
        {
           
        }

        private void PackageOptionsPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            CurrentCmProgram =
                    GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

            if (GlobalObjects.ViewModel.CmPackage.DeploymentSource == DeploymentSource.DistributionPoint)
            {
                lblDistributionPoint.IsEnabled = true;
                lblDistributionPointGroupName.IsEnabled = true;

                if (DistributionPoint.Text.Length == 0 && DistributionPointGroupName.Text.Length > 0)
                {
                    DistributionPoint.IsEnabled = false;
                    DistributionPointGroupName.IsEnabled = true;
                }
                else if (DistributionPoint.Text.Length > 0 && DistributionPointGroupName.Text.Length == 0)
                {
                    DistributionPoint.IsEnabled = true;
                    DistributionPointGroupName.IsEnabled = false;
                }
                else
                {
                    DistributionPoint.IsEnabled = true;
                    DistributionPointGroupName.IsEnabled = true;

                }
               
            }
            else
            {
                lblDistributionPoint.IsEnabled = false;
                lblDistributionPointGroupName.IsEnabled = false;
                DistributionPoint.IsEnabled = false;
                DistributionPointGroupName.IsEnabled = false;
            }

            ToggleNext();
        }

        private void ToggleNext()
        {

            if (GlobalObjects.ViewModel.CmPackage.DeploymentSource == DeploymentSource.DistributionPoint)
            {
                if (GlobalObjects.ViewModel.CmPackage.DistributionPointGroupName.Length > 0 || 
                    GlobalObjects.ViewModel.CmPackage.DistributionPoint.Length > 0)
                {
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
            else
            {
                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = true
                });    
            }
           
        }

        private void DistributionPoint_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;
            GlobalObjects.ViewModel.CmPackage.DistributionPoint = text;

            DistributionPointGroupName.IsEnabled = false;

            if (DistributionPoint.Text.Length == 0)
                DistributionPointGroupName.IsEnabled = true;

            ToggleNext();
        }

        private void DistributionPointGroupName_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;
            GlobalObjects.ViewModel.CmPackage.DistributionPointGroupName = text;

            DistributionPoint.IsEnabled = false;

            if (DistributionPointGroupName.Text.Length == 0)
                DistributionPoint.IsEnabled = true;

            ToggleNext();
        }



        private void DeploymentExpiryDurationInDays_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            if (text.All(char.IsDigit))
                GlobalObjects.ViewModel.CmPackage.DeploymentExpiryDurationInDays = Convert.ToInt32(text);
            else
            {
                textbox.Text = GlobalObjects.ViewModel.CmPackage.DeploymentExpiryDurationInDays.ToString();
            }
        }

        private void PackageName_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            GlobalObjects.ViewModel.CmPackage.CustomPackageShareName = text;
        }

        private void CMPSModulePath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            if (Directory.Exists(text))
                GlobalObjects.ViewModel.CmPackage.CMPSModulePath = text;
        }

        private void TsMoveFiles_OnChecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.CmPackage.MoveFiles = true;

        }

        private void TsMoveFiles_OnUnchecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.CmPackage.MoveFiles = false;

        }

        private void TsUpdateBits_OnChecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.CmPackage.UpdateOnlyChangedBits = true;
        }

        private void TsUpdateBits_OnUnchecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.CmPackage.UpdateOnlyChangedBits = false;
        }

        private void BrowseButton_OnClick(object sender, RoutedEventArgs e)
        {
            var filepath = CMPSModulePath.Text;
            var fileBrowser = new FolderBrowserDialog();

            if (filepath != null && Directory.Exists(filepath))
            {
                fileBrowser.SelectedPath = filepath;
            }
            fileBrowser.ShowDialog();
            CMPSModulePath.Text = fileBrowser.SelectedPath;
        }
    }
}
