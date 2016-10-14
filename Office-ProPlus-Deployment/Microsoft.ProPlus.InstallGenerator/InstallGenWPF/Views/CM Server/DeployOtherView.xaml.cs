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

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config
{
    /// <summary>
    /// Interaction logic for DeployOtherView.xaml
    /// </summary>
    public partial class DeployOtherView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;


        public DeployOtherView()
        {
            InitializeComponent();
        }

      

        private void DeployOtherView_OnLoaded(object sender, RoutedEventArgs e)
        {
           
        }

        private void OtherOptionsPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {

            ToggleNext();


            switch (GlobalObjects.ViewModel.SccmConfiguration.Scenario)
            {
                case SccmScenario.Deploy:

                    lblDistributionPoint.Visibility = Visibility.Visible;
                    lblDistributionPointGroupName.Visibility = Visibility.Visible;
                    DistributionPoint.Visibility = Visibility.Visible;
                    DistributionPointGroupName.Visibility = Visibility.Visible;

                    lblDeploymentExpiryDurationInDays.SetValue(Grid.RowProperty, 4);
                    DeploymentExpiryDurationInDays.SetValue(Grid.RowProperty, 4);
                    lblPackName.SetValue(Grid.RowProperty, 5);
                    PackageName.SetValue(Grid.RowProperty, 5);


                    if (GlobalObjects.ViewModel.SccmConfiguration.DeploymentSource != DeploymentSource.DistributionPoint)
                    {
                        lblDeploymentExpiryDurationInDays.SetValue(Grid.RowProperty, 2);
                        DeploymentExpiryDurationInDays.SetValue(Grid.RowProperty, 2);
                        lblPackName.SetValue(Grid.RowProperty, 3);
                        PackageName.SetValue(Grid.RowProperty, 3);

                        lblDistributionPoint.Visibility = Visibility.Collapsed;
                        lblDistributionPointGroupName.Visibility = Visibility.Collapsed;
                        DistributionPoint.Visibility = Visibility.Collapsed;
                        DistributionPointGroupName.Visibility = Visibility.Collapsed;


                    }
                    break;
                case SccmScenario.ChangeChannel:
                    break;
                case SccmScenario.Rollback:
                    break;
                case SccmScenario.UpdateConfigMgr:
                    break;
                case SccmScenario.UpdateScheduledTask:
                    break;
                default:
                    break;
            }


        }

        private void BrowseButton_OnClick(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }


        private void DistributionPoint_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;
            GlobalObjects.ViewModel.SccmConfiguration.DistributionPoint = text;

            DistributionPointGroupName.IsEnabled = false;

            if (DistributionPoint.Text.Length == 0)
                DistributionPointGroupName.IsEnabled = true;

            ToggleNext();
        }

        private void DistributionPointGroupName_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox) sender;
            var text = textbox.Text;
            GlobalObjects.ViewModel.SccmConfiguration.DistributionPointGroupName = text;

            DistributionPoint.IsEnabled = false;

            if (DistributionPointGroupName.Text.Length == 0)
                DistributionPoint.IsEnabled = true;

            ToggleNext();
        }

        private void ToggleNext()
        {
            var SccmConfig = GlobalObjects.ViewModel.SccmConfiguration;

            switch (GlobalObjects.ViewModel.SccmConfiguration.Scenario)
            {
                case SccmScenario.Deploy:

                    if (SccmConfig.DeploymentSource != DeploymentSource.DistributionPoint)
                    {
                        if (Collection.Text.Length > 0)
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
                        if (Collection.Text.Length > 0 &&
                            (DistributionPointGroupName.Text.Length > 0 || DistributionPoint.Text.Length > 0))
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


                    break;
                case SccmScenario.ChangeChannel:
                    break;
                case SccmScenario.Rollback:
                    break;
                case SccmScenario.UpdateConfigMgr:
                    break;
                case SccmScenario.UpdateScheduledTask:
                    break;
                default:
                    break;
            }
        }

        private void DeploymentExpiryDurationInDays_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            if (text.All(char.IsDigit))
                GlobalObjects.ViewModel.SccmConfiguration.DeploymentExpiryDurationInDays = Convert.ToInt32(text);
            else
            {
                textbox.Text = GlobalObjects.ViewModel.SccmConfiguration.DeploymentExpiryDurationInDays.ToString();
            }
        }

        private void PackageName_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            GlobalObjects.ViewModel.SccmConfiguration.CustomPackageShareName = text; 
        }

        private void CMPSModulePath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            if(Directory.Exists(text))
                GlobalObjects.ViewModel.SccmConfiguration.CMPSModulePath = text;
        }

        private void SiteCode_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            GlobalObjects.ViewModel.SccmConfiguration.SiteCode = text;
        }

        private void ScriptName_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            GlobalObjects.ViewModel.SccmConfiguration.ScriptName = text; 
        }

        private void ConfigurationXml_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            GlobalObjects.ViewModel.SccmConfiguration.ConfigurationXml = text;
        }

        private void CustomName_OnTextChanged(object sender, TextChangedEventArgs e)
        {

            var textbox = (TextBox)sender;
            var text = textbox.Text;

            GlobalObjects.ViewModel.SccmConfiguration.CustomName = text;
        }

        private void Collection_OnTextChanged(object sender, TextChangedEventArgs e)
        {

            var textbox = (TextBox)sender;
            var text = textbox.Text;

            GlobalObjects.ViewModel.SccmConfiguration.Collection = text;

            ToggleNext();
        }

        private void CbDeploymentPurpose_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = (ComboBox) sender;
            var index = comboBox.SelectedIndex;
            var value = GlobalObjects.ViewModel.DeploymentPurposes[index];

            GlobalObjects.ViewModel.SccmConfiguration.DeploymentPurpose = value; 
        }

        private void CbDeploymentType_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = (ComboBox)sender;
            var index = comboBox.SelectedIndex;
            var value = GlobalObjects.ViewModel.DeploymentTypes[index];

            GlobalObjects.ViewModel.SccmConfiguration.DeploymentType = value;
        }

        private void TsMoveFiles_OnChecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.SccmConfiguration.MoveFiles = true;

        }

        private void TsMoveFiles_OnUnchecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.SccmConfiguration.MoveFiles = false;

        }

        private void TsUpdateBits_OnChecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.SccmConfiguration.UpdateOnlyChangedBits = true;
        }

        private void TsUpdateBits_OnUnchecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.SccmConfiguration.UpdateOnlyChangedBits = false;
        }
    }
}
