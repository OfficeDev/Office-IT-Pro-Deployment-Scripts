using System;
using System.Collections.Generic;
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




            switch(GlobalObjects.ViewModel.SccmConfiguration.Scenario)
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


        private void CbDeploymentPurpose_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //throw new NotImplementedException();
        }

        private void CbDeploymentType_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //throw new NotImplementedException();
        }

        private void DistributionPoint_OnTextChanged(object sender, TextChangedEventArgs e)
        {
           
            DistributionPointGroupName.IsEnabled = false;

            if (DistributionPoint.Text.Length == 0)
                DistributionPointGroupName.IsEnabled = true;

        }

        private void DistributionPointGroupName_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            DistributionPoint.IsEnabled = false;

            if (DistributionPointGroupName.Text.Length == 0)
                DistributionPoint.IsEnabled = true;
        }
    }
}
