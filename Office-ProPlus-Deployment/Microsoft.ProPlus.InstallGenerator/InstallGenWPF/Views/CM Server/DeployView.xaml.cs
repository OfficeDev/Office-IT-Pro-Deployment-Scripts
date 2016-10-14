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
using MetroDemo;
using MetroDemo.Events;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config

{
    /// <summary>
    /// Interaction logic for DeployView.xaml
    /// </summary>
    public partial class DeployView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;


        public DeployView()
        {
            InitializeComponent();
        }

        private void DeployView_OnLoaded(object sender, RoutedEventArgs e)
        {

        }

        private void DeployPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
           
        }

        private void DeployButton_OnClick(object sender, RoutedEventArgs e)
        {
            
        }

        #region helpers

        private async Task GetScripts()
        {
            
        }

        #endregion
    }
}
