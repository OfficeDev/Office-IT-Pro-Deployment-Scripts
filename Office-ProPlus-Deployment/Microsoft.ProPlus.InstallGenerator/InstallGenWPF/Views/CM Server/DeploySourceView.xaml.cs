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

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config
{
    /// <summary>
    /// Interaction logic for DeploySource.xaml
    /// </summary>
    public partial class DeploySourceView : UserControl
    {
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
            
        }
    }
}
