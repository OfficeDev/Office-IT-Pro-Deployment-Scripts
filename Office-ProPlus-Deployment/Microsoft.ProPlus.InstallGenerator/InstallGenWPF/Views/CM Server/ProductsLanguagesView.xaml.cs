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
using MetroDemo.Events;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config
{
    /// <summary>
    /// Interaction logic for ProductsLanguagesView.xaml
    /// </summary>
    public partial class ProductsLanguagesView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;

        public ProductsLanguagesView()
        {
            InitializeComponent();
        }

        private void ProductsLanguagesView_OnLoaded(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}
