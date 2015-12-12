using System.Windows;
using System.Windows.Controls;

namespace MahApps.Metro.Controls.XmlBrowser
{
    /// <summary>
    /// Interaction logic for ValidationResultControl.xaml
    /// </summary>
    public partial class ValidationResultControl : UserControl
    {
        public ValidationResultControl()
        {
            InitializeComponent();
        }

        private void ButtonClick(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(ResultTextBox.Text);
        }

        private void ButtonClick1(object sender, RoutedEventArgs e)
        {
            ((Window)this.Parent).Close();
        }
    }
}
