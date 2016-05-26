using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace MetroDemo.ExampleWindows
{
    /// <summary>
    /// Interaction logic for RemoteChannelVersionDialog.xaml
    /// </summary>
    public partial class RemoteMachinesDialog : IDisposable
    {
        public RemoteMachinesDialog()
        {
            InitializeComponent();
        }
        public DialogResult Result = System.Windows.Forms.DialogResult.Cancel;
        private List<Channel> items;

        public void Launch()
        {
            
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Result = System.Windows.Forms.DialogResult.OK;
            
                this.Close();                
            }
            catch (Exception ex)
            {
                ex.LogException();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Result = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

    }
}
