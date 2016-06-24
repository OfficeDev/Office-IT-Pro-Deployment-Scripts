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
    public partial class RemoteClientInfoDialog : IDisposable
    {
        public RemoteClientInfoDialog()
        {
            InitializeComponent();
        }
        public DialogResult Result = System.Windows.Forms.DialogResult.Cancel;

        public void Launch()
        {
           
            this.Show();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if(!string.IsNullOrEmpty(txtBxAddMachines.Text))
                {
                    Result = System.Windows.Forms.DialogResult.OK;
                    GlobalObjects.ViewModel.RemoteConnectionInfo(txtBxAddMachines.Text);
                }
              
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

        private List<String> getVersions(OfficeBranch currentChannel, List<String> versions, string currentVersion)
        {

            foreach (var version in currentChannel.Versions)
            {
                if (version.Version.ToString() != currentVersion)
                {
                    versions.Add(version.Version.ToString());
                }
            }

            return versions;
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

    }
}
