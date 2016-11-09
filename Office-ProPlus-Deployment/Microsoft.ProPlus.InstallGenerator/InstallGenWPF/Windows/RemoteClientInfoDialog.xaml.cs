using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
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
                    if(txtBxAddMachines.Text.Contains(" "))
                        throw new Exception("Please add IP addresses or machine names only");
                    GlobalObjects.ViewModel.RemoteConnectionInfo(txtBxAddMachines.Text.Trim());
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
        

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        private void ImportComputers_OnClick_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".png",
                Filter = "CSV Files (.csv)|*.csv"
            };

            var result = dlg.ShowDialog();
            if (result == true)
            {
                string line;
                txtBxAddMachines.Text = "";
                try
                {

                    StreamReader file = new StreamReader(dlg.FileName);

                    while ((line = file.ReadLine()) != null)
                    {
                        txtBxAddMachines.Text += line + Environment.NewLine;
                    }
                    txtBxAddMachines.Text = txtBxAddMachines.Text.TrimEnd();
                }
                catch (Exception ex)
                {

                }
            }
        }

        private void AddCredentials_OnClick(object sender, RoutedEventArgs e)
        {
            var credentialsDialogue = new CredentialsDialog();
            credentialsDialogue.Closing += CredentialsDialogue_Closing;
            
            credentialsDialogue.Launch();
        }

        private void CredentialsDialogue_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //do stuff here, I guess
        }
    }
}
