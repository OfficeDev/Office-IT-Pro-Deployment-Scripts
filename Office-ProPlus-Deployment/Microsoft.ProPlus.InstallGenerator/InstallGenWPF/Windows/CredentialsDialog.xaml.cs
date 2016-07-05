using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
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
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class CredentialsDialog : IDisposable
    {
        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public void Launch()
        {
            this.Show();
            string username = GlobalObjects.ViewModel.GetUsername();
            string password = GlobalObjects.ViewModel.GetPassword();
            string domain = GlobalObjects.ViewModel.GetDomain();
            if ((!string.IsNullOrEmpty(username)) && (!string.IsNullOrEmpty(password)))
            {
                txtBoxUserName.Text = username;
                txtBoxPassword.Password = password;
                txtBoxDomain.Text = domain;
            }
        }

        public CredentialsDialog()
        {
            InitializeComponent();
        }

        public DialogResult Result = System.Windows.Forms.DialogResult.Cancel;

        private void OkButton_OnClick(object sender, RoutedEventArgs e)
        {
            try { 
            if ((!string.IsNullOrEmpty(txtBoxUserName.Text)) && (!string.IsNullOrEmpty(txtBoxPassword.Password)))
            {
                Result = System.Windows.Forms.DialogResult.OK;
                GlobalObjects.ViewModel.SetCredentials(txtBoxUserName.Text, txtBoxPassword.Password, txtBoxDomain.Text);
            }
            else
            {
                throw  new Exception("Please provide username and password");
            }

            this.Close();
        }
        catch (Exception ex)
        {
            ex.LogException();
        }
    }


    private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            Result = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }
    }
}
