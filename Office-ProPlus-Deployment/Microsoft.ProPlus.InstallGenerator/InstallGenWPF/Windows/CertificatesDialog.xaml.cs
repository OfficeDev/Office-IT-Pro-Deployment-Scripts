using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using MetroDemo.Models;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using Application = System.Windows.Application;
using MessageBox = System.Windows.Forms.MessageBox;

namespace MetroDemo.ExampleWindows
{
    public partial class CertificatesDialog : IDisposable
    {

        public List<Certificate> Certificatesource { get; set; } 

        private bool _disposed;
        private bool _hideOnClose = true;

        public CertificatesDialog()
        {
            this.DataContext = GlobalObjects.ViewModel;
            this.InitializeComponent();
            this.Closing += (s, e) =>
                {
                    if (_hideOnClose)
                    {
                        Hide();
                        e.Cancel = true;
                    }
                };
            
            var mainWindow = (MetroWindow)this;
            var windowPlacementSettings = mainWindow.GetWindowPlacementSettings();
            if (windowPlacementSettings.UpgradeSettings)
            {
                windowPlacementSettings.Upgrade();
                windowPlacementSettings.UpgradeSettings = false;
                windowPlacementSettings.Save();
            }

        }


        public void Launch()
        {
            Owner = Application.Current.MainWindow;
            // only for this window, because we allow minimizing
            if (WindowState == WindowState.Minimized)
            {
                WindowState = WindowState.Normal;
            }
            Show();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (CertificateList.SelectedItems.Count > 0)
                {
                    var tempCert = (Certificate) CertificateList.SelectedItem;
                    if (tempCert != null)
                    {
                        GlobalObjects.ViewModel.SelectedCertificate = tempCert;
                        GlobalObjects.ViewModel.SelectedCertificate.FriendlyName = tempCert.FriendlyName;
                        GlobalObjects.ViewModel.SelectedCertificate.IssuerName = tempCert.IssuerName;
                        GlobalObjects.ViewModel.SelectedCertificate.ThumbPrint = tempCert.ThumbPrint;
                    }
                }
                else
                {
                    GlobalObjects.ViewModel.SelectedCertificate = new Certificate();
                }

                Result = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                ex.LogException();
            }
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.SelectedCertificate = new Certificate();
            Result = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        public DialogResult Result = System.Windows.Forms.DialogResult.Cancel;


        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                _disposed = true;
                _hideOnClose = false;
                Close();
            }
        }


    }
}
