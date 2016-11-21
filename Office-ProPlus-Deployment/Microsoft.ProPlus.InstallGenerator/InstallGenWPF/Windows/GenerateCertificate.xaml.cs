using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
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
    public partial class GenerateCertificate : IDisposable
    {

        public List<Certificate> Certificatesource { get; set; } 

        private bool _disposed;
        private bool _hideOnClose = true;

        public GenerateCertificate()
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

        private X509Certificate2 GetThumbPrint( int serialNumber)
        {
            var localStore = new X509Store(StoreLocation.CurrentUser);
            var thumbprint = "";
            try
            {
                localStore.Open(OpenFlags.ReadOnly);
                if (localStore.Certificates.Count > 0)
                {
                    foreach (var certificate in localStore.Certificates)
                    {
                        var currentSerialNumber = certificate.SerialNumber;
                        var matchSerialNumber = serialNumber.ToString("X6");

                        if (currentSerialNumber == matchSerialNumber)
                        {
                            return certificate;
                        }
                    }
                }
                return new X509Certificate2();
            }
            finally
            {
                localStore.Close();
            }
        }

        private Certificate CreateCertificate(string publisher)
        {
            try
            {
                var getRandom = new Random();
                var makeCertPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "makecert.exe");
                var startDate = "01/01/" + DateTime.Now.Year.ToString();
                var endDate = "01/01/" + DateTime.Now.AddYears(2).Year.ToString(); 
                var serialNumber = getRandom.Next(0, 1000000);

                System.IO.File.WriteAllBytes(makeCertPath, Microsoft.OfficeProPlus.InstallGen.Presentation.Properties.Resources.makecert);

                var createProcess = new Process
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = makeCertPath,
                        Arguments =
                            " -r -pe -n \"CN=" + publisher + "\" -b " + startDate + " -e " + endDate +
                            " -eku 1.3.6.1.5.5.7.3.3 -ss My -# " + serialNumber,
                        CreateNoWindow = true,
                        UseShellExecute = false
                    }
                };

                createProcess.Start();
                createProcess.WaitForExit();
                var cert = GetThumbPrint(serialNumber);

                if (cert == null) return new Certificate();

                var name = cert.FriendlyName;
                if (string.IsNullOrEmpty(name))
                {
                    name = cert.SubjectName.Name;
                }

                return new Certificate()
                {
                    ThumbPrint = cert.Thumbprint,
                    FriendlyName = name,
                    IssuerName = cert.IssuerName.Name
                };
            }
            catch (Exception ex)
            {
                ex.LogException();
                throw;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                GlobalObjects.ViewModel.SelectedCertificate = new Certificate();
                var publisher = CertPublisher.Text;
                if (!string.IsNullOrEmpty(publisher))
                {
                    var certificate = CreateCertificate(publisher);
                    if (certificate != null)
                    {
                        GlobalObjects.ViewModel.SelectedCertificate = certificate;
                    }
                    Result = System.Windows.Forms.DialogResult.OK;
                    this.Close();
                }
                else
                {
                    throw (new Exception("Publisher name required"));
                }
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
