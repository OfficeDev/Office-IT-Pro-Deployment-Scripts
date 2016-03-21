using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
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

        private void CreateCertificate(string publisher)
        {
            try
            {


                var makeCertPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "makecert.exe");
                var startDate = DateTime.Now.AddDays(-1).ToString("MM/dd/yyyy").Split(' ')[0];
                var endDate = DateTime.Now.AddYears(2).ToString("MM/dd/yyyy").Split(' ')[0];

                Console.WriteLine(startDate);

                System.IO.File.WriteAllBytes(makeCertPath,
                    Microsoft.OfficeProPlus.InstallGen.Presentation.Properties.Resources.makecert);

                Process createProcess = new Process
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = makeCertPath,
                        Arguments =
                            " -r -pe -n CN=" + publisher +" -b " + startDate + " -e " + endDate +
                            " -eku 1.3.6.1.5.5.7.3.3 -ss My",
                        CreateNoWindow = true,
                        UseShellExecute = false
                    }
                };

                createProcess.Start();

            }
            catch (Exception ex)
            {
                ex.LogException();
            }

        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var publisher = CertPublisher.Text;
                if (!String.IsNullOrEmpty(publisher))
                {
                    CreateCertificate(publisher);
                    this.Close();
 
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
            this.Close();

        }

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
