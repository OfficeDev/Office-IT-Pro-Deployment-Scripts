using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Navigation;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using MetroDemo.Models;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using Application = System.Windows.Application;
using MessageBox = System.Windows.Forms.MessageBox;

namespace MetroDemo.ExampleWindows
{
    public partial class InformationDialog : IDisposable
    {

        public List<Language> SelectedItems { get; set; }

        public List<Language> LanguageSource { get; set; } 

        private bool _disposed;
        private bool _hideOnClose = true;

        public InformationDialog()
        {
            try
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

                var mainWindow = (MetroWindow) this;
                var windowPlacementSettings = mainWindow.GetWindowPlacementSettings();
                if (windowPlacementSettings.UpgradeSettings)
                {
                    windowPlacementSettings.Upgrade();
                    windowPlacementSettings.UpgradeSettings = false;
                    windowPlacementSettings.Save();
                }
            }
            catch (Exception ex)
            {
                ex.LogException(true);
            }
        }

        private void InformationDialog_OnLoaded(object sender, RoutedEventArgs e)
        {

        }

        private void HelpInfo_OnNavigated(object sender, NavigationEventArgs e)
        {
            try
            {
                Task.Run(() => Dispatcher.Invoke(async () =>
                {
                    await Task.Delay(500);
                    HelpInfo.Visibility = Visibility.Visible;
                }));
            }
            catch (Exception ex)
            {
                ex.LogException(true);
            }
        }

        public void Launch()
        {
            try { 
                Owner = Application.Current.MainWindow;
                // only for this window, because we allow minimizing
                if (WindowState == WindowState.Minimized)
                {
                    WindowState = WindowState.Normal;
                }
                Show();
            }
            catch (Exception ex)
            {
                ex.LogException(true);
            }
        }

        private void CloseWindow_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception ex)
            {
                ex.LogException();
            }
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
