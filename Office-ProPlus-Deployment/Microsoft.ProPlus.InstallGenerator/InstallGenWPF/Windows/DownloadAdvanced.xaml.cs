using System;
using System.Collections.Generic;
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
    public partial class DownloadAdvanced : IDisposable
    {
        private bool _disposed;
        private bool _hideOnClose = true;
        private bool localOverride = true;

        public DownloadAdvanced()
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

        private void DownloadAdvanced_OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                localOverride = false;
                AllowMultipleDownloads.IsChecked = GlobalObjects.ViewModel.AllowMultipleDownloads;
                UseFolderShortNames.IsChecked = GlobalObjects.ViewModel.UseFolderShortNames;
            }
            catch (Exception ex)
            {
                ex.LogException(true);
            }
            finally
            {
                localOverride = true;
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

        #region Events

        private bool allowCheck = true;
        private void ToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!allowCheck) return;
                var chkBox = (System.Windows.Controls.CheckBox)sender;
                if (GlobalObjects.ViewModel.BlockNavigation && localOverride)
                {
                    allowCheck = false;
                    chkBox.IsChecked = !chkBox.IsChecked;
                }
            }
            catch (Exception ex)
            {
                ex.LogException(true);
            }
            finally
            {
                allowCheck = true;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                GlobalObjects.ViewModel.UseFolderShortNames = UseFolderShortNames.IsChecked.HasValue && UseFolderShortNames.IsChecked.Value;
                GlobalObjects.ViewModel.AllowMultipleDownloads = AllowMultipleDownloads.IsChecked.HasValue && AllowMultipleDownloads.IsChecked.Value;

                this.Close();
            }
            catch (Exception ex)
            {
                ex.LogException();
            }
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
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

        #endregion


    }
}
