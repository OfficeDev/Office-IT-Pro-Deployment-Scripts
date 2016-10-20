using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using MetroDemo.Models;
using Microsoft.ConfigurationManagement.ManagementProvider;
using Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using Microsoft.Win32;
using Application = System.Windows.Application;
using MessageBox = System.Windows.Forms.MessageBox;

namespace MetroDemo.ExampleWindows
{
    public partial class CMAddCollection : IDisposable
    {

        public List<string> SelectedItems { get; set; }

        public List<string> CollectionSource { get; set; } 


        private bool _disposed;
        private bool _hideOnClose = true;
        private CmProgram CurrentCmProgram = GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

        public CMAddCollection()
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

        private void ProductsDialog_OnLoaded(object sender, RoutedEventArgs e)
        {
            CollectionSource = GetCollections();

            CollectionList.ItemsSource = CollectionSource;
        }

        private List<string> GetCollections()
        {
            var collections = new List<string>();
            var siteCode = GlobalObjects.ViewModel.CmPackage.SiteCode;
            var sitePath = @"SOFTWARE\Microsoft\SMS\Providers\Sites";
            var siteKey = Registry.LocalMachine.OpenSubKey(sitePath + $"\\{siteCode}");
            var dbKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\SMS\SQL Server\Site System SQL Account");
            var dbName = dbKey.GetValue("Database Name").ToString();
            var sqlServerName = siteKey.GetValue("SQL Server Name").ToString();
            var connectionString = $"Server= {sqlServerName}; Database= {dbName};Integrated Security=SSPI;";
            var query = $"select [CollectionName] from [{dbName}].[dbo].[Collections_G]";
            SqlDataReader dataReader;

            var connection = new SqlConnection(connectionString);

           
            connection.Open();
            var command = new SqlCommand(query, connection);
            dataReader = command.ExecuteReader();

            while (dataReader.Read())
            {
                collections.Add(dataReader.GetValue(0).ToString());
            }

            dataReader.Close();
            command.Dispose();
            connection.Close();


            return collections;
         

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
                if (CollectionList.SelectedItems.Count > 0)
                {
                    SelectedItems = (List<string>)CollectionList.SelectedItems.Cast<string>().ToList();
                }
                else
                {
                    SelectedItems = new List<string>();
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
            SelectedItems = new List<string>();
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
