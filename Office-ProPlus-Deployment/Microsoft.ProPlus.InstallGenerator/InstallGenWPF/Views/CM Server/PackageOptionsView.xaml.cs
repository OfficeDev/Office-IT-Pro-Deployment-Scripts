using System;
using System.Collections.Generic;
using System.Data.SqlClient;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using MetroDemo;
using MetroDemo.Events;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.Win32;
using ComboBox = System.Windows.Controls.ComboBox;
using TextBox = System.Windows.Controls.TextBox;
using UserControl = System.Windows.Controls.UserControl;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config
{
    /// <summary>
    /// Interaction logic for DeployOtherView.xaml
    /// </summary>
    public partial class PackageOptionsView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;
        private CmProgram CurrentCmProgram = GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1]; 
    


        public PackageOptionsView()
        {
            InitializeComponent();
        }

      

        private void PackageOptionsView_OnLoaded(object sender, RoutedEventArgs e)
        {
            GetDistributionPointGroups();
            GetDistributionPoints();
        }

        private void PackageOptionsPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            CurrentCmProgram =
                    GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

            GetDistributionPointGroups();
            GetDistributionPoints();

            lblDistributionPoint.IsEnabled = true;
            lblDistributionPointGroupName.IsEnabled = true;

            rdbDistributionPointGroup.IsEnabled = true;
            rdbDistributionPoint.IsEnabled = true;

            if (rdbDistributionPointGroup.IsChecked.Value)
            {
                DistributionPoint.IsEnabled = false;
                DistributionPointGroup.IsEnabled = true;
            }
            else
            {
                DistributionPoint.IsEnabled = true;
                DistributionPointGroup.IsEnabled = true;

            }
               
            
            ToggleNext();
        }

        private void ToggleNext()
        {

            if (GlobalObjects.ViewModel.CmPackage.DeploymentSource == DeploymentSource.DistributionPoint)
            {
                if (GlobalObjects.ViewModel.CmPackage.DistributionPointGroupName.Length > 0 || 
                    GlobalObjects.ViewModel.CmPackage.DistributionPoint.Length > 0)
                {
                    ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                    {
                        Enabled = true
                    });
                }
                else
                {
                    ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                    {
                        Enabled = false
                    });
                }
            }
            else
            {
                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = true
                });    
            }
           
        }

        private void DeploymentExpiryDurationInDays_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            if (text.All(char.IsDigit))
                GlobalObjects.ViewModel.CmPackage.DeploymentExpiryDurationInDays = Convert.ToInt32(text);
            else
            {
                textbox.Text = GlobalObjects.ViewModel.CmPackage.DeploymentExpiryDurationInDays.ToString();
            }
        }

        private void PackageName_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            GlobalObjects.ViewModel.CmPackage.CustomPackageShareName = text;
        }

        private void CMPSModulePath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            if (Directory.Exists(text))
                GlobalObjects.ViewModel.CmPackage.CMPSModulePath = text;
        }

        private void TsMoveFiles_OnChecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.CmPackage.MoveFiles = true;

        }

        private void TsMoveFiles_OnUnchecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.CmPackage.MoveFiles = false;

        }

        private void TsUpdateBits_OnChecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.CmPackage.UpdateOnlyChangedBits = true;
        }

        private void TsUpdateBits_OnUnchecked(object sender, RoutedEventArgs e)
        {
            GlobalObjects.ViewModel.CmPackage.UpdateOnlyChangedBits = false;
        }

        private void BrowseButton_OnClick(object sender, RoutedEventArgs e)
        {
            var filepath = CMPSModulePath.Text;
            var fileBrowser = new FolderBrowserDialog();

            if (filepath != null && Directory.Exists(filepath))
            {
                fileBrowser.SelectedPath = filepath;
            }
            fileBrowser.ShowDialog();
            CMPSModulePath.Text = fileBrowser.SelectedPath;
        }

        private void DistributionPoint_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = (ComboBox) sender;
            var value = comboBox.SelectedValue.ToString();

            GlobalObjects.ViewModel.CmPackage.DistributionPoint = value; 
        }

        private void DistributionPointGroup_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = (ComboBox)sender;
            var value = comboBox.SelectedValue.ToString();

            GlobalObjects.ViewModel.CmPackage.DistributionPointGroupName = value;
        }

        private void RdbDistributionPoint_OnChecked(object sender, RoutedEventArgs e)
        {
            DistributionPoint.IsEnabled = true;
            DistributionPointGroup.IsEnabled = false;

            if (DistributionPoint.Items.Count > 0)
            {
                DistributionPoint.SelectedIndex = 0;
                GlobalObjects.ViewModel.CmPackage.DistributionPointGroupName = "";
                GlobalObjects.ViewModel.CmPackage.DistributionPoint = DistributionPoint.SelectedValue.ToString(); 
            }

        }

        private void RdbDistributionPointGroup_OnChecked(object sender, RoutedEventArgs e)
        {
            DistributionPoint.IsEnabled = false;
            DistributionPointGroup.IsEnabled = true;

            if (DistributionPointGroup.Items.Count > 0)
            {
                DistributionPointGroup.SelectedIndex = 0;
                GlobalObjects.ViewModel.CmPackage.DistributionPoint = "";
                GlobalObjects.ViewModel.CmPackage.DistributionPointGroupName = DistributionPointGroup.SelectedValue.ToString();

            }
        }


        #region helpers
        private void GetDistributionPointGroups()
        {
            var dpGroups = new List<string>();
            var siteCode = GlobalObjects.ViewModel.CmPackage.SiteCode;
            var sitePath = @"SOFTWARE\Microsoft\SMS\Providers\Sites";
            var siteKey = Registry.LocalMachine.OpenSubKey(sitePath + $"\\{siteCode}");
            var dbKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\SMS\SQL Server\Site System SQL Account");
            var dbName = dbKey.GetValue("Database Name").ToString();
            var sqlServerName = siteKey.GetValue("SQL Server Name").ToString();
            var connectionString = $"Server= {sqlServerName}; Database= {dbName};Integrated Security=SSPI;";
            var query = $"select [Name] from [{dbName}].[dbo].[DistributionPointGroup]";
            SqlDataReader dataReader;

            var connection = new SqlConnection(connectionString);

            connection.Open();
            var command = new SqlCommand(query, connection);
            dataReader = command.ExecuteReader();

            while (dataReader.Read())
            {
                dpGroups.Add(dataReader.GetValue(0).ToString());
            }

            dataReader.Close();
            command.Dispose();
            connection.Close();

            if(dpGroups.Count > 0)
                DistributionPointGroup.ItemsSource = dpGroups;
        }


        private void GetDistributionPoints()
        {
            var distribtutionPoints = new List<string>();
            var siteCode = GlobalObjects.ViewModel.CmPackage.SiteCode;
            var sitePath = @"SOFTWARE\Microsoft\SMS\Providers\Sites";
            var siteKey = Registry.LocalMachine.OpenSubKey(sitePath + $"\\{siteCode}");
            var dbKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\SMS\SQL Server\Site System SQL Account");
            var dbName = dbKey.GetValue("Database Name").ToString();
            var sqlServerName = siteKey.GetValue("SQL Server Name").ToString();
            var connectionString = $"Server= {sqlServerName}; Database= {dbName};Integrated Security=SSPI;";
            var query = $"select [ServerName] from [{dbName}].[dbo].[DistributionPoints]";
            SqlDataReader dataReader;

            var connection = new SqlConnection(connectionString);

            connection.Open();
            var command = new SqlCommand(query, connection);
            dataReader = command.ExecuteReader();

            while (dataReader.Read())
            {
                distribtutionPoints.Add(dataReader.GetValue(0).ToString());
            }

            dataReader.Close();
            command.Dispose();
            connection.Close();

            if(distribtutionPoints.Count > 0)
                DistributionPoint.ItemsSource = distribtutionPoints;
        }

        #endregion
    }
}
