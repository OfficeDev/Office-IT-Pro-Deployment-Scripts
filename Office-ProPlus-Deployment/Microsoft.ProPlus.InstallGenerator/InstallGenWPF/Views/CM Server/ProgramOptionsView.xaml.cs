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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using MetroDemo;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.ConfigurationManagement;
using Microsoft.ConfigurationManagement.ManagementProvider;
using Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.Win32;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config
{
    /// <summary>
    /// Interaction logic for DeployOtherView.xaml
    /// </summary>
    public partial class ProgramOptionsView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;
        public event MessageEventHandler ErrorMessage;

        private CmProgram CurrentCmProgram = GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];
        private CMAddCollection AddCollectionDialog = null;
        private CMRemoveCollection RemoveCollectionsDialog = null; 

        
        public ProgramOptionsView()
        {

            InitializeComponent();

            if (GlobalObjects.ViewModel.CmPackage.SiteCode != string.Empty)
            {
                cbSiteCode.SelectedItem = GlobalObjects.ViewModel.CmPackage.SiteCode;
            }
            else
            {
                cbSiteCode.SelectedIndex = 0; 
            }

        }

      

        private void ProgramOptionsView_OnLoaded(object sender, RoutedEventArgs e)
        {
        }

        private void OtherOptionsPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            CurrentCmProgram =
              GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

            if (GlobalObjects.ViewModel.CmPackage.Programs.Count > 1 && CurrentCmProgram.CollectionNames.Count == 0)
            {
                cbSiteCode.IsEnabled = false;
                cbSiteCode.Text = GlobalObjects.ViewModel.CmPackage.SiteCode; 
            }
            else
            {
                cbSiteCode.IsEnabled = true;
            }

            ToggleNext();

            if(cbSiteCode.Items.Count == 0 && GlobalObjects.ViewModel.CmPackage.Programs.Count == 1)
                GetSiteCodes();
        }

        private void ToggleNext()
        {
            if (CurrentCmProgram.CollectionNames.Count > 0)
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
 
        private void ScriptName_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            CurrentCmProgram =
              GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

            var textbox = (TextBox)sender;
            var text = textbox.Text;

            CurrentCmProgram.ScriptName = text; 
        }

        private void ConfigurationXml_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            CurrentCmProgram =
              GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

            var textbox = (TextBox)sender;
            var text = textbox.Text;

            CurrentCmProgram.ConfigurationXml = text;
        }

        private void CustomName_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            CurrentCmProgram =
             GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

            var textbox = (TextBox)sender;
            var text = textbox.Text;

            CurrentCmProgram.CustomName = text;
        }

        private void CbDeploymentPurpose_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                CurrentCmProgram =
                GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

                var comboBox = (ComboBox) sender;
                var index = comboBox.SelectedIndex;
                var value = GlobalObjects.ViewModel.DeploymentPurposes[index];

                CurrentCmProgram.DeploymentPurpose = value;
            }
            catch (Exception ex)
            {
                
            }
           
        }

        private void CbDeploymentType_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                CurrentCmProgram =
                GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

                var comboBox = (ComboBox) sender;
                var index = comboBox.SelectedIndex;
                var value = GlobalObjects.ViewModel.DeploymentTypes[index];

                CurrentCmProgram.DeploymentType = value;
            }
            catch (Exception ex)
            {
                
            }
        }

        private void AddProgram_OnChecked(object sender, RoutedEventArgs e)
        {

        }

        private void AddProgram_OnUnchecked(object sender, RoutedEventArgs e)
        {
            CurrentCmProgram =
                    GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

            if(CurrentCmProgram.Channels.Count == 0)
            GlobalObjects.ViewModel.CmPackage.Programs.RemoveAt(
            GlobalObjects.ViewModel.CmPackage.Programs.Count-1);
        }

        private void BAddCollection_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                CurrentCmProgram =
                GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

                if (AddCollectionDialog == null)
                {
                    var collectionList = GetCollections();

                    AddCollectionDialog = new CMAddCollection
                    {
                        CollectionSource = collectionList
                    };
                    AddCollectionDialog.Closed += (o, args) =>
                    {
                        AddCollectionDialog = null;
                    };
                    AddCollectionDialog.Closing += (o, args) =>
                    {

                        var selectedCollections = AddCollectionDialog.SelectedItems;

                        selectedCollections.ForEach(c =>
                        {
                            if (CurrentCmProgram.CollectionNames.IndexOf(c) == -1)
                            {
                                CurrentCmProgram.CollectionNames.Add(c);
                            }
                        });
                        AddCollectionDialog = null;
                        CollectionList.ItemsSource = null;
                        CollectionList.ItemsSource = CurrentCmProgram.CollectionNames;

                        ToggleNext();
                    };
                }
                AddCollectionDialog.Launch();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }

          
        }

        private void BRemoveCollection_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (RemoveCollectionsDialog == null)
                {
                    var collectionList = CurrentCmProgram.CollectionNames.ToList();

                    RemoveCollectionsDialog = new CMRemoveCollection()
                    {
                        CollectionSource = collectionList
                    };
                    RemoveCollectionsDialog.Closed += (o, args) =>
                    {
                        RemoveCollectionsDialog = null;
                    };
                    RemoveCollectionsDialog.Closing += (o, args) =>
                    {
                        var selectedProducts = RemoveCollectionsDialog.SelectedItems;

                        selectedProducts.ForEach(p =>
                        {
                            if (CurrentCmProgram.CollectionNames.IndexOf(p) > -1)
                                CurrentCmProgram.CollectionNames.Remove(p);
                        });
                        CollectionList.ItemsSource = null;
                        CollectionList.ItemsSource = CurrentCmProgram.CollectionNames;
                        RemoveCollectionsDialog = null;
                        ToggleNext();
                    };
                }
                RemoveCollectionsDialog.Launch();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void CbSiteCode_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CurrentCmProgram =
            GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

            var comboBox = (ComboBox)sender;
            var siteCode = comboBox.SelectedValue.ToString();

            GlobalObjects.ViewModel.CmPackage.SiteCode = siteCode;
            CollectionList.ItemsSource = null;

            ToggleNext();

        }

        #region helpers

        private void GetSiteCodes()
        {
            var siteCodes = new List<string>();
            var sitePath = @"SOFTWARE\Microsoft\SMS\Providers\Sites";
            var siteKey = Registry.LocalMachine.OpenSubKey(sitePath);

            if (siteKey != null)
                siteCodes = siteKey.GetSubKeyNames().ToList();

            if (siteCodes.Count == 0)
                siteCodes.Add("S01");

            cbSiteCode.ItemsSource = siteCodes;
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

        private void LogErrorMessage(Exception ex)
        {
            ex.LogException(false);
            if (ErrorMessage != null)
            {
                ErrorMessage(this, new MessageEventArgs()
                {
                    Title = "Error",
                    Message = ex.Message
                });
            }
        }

        #endregion

    }
}
