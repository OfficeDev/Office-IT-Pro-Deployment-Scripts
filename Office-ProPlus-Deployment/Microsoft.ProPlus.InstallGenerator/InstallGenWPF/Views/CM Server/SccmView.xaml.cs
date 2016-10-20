using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;
using MahApps.Metro.Controls;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using MetroDemo.Models;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using OfficeInstallGenerator.Model;
using System.Xml;
using MahApps.Metro.Converters;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config;
using Button = System.Windows.Controls.Button;
using CheckBox = System.Windows.Controls.CheckBox;
using ComboBox = System.Windows.Controls.ComboBox;
using DataGrid = System.Windows.Controls.DataGrid;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using TextBox = System.Windows.Controls.TextBox;
using UserControl = System.Windows.Controls.UserControl;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class CMView : UserControl
    {
    #region declarations
        private LanguagesDialog languagesDialog = null;
        private CancellationTokenSource _tokenSource = new CancellationTokenSource();

        public event TransitionTabEventHandler TransitionTab;
        public event MessageEventHandler ErrorMessage;
        public event ToggleNextEventHandler ToggleNextButton;
        public event MessageEventHandler InfoMessage;


        public List<Language> PackageLanguages { get; set; }
        public List<Bitness> OfficeBitnesses { get; set; }

        public List<SelectedChannel> OfficeChannels { get; set; }
        public string OfficeBitness { get; set; }
        public string ChannelDownloadLocation { get; set; }



        #endregion



        private int _cachedIndex = 0;
   
        public CMView()
        {
            InitializeComponent();

        }

        private  void ToggleNext(object sender, Events.ToggleEventArgs e)
        {
            NextButton.IsEnabled = e.Enabled;
        }

        private void CMView_Loaded(object sender, RoutedEventArgs e)             
        {
            try
            {

                if (cbActions.Items.Count < 1)
                {
                    cbActions.Items.Add("Start an Office ProPlus Deployment");
                    cbActions.Items.Add("Manage an Existing Office ProPlus Deployment");
                    cbActions.Items.Add("Update an Existing Office ProPlus Deployment");
                    cbActions.Items.Add("Remove an Instance of Office ProPlus");
                }

                cbActions.SelectedIndex = 0; 

                Dispatcher.Invoke(() =>
                {
                    StartTab.Visibility = Visibility.Visible;
                    StartPage.Visibility = Visibility.Visible;
                    StartTab.IsSelected = true;
                    PreviousButton.Visibility = Visibility.Collapsed;
                    NextButton.Visibility = Visibility.Collapsed;
                });

                OfficeChannels = new List<SelectedChannel>();
                PackageLanguages = new List<Language>();
                OfficeBitnesses = new List<Bitness>();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LogAnaylytics(string path, string pageName)
        {
            try
            {
                GoogleAnalytics.Log(path, pageName);
            }
            catch { }
        }

        private void TransitionCMTabs(TransitionTabDirection direction)
        {
            var currentIndex = MainTabControl.SelectedIndex;
            var tmpIndex = currentIndex;
            if (direction == TransitionTabDirection.Forward)
            {
                if (MainTabControl.SelectedIndex < MainTabControl.Items.Count - 1)
                {
                    while (tmpIndex < MainTabControl.Items.Count - 1)
                    {
                        tmpIndex++;
                        var item = (TabItem)MainTabControl.Items[tmpIndex];

                        if (item.IsVisible)
                        {
                            MainTabControl.SelectedIndex = tmpIndex;
                            break;
                        }
                    }
                }      
            }
            else
            {
                if (MainTabControl.SelectedIndex > 0)
                {
                    while (tmpIndex != 0)
                    {
                        tmpIndex--;
                        var item = (TabItem)MainTabControl.Items[tmpIndex];

                        if (item.IsVisible)
                        {
                            MainTabControl.SelectedIndex = tmpIndex;
                            break;
                        }
                    }
                }
            }
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

        private void SetTabStatus(bool enabled)
        {
            Dispatcher.Invoke(() =>
            {
                StartTab.IsEnabled = enabled;
            });
        }

        #region "Events"

        private void MainTabControl_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (GlobalObjects.ViewModel.BlockNavigation)
                {
                    MainTabControl.SelectedIndex = _cachedIndex;
                    return;
                }

                              
                switch (MainTabControl.SelectedIndex)
                {
                    case 0:
                        StartPage.Visibility = Visibility.Visible;
                        PreviousButton.Visibility = Visibility.Collapsed;
                        NextButton.Visibility = Visibility.Collapsed;

                        var tabIndex = MainTabControl.Items.Count - 1;

                        while (tabIndex > 0)
                        {
                            var tabItem = (TabItem) MainTabControl.Items[tabIndex];
                            tabItem.Visibility = Visibility.Collapsed;
                            tabIndex--; 
                        } 

                        LogAnaylytics("/CMView", "Start");
                        break;
                    case 1:
                        //DownloadPage.Visibility = Visibility.Visible;
                        LogAnaylytics("/CMView", "Download");
                        break;
                    case 2:
                        LogAnaylytics("/CMView", "Optional");
                        break;
                    case 3:
                        LogAnaylytics("/CMView", "Excluded");
                        break;
                }

                _cachedIndex = MainTabControl.SelectedIndex;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                var currentProgram =
                    GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

                if (GlobalObjects.ViewModel.CmPackage.Scenario == CMScenario.Deploy && MainTabControl.SelectedIndex == 4 && currentProgram.Channels.Count == 0)
                {

                    var ChannelVersionView = new ChannelVersionView();
                    var ProductsLanguagesView = new ProductsLanguagesView();
                    var ProgramOptionsView = new ProgramOptionsView();

                    ChannelVersionView.ToggleNextButton += ToggleNext;
                    ProductsLanguagesView.ToggleNextButton += ToggleNext;
                    ProgramOptionsView.ToggleNextButton += ToggleNext;

                    ProgramOptionsView.ErrorMessage += ErrorMessage;

                    ChannelVersionView.MainTabControl.Items.Remove(ChannelVersionView.ChannelVersionTab);
                    ProductsLanguagesView.MainTabControl.Items.Remove(ProductsLanguagesView.ProductsLanguagesTab);
                    ProgramOptionsView.MainTabControl.Items.Remove(ProgramOptionsView.OtherTab);

                    MainTabControl.Items[2] = ChannelVersionView.ChannelVersionTab;
                    MainTabControl.Items[3] = ProductsLanguagesView.ProductsLanguagesTab;

                    MainTabControl.SelectedIndex = 2;

                    var tab = (TabItem) MainTabControl.Items[4];
                    tab.Visibility = Visibility.Collapsed;


                    MainTabControl.Items[4] = ProgramOptionsView.OtherTab;
                }
                else
                {
                    TransitionCMTabs(TransitionTabDirection.Forward);
                }

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void PreviousButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TransitionCMTabs(TransitionTabDirection.Back);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void cbActions_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (cbActions.SelectedIndex)
            {
                case 0:
                    txtBlock.Text = "Select this option to begin an Office ProPlus deployment.";
                    break;
                case 1:
                    txtBlock.Text = "Select this option to manage an existing Office ProPlus deployment.";
                    break;
                case 2:
                    txtBlock.Text = "Select this option to update an exist Office ProPlus deployment.";
                    break;
                case 3:
                    txtBlock.Text = "Select this option to remove an instance of Office ProPlus.";
                    break;
                default:
                    txtBlock.Text = "";
                    break;
            }
        }

        private void strtButton_Click(object sender, RoutedEventArgs e)
        {

            switch (cbActions.SelectedIndex)
            {
                case 0:
                    GlobalObjects.ViewModel.CmPackage = new CmPackage();
                    DeployOffice();
                    break;
                case 1:
                    GlobalObjects.ViewModel.CmPackage = new CmPackage();
                    break;  
                case 2:
                    GlobalObjects.ViewModel.CmPackage = new CmPackage();

                    break;
                case 3:
                    GlobalObjects.ViewModel.CmPackage = new CmPackage();

                    break;
                case 4:
                    GlobalObjects.ViewModel.CmPackage = new CmPackage();

                    break;
                case 5:
                    GlobalObjects.ViewModel.CmPackage = new CmPackage();

                    break;
                default:
                    LogErrorMessage(new Exception("invalid selection"));
                    break;
            }
        }

        private void DeployOffice()
        {
            var SourceView = new DeploySourceView();
            var ChannelVersionView = new ChannelVersionView();
            var ProductsLanguagesView = new ProductsLanguagesView();
            var ProgramOptionsView = new ProgramOptionsView();
            var PackageOptionsView = new PackageOptionsView();
            var DeploymentStagingView =  new DeploymentStagingView();

            SourceView.ToggleNextButton += ToggleNext;
            ChannelVersionView.ToggleNextButton += ToggleNext;
            ProductsLanguagesView.ToggleNextButton += ToggleNext;
            ProgramOptionsView.ToggleNextButton += ToggleNext;
            PackageOptionsView.ToggleNextButton += ToggleNext;
            DeploymentStagingView.ToggleNextButton += ToggleNext;

            DeploymentStagingView.ErrorMessage += ErrorMessage;
            ProgramOptionsView.ErrorMessage += ErrorMessage;

            SourceView.MainTabControl.Items.Remove(SourceView.SourceTab);
            ChannelVersionView.MainTabControl.Items.Remove(ChannelVersionView.ChannelVersionTab);
            ProductsLanguagesView.MainTabControl.Items.Remove(ProductsLanguagesView.ProductsLanguagesTab);
            ProgramOptionsView.MainTabControl.Items.Remove(ProgramOptionsView.OtherTab);
            PackageOptionsView.MainTabControl.Items.Remove(PackageOptionsView.PackageTab);
            DeploymentStagingView.MainTabControl.Items.Remove(DeploymentStagingView.StagingTab);

            MainTabControl.Items.Add(SourceView.SourceTab);
            MainTabControl.Items.Add(ChannelVersionView.ChannelVersionTab);
            MainTabControl.Items.Add(ProductsLanguagesView.ProductsLanguagesTab);
            MainTabControl.Items.Add(ProgramOptionsView.OtherTab);
            MainTabControl.Items.Add(PackageOptionsView.PackageTab);
            MainTabControl.Items.Add(DeploymentStagingView.StagingTab);

            ChannelVersionView.cbChannelVersion.SelectedIndex = 0;
            
            var tabIndex = 2;
            while (tabIndex < MainTabControl.Items.Count)
            {
                var tempTab = (TabItem) MainTabControl.Items[tabIndex];
                tempTab.IsEnabled = false;

                tabIndex++; 
            }

            var sourceTab = (TabItem) MainTabControl.Items[1];
            sourceTab.IsSelected = true;
            sourceTab.IsEnabled = true;

            GlobalObjects.ViewModel.CmPackage.Scenario = CMScenario.Deploy;

            NextButton.Visibility = Visibility.Visible;
            PreviousButton.Visibility = Visibility.Visible;
        }


        #endregion

        #region "Info"

        private void ProductInfo_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic) sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }


        private InformationDialog informationDialog = null;

        private void LaunchInformationDialog(string sourceName)
        {
            try
            {
                if (informationDialog == null)
                {

                    informationDialog = new InformationDialog
                    {
                        Height = 500,
                        Width = 400
                    };
                    informationDialog.Closed += (o, args) =>
                    {
                        informationDialog = null;
                    };
                    informationDialog.Closing += (o, args) =>
                    {

                    };
                }
                
                informationDialog.Height = 500;
                informationDialog.Width = 400;

                var filePath = AppDomain.CurrentDomain.BaseDirectory + @"HelpFiles\" + sourceName + ".html";
                var helpFile = File.ReadAllText(filePath);

                informationDialog.HelpInfo.NavigateToString(helpFile);
                informationDialog.Launch();

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        #endregion


    }
}

