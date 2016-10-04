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
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class SccmView : UserControl
    {
    #region declarations
        private LanguagesDialog languagesDialog = null;
        private CancellationTokenSource _tokenSource = new CancellationTokenSource();

        public event TransitionTabEventHandler TransitionTab;
        public event MessageEventHandler ErrorMessage;
        public event MessageEventHandler InfoMessage;


        public List<Language> PackageLanguages { get; set; }
        public List<Bitness> OfficeBitnesses { get; set; }

        public List<SelectedChannel> OfficeChannels { get; set; }
        public string OfficeBitness { get; set; }
        public string ChannelDownloadLocation { get; set; }



        #endregion



        private int _cachedIndex = 0;
   
        public SccmView()
        {
            InitializeComponent();
            
        }

        private void SccmView_Loaded(object sender, RoutedEventArgs e)             
        {
            try
            {

                if (cbActions.Items.Count < 1)
                {
                    cbActions.Items.Add("Deploy Office 365 ProPlus");
                    cbActions.Items.Add("Change the channel of an Office 365 client");
                    cbActions.Items.Add("Rollback the version of an Office 365 client");
                    cbActions.Items.Add("Update an Office 365 ProPlus client with ConfigMgr");
                    cbActions.Items.Add("Update an Office 365 ProPlus client using a scheduled task");
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

        private void TransitionSccmTabs(TransitionTabDirection direction)
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

        private void DownloadPage_Loaded()
        {
            PreviousButton.Visibility = Visibility.Visible;
            NextButton.Visibility = Visibility.Visible;
            NextButton.IsEnabled = false;


            StartPage.Visibility = Visibility.Collapsed;
            StartTab.IsSelected = false;

            //DownloadTab.Visibility = Visibility.Visible;
            //DownloadPage.Visibility = Visibility.Visible;
            //DownloadTab.IsSelected = true;

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

                        LogAnaylytics("/SccmView", "Start");
                        break;
                    case 1:
                        //DownloadPage.Visibility = Visibility.Visible;
                        LogAnaylytics("/SccmView", "Download");
                        break;
                    case 2:
                        LogAnaylytics("/SccmView", "Optional");
                        break;
                    case 3:
                        LogAnaylytics("/SccmView", "Excluded");
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
                TransitionSccmTabs(TransitionTabDirection.Forward);
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
                TransitionSccmTabs(TransitionTabDirection.Back);
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
                    txtBlock.Text = "Select this option if you would like to deploy Office 365 ProPlus.";
                    break;
                case 1:
                    txtBlock.Text = "Select this option if would like to change the installed channel of an Office 365 client.";
                    break;
                case 2:
                    txtBlock.Text = "Select this option if you would like to rollback the version of Office 365 installed on a client.";
                    break;
                case 3:
                    txtBlock.Text = "Select this option if you would like to update the version of Office 365 ProPlus installed on a client via ConfigMgr.";
                    break;
                case 4:
                    txtBlock.Text = "Select this option if you would like to update the version of Office 365 ProPlus installed on a client via a scheduled task.";
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
                    DeployOffice();
                    break;
                case 1:
                    GlobalObjects.ViewModel.SccmConfiguration = new SccmConfiguration();
                    break;  
                case 2:
                    GlobalObjects.ViewModel.SccmConfiguration = new SccmConfiguration();

                    break;
                case 3:
                    GlobalObjects.ViewModel.SccmConfiguration = new SccmConfiguration();

                    break;
                case 4:
                    GlobalObjects.ViewModel.SccmConfiguration = new SccmConfiguration();

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
            var DeployOtherView = new DeployOtherView();


            SourceView.MainTabControl.Items.Remove(SourceView.SourceTab);
            ChannelVersionView.MainTabControl.Items.Remove(ChannelVersionView.ChannelVersionTab);
            ProductsLanguagesView.MainTabControl.Items.Remove(ProductsLanguagesView.ProductsLanguagesTab);
            DeployOtherView.MainTabControl.Items.Remove(DeployOtherView.OtherTab);

            MainTabControl.Items.Add(SourceView.SourceTab);
            MainTabControl.Items.Add(ChannelVersionView.ChannelVersionTab);
            MainTabControl.Items.Add(ProductsLanguagesView.ProductsLanguagesTab);
            MainTabControl.Items.Add(DeployOtherView.OtherTab);

            var tabIndex = 2;`
            while (tabIndex < MainTabControl.Items.Count)
            {
                var tempTab = (TabItem) MainTabControl.Items[tabIndex];
                tempTab.IsEnabled = false;

                tabIndex++; 
            }

            var sourceTab = (TabItem) MainTabControl.Items[1];
            sourceTab.IsSelected = true;
            sourceTab.IsEnabled = true;

            GlobalObjects.ViewModel.SccmConfiguration = new SccmConfiguration();
            GlobalObjects.ViewModel.SccmConfiguration.Scenario = SccmScenario.Deploy;

            NextButton.Visibility = Visibility.Visible;
            PreviousButton.Visibility = Visibility.Visible;

            NextButton.IsEnabled = false;
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

