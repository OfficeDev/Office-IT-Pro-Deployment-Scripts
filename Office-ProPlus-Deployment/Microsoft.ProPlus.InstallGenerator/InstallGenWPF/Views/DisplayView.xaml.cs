using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Enums;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class DisplayView : UserControl
    {

        public event MessageEventHandler InfoMessage;
        public event MessageEventHandler ErrorMessage;
        public event TransitionTabEventHandler TransitionTab;

        public DisplayView()
        {
            InitializeComponent();
        }

        private void DisplayView_OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (GlobalObjects.ViewModel == null) return;
                GlobalObjects.ViewModel.PropertyChanged += ViewModel_PropertyChanged;

                LogAnaylytics("/DisplayView", "Load");
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        public void LoadXml()
        {
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Display != null)
            {

                DisplayLevel.IsChecked = (configXml.Display.Level.HasValue &&
                    configXml.Display.Level.Value == Micorosft.OfficeProPlus.ConfigurationXml.Enums.DisplayLevel.Full);

                AcceptEula.IsChecked = configXml.Display.AcceptEULA.HasValue && configXml.Display.AcceptEULA.Value;
            }

            if (configXml.Properties != null)
            {
                AutoActivate.IsChecked = configXml.Properties.AutoActivate.HasValue && configXml.Properties.AutoActivate.Value == YesNo.Yes;
                ForceAppShutdown.IsChecked = configXml.Properties.ForceAppShutdown.HasValue && configXml.Properties.ForceAppShutdown.Value;
                SharedComputerLicensing.IsChecked = configXml.Properties.SharedComputerLicensing.HasValue && configXml.Properties.SharedComputerLicensing.Value;
                PinIconsToTaskbar.IsChecked = configXml.Properties.PinIconsToTaskbar.HasValue && configXml.Properties.PinIconsToTaskbar.Value;
            }
        }

        public void Reset()
        {
            DisplayLevel.IsChecked = true;
            AcceptEula.IsChecked = false;
            AutoActivate.IsChecked = false;
            ForceAppShutdown.IsChecked = false;
            SharedComputerLicensing.IsChecked = false;
            PinIconsToTaskbar.IsChecked = true;
        }

        public void UpdateXml()
        {
            UpdateDisplayXml();

            UpdatePropertiesXml();

            UpdateConfigManagerXml();
        }

        private void UpdateDisplayXml()
        {
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Display == null)
            {
                configXml.Display = new ODTDisplay();
            }

            configXml.Display.Level = null;
            if (DisplayLevel.IsChecked.HasValue)
            {
                if (DisplayLevel.IsChecked.Value)
                {
                    configXml.Display.Level = (DisplayLevel) Enum.Parse(typeof (DisplayLevel), DisplayLevel.OnLabel);
                }
                else
                {
                    configXml.Display.Level = (DisplayLevel)Enum.Parse(typeof(DisplayLevel), DisplayLevel.OffLabel);
                }
            }

            configXml.Display.AcceptEULA = null;
            if (AcceptEula.IsChecked.HasValue)
            {
                configXml.Display.AcceptEULA = AcceptEula.IsChecked.Value;
            }
        }

        private void UpdatePropertiesXml()
        {
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Properties == null)
            {
                configXml.Properties = new ODTProperties();
            }

            configXml.Properties.AutoActivate = null;
            if (AutoActivate.IsChecked.HasValue)
            {
                if (AutoActivate.IsChecked.Value)
                {
                    configXml.Properties.AutoActivate = (YesNo)Enum.Parse(typeof(YesNo), AutoActivate.OnLabel);
                }
                else
                {
                    configXml.Properties.AutoActivate = (YesNo)Enum.Parse(typeof(YesNo), AutoActivate.OffLabel);
                }
            }

            configXml.Properties.ForceAppShutdown = null;
            if (ForceAppShutdown.IsChecked.HasValue)
            {
                configXml.Properties.ForceAppShutdown = ForceAppShutdown.IsChecked.Value;
            }

            configXml.Properties.SharedComputerLicensing = null;
            if (SharedComputerLicensing.IsChecked.HasValue)
            {
                configXml.Properties.SharedComputerLicensing = SharedComputerLicensing.IsChecked.Value;
            }

            configXml.Properties.PinIconsToTaskbar = null;
            if (PinIconsToTaskbar.IsChecked.HasValue)
            {
                configXml.Properties.PinIconsToTaskbar = PinIconsToTaskbar.IsChecked.Value;
            }

        }

        private void UpdateConfigManagerXml()
        {
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Add != null)
            {
                if (EnableSCCMSupport.IsChecked.HasValue)
                {
                    configXml.Add.OfficeMgmtCOM = EnableSCCMSupport.IsChecked.Value;
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

        private void LogAnaylytics(string path, string pageName)
        {
            try
            {
                GoogleAnalytics.Log(path, pageName);
            }
            catch { }
        }

        #region Events

        private void ViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            try
            {
                if (e.PropertyName.ToUpper() == "SilentInstall".ToUpper())
                {
                    var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;

                    if (configXml.Display.Level.HasValue &&
                        configXml.Display.Level.Value ==
                        Micorosft.OfficeProPlus.ConfigurationXml.Enums.DisplayLevel.None)
                    {
                        DisplayLevel.IsChecked = false;
                    }
                    else
                    {
                        DisplayLevel.IsChecked = true;
                    }

                    if (configXml.Display.AcceptEULA.HasValue && configXml.Display.AcceptEULA.Value)
                    {
                        AcceptEula.IsChecked = true;
                    }
                    else
                    {
                        AcceptEula.IsChecked = false;
                    }
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void NextButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateDisplayXml();

                UpdatePropertiesXml();

                UpdateConfigManagerXml();

                this.TransitionTab(this, new TransitionTabEventArgs()
                {
                    Direction = TransitionTabDirection.Forward
                });
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void PreviousButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateDisplayXml();

                UpdatePropertiesXml();

                UpdateConfigManagerXml();

                this.TransitionTab(this, new TransitionTabEventArgs()
                {
                    Direction = TransitionTabDirection.Back
                });
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        #endregion

        #region Info

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

        private void ProductInfo_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic)sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        #endregion


    }
}
