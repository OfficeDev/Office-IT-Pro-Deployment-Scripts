﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MetroDemo.Events;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Enums;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class DisplayView : UserControl
    {
        public MainWindowViewModel ViewModel { get; set; }

        public DisplayView()
        {
            InitializeComponent();
        }

        public void UpdateXml()
        {
            UpdateDisplayXml();

            UpdatePropertiesXml();
        }

        private void UpdateDisplayXml()
        {
            var configXml = ViewModel.ConfigXmlParser.ConfigurationXml;
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

            var xml = ViewModel.ConfigXmlParser.Xml;
            if (xml != null)
            {

            }
        }

        private void UpdatePropertiesXml()
        {
            var configXml = ViewModel.ConfigXmlParser.ConfigurationXml;
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

            var xml = ViewModel.ConfigXmlParser.Xml;
            if (xml != null)
            {

            }
        }

        public event TransitionTabEventHandler TransitionTab;

        private void NextButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateDisplayXml();

                UpdatePropertiesXml();

                this.TransitionTab(this, new TransitionTabEventArgs()
                {
                    Direction = TransitionTabDirection.Forward
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void PreviousButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateDisplayXml();

                UpdatePropertiesXml();

                this.TransitionTab(this, new TransitionTabEventArgs()
                {
                    Direction = TransitionTabDirection.Back
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

    }
}
