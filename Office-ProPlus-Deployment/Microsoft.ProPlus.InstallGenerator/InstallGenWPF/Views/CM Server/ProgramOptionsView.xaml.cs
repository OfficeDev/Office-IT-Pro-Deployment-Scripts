using System;
using System.Collections.Generic;
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
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config
{
    /// <summary>
    /// Interaction logic for DeployOtherView.xaml
    /// </summary>
    public partial class ProgramOptionsView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;
        private CmProgram CurrentCmProgram = GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1]; 


        public ProgramOptionsView()
        {
            InitializeComponent();
        }

      

        private void ProgramOptionsView_OnLoaded(object sender, RoutedEventArgs e)
        {
           
        }

        private void OtherOptionsPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            CurrentCmProgram =
                    GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];
            ToggleNext();

            if (CurrentCmProgram.CollectionNames.Count == 0 && CurrentCmProgram.ScriptName == string.Empty &&
                CurrentCmProgram.ConfigurationXml == null && CurrentCmProgram.CustomName == string.Empty)
            {
                ScriptName.Text = string.Empty;
                ConfigurationXml.Text = string.Empty;
                CustomName.Text = string.Empty;
                cbDeploymentPurpose.SelectedIndex = 1;
                cbDeploymentType.SelectedIndex = 1;
                AddProgram.IsChecked = false;

                OptionsTab1.IsSelected = true;
            }
        }

        private void ToggleNext()
        {     
            if (Collection.Text.Length > 0)
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
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            CurrentCmProgram.ScriptName = text; 
        }

        private void ConfigurationXml_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = (TextBox)sender;
            var text = textbox.Text;

            CurrentCmProgram.ConfigurationXml = text;
        }

        private void CustomName_OnTextChanged(object sender, TextChangedEventArgs e)
        {

            var textbox = (TextBox)sender;
            var text = textbox.Text;

            CurrentCmProgram.CustomName = text;
        }

        private void Collection_OnTextChanged(object sender, TextChangedEventArgs e)
        {

            var textbox = (TextBox)sender;
            var text = textbox.Text;

            CurrentCmProgram.CollectionNames.Add(text);

            ToggleNext();
        }

        private void CbDeploymentPurpose_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = (ComboBox) sender;
            var index = comboBox.SelectedIndex;
            var value = GlobalObjects.ViewModel.DeploymentPurposes[index];

            CurrentCmProgram.DeploymentPurpose = value; 
        }

        private void CbDeploymentType_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = (ComboBox)sender;
            var index = comboBox.SelectedIndex;
            var value = GlobalObjects.ViewModel.DeploymentTypes[index];

            CurrentCmProgram.DeploymentType = value;
        }


        private void AddProgram_OnChecked(object sender, RoutedEventArgs e)
        {
          var newProgram = new CmProgram();
          GlobalObjects.ViewModel.CmPackage.Programs.Add(newProgram);
        }

        private void AddProgram_OnUnchecked(object sender, RoutedEventArgs e)
        {
            CurrentCmProgram =
                    GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

            if(CurrentCmProgram.Channels.Count == 0)
            GlobalObjects.ViewModel.CmPackage.Programs.RemoveAt(
            GlobalObjects.ViewModel.CmPackage.Programs.Count-1);
        }
    }
}
