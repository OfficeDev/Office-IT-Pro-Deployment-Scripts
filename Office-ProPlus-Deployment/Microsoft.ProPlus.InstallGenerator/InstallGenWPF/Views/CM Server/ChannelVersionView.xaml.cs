using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
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
    /// Interaction logic for ChannelVersionView.xaml
    /// </summary>
    public partial class ChannelVersionView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;
        private CmProgram CurrentCmProgram = GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1]; 

        public ChannelVersionView()
        {
            InitializeComponent();
            ToggleNextButton?.Invoke(this, new ToggleEventArgs()
            {
                Enabled = false
            });
        }

        private void ChannelVersionView_OnLoaded(object sender, RoutedEventArgs e)
        {
            
        }

        private void ChannelVersionPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var grid = (Grid)sender;

            if (grid.Visibility == Visibility.Visible)
            {
                CurrentCmProgram =
                    GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];

                if (CurrentCmProgram.Channels.Count == 0 && CurrentCmProgram.Bitnesses.Count == 0)
                {
                    ChannelList.ItemsSource = null;
                    ChannelList.ItemsSource = GlobalObjects.ViewModel.Branches;
                    cbChannelVersion.SelectedIndex = 0;
                    Bit32ToggleButton.IsChecked = false;
                    Bit64ToggleButton.IsChecked = false; 
                }

                ToggleNext();
            }
        }

        private void ChannelToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox) sender;
            var branch = checkbox.DataContext as OfficeBranch;

            var selectedVersion =
                (BranchVersion) Enum.Parse(typeof(BranchVersion), cbChannelVersion.SelectedValue.ToString(), true);

            var selectedBranch = new SelectedChannel()
            {
                Branch = branch,
                SelectedVersion = (BranchVersion)Enum.Parse(typeof(BranchVersion), selectedVersion.ToString(), true)
            };

            if (CurrentCmProgram.Channels != null && !CurrentCmProgram.Channels.Contains(selectedBranch))
                CurrentCmProgram.Channels.Add(selectedBranch); 

            ToggleNext();
        }

        private void ChannelToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox)sender;
            var branch = checkbox.DataContext as OfficeBranch;

            var selectedVersion =
                (BranchVersion)Enum.Parse(typeof(BranchVersion), cbChannelVersion.SelectedValue.ToString(), true);

            var selectedBranch = new SelectedChannel()
            {
                Branch = branch,
                SelectedVersion = (BranchVersion)Enum.Parse(typeof(BranchVersion), selectedVersion.ToString(), true)
            };

            foreach (var channel in CurrentCmProgram.Channels)
            {
                if (channel.Branch.Branch == selectedBranch.Branch.Branch)
                {
                    CurrentCmProgram.Channels.Remove(channel);
                    break;
                }
            }

            ToggleNext();
        }

        private void ToggleNext()
        {
            
            if (CurrentCmProgram.Channels.Count > 0 && CurrentCmProgram.Bitnesses.Count > 0)
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

        private void Bit64ToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {

            CurrentCmProgram.Bitnesses.Add(GlobalObjects.ViewModel.OfficeBitnesses[0]);
            ToggleNext();
        }

        private void Bit32ToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            CurrentCmProgram.Bitnesses.Add(GlobalObjects.ViewModel.OfficeBitnesses[1]);
            ToggleNext();
        }

        private void Bit64ToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            CurrentCmProgram.Bitnesses.Remove(GlobalObjects.ViewModel.OfficeBitnesses[0]);
            ToggleNext();
        }

        private void Bit32ToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            CurrentCmProgram.Bitnesses.Add(GlobalObjects.ViewModel.OfficeBitnesses[1]);
            ToggleNext();
        }


        private childItem FindVisualChild<childItem>(DependencyObject obj)
         where childItem : DependencyObject
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                    if (child != null && child is childItem)
                        return (childItem)child;
                    else
                    {
                        childItem childOfChild = FindVisualChild<childItem>(child);
                        if (childOfChild != null)
                            return childOfChild;
                    }
                }
                return null;
        }

        private void CbChannelVersion_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var combobox = (ComboBox) sender;
            var selectedVersion = combobox.SelectedValue;
            
            CurrentCmProgram.Channels.ForEach(c =>
            {
                c.SelectedVersion = (BranchVersion)Enum.Parse(typeof(BranchVersion), selectedVersion.ToString(), true);
            });
        }

       
    }
}
