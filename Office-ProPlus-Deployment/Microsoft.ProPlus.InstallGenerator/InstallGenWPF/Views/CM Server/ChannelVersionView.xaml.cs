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
        private BranchVersion SelectedVersion = BranchVersion.Current;

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

        private void ChannelToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox) sender;
            var branch = checkbox.DataContext as OfficeBranch;
            var index = ChannelList.Items.IndexOf(branch);

            var row = (DataGridRow) ChannelList.ItemContainerGenerator.ContainerFromIndex(index);
            var comobBox = FindVisualChild<ComboBox>(row);
            var selectedVersion = comobBox.SelectedValue;

            var selectedBranch = new SelectedChannel()
            {
                Branch = branch,
                SelectedVersion = (BranchVersion)Enum.Parse(typeof(BranchVersion), selectedVersion.ToString(), true)
            };

        
            if (GlobalObjects.ViewModel.SccmConfiguration.Channels.IndexOf(selectedBranch) == -1)
            {
                GlobalObjects.ViewModel.SccmConfiguration.Channels.Add(selectedBranch);
            }
            else
            {
                foreach (var channel in GlobalObjects.ViewModel.SccmConfiguration.Channels)
                {
                    if (channel.Branch == branch && channel.SelectedVersion != SelectedVersion)
                    {
                        channel.SelectedVersion = SelectedVersion;
                    }
                }
            }
            ToggleNext();
        }

        private void ChannelToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox)sender;
            var branch = checkbox.DataContext as OfficeBranch;

            var unSelectedBranch = new SelectedChannel()
            {
                Branch = branch,
                SelectedVersion = SelectedVersion
            };


            foreach (var channel in GlobalObjects.ViewModel.SccmConfiguration.Channels)
            {
                if (channel.Branch == branch)
                {
                    GlobalObjects.ViewModel.SccmConfiguration.Channels.Remove(channel);
                    break;
                }
            }

            ToggleNext();
        }

        private void ToggleNext()
        {
            var SccmConfiguration = GlobalObjects.ViewModel.SccmConfiguration;

            if (SccmConfiguration.Channels.Count > 0 && SccmConfiguration.Bitnesses.Count > 0)
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

        private void ToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var checkBox = (CheckBox) sender;
            var bitness = checkBox.DataContext as Bitness;

            if (GlobalObjects.ViewModel.SccmConfiguration.Bitnesses.IndexOf(bitness) == -1)
            {
                GlobalObjects.ViewModel.SccmConfiguration.Bitnesses.Add(bitness);
            }

            ToggleNext();
        }

        private void ToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            var checkBox = (CheckBox)sender;
            var bitness = checkBox.DataContext as Bitness;

            foreach (var bitnesses in GlobalObjects.ViewModel.SccmConfiguration.Bitnesses)
            {
                if (bitnesses == bitness)
                {
                    GlobalObjects.ViewModel.SccmConfiguration.Bitnesses.Remove(bitnesses);
                    break;
                }
            }

            ToggleNext();
        }

        private void ChannelVersionPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var grid = (Grid) sender;

            if (grid.Visibility == Visibility.Visible)
            {
                ToggleNext();
            }
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
            var row = combobox.TryFindParent<DataGridRow>(); 
    
            var checkBox = FindVisualChild<CheckBox>(row);
            var branch = checkBox.DataContext as OfficeBranch;

            if (checkBox.IsChecked.Value)
            {
                foreach (var channel in GlobalObjects.ViewModel.SccmConfiguration.Channels)
                {
                    if (channel.Branch.Name == branch.Name)
                    {
                        channel.SelectedVersion =
                            (BranchVersion)Enum.Parse(typeof(BranchVersion), selectedVersion.ToString(), true);
                        break;
                    }
                }
            }
        }
    }
}
