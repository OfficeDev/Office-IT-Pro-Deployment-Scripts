using System;
using System.Collections.Generic;
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

            var selectedBranch = new SelectedChannel()
            {
                Branch = branch,
                SelectedVersion = SelectedVersion
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

            DisplaySelectedChannels();
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

            DisplaySelectedChannels();
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


            DisplaySelectedBitnesses();
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

            DisplaySelectedBitnesses();
            ToggleNext();
        }

        private void CbDownloadChannel_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cbDownloadChannel.Text = null;
        }

        private void CbDownloadBitness_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cbDownloadBitness.Text = null;
        }

        private void CbChannelVersion_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            GlobalObjects.ViewModel.SccmConfiguration.Version = (BranchVersion) cbChannelVersion.SelectedItem;
        }

        private void DisplaySelectedChannels()
        {
            tbSelectedChannels.Text = "Selected: ";
            GlobalObjects.ViewModel.SccmConfiguration.Channels.ForEach(c =>
            {
                tbSelectedChannels.Text += c.Branch.Name + ", ";
            });
        }

        private void DisplaySelectedBitnesses()
        {
            tbSelectedBitnesses.Text = "Selected: ";
            GlobalObjects.ViewModel.SccmConfiguration.Bitnesses.ForEach(b =>
            {
                tbSelectedBitnesses.Text += b.Name + ", ";
            });
        }


        private void ChannelVersionPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var grid = (Grid) sender;

            if (grid.Visibility == Visibility.Visible)
            {
                ToggleNext();
            }
        }
    }
}
