using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using MetroDemo.Models;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MetroDemo.Events;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for ChangeChannelView.xaml
    /// </summary>
    public partial class ChangeChannelView : UserControl
    {


        public event TransitionTabEventHandler TransitionTab;
        public event MessageEventHandler InfoMessage;
        public event MessageEventHandler ErrorMessage;
        public ChangeChannelView()
        {
            InitializeComponent();
        }

        private void ChangeChannelView_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void MainTabControl_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void ChangeChannelChannel_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //rollback_checked(sender, new RoutedEventArgs());
        }

        private void rollback_checked(object sender, RoutedEventArgs e)
        {
            if (GlobalObjects.ViewModel.ApplicationMode == Microsoft.OfficeProPlus.InstallGen.Presentation.Enums.ApplicationMode.ChangeChannel)
            {
                GlobalObjects.ViewModel.ChangeChannel = "$scriptPath = \".\"" + Environment.NewLine + Environment.NewLine + "if ($PSScriptRoot) {" + Environment.NewLine + "$scriptPath = $PSScriptRoot" + Environment.NewLine + "} else {" + Environment.NewLine + "$scriptPath = (Get-Item -Path \".\\\").FullName" + Environment.NewLine + "}" + Environment.NewLine + Environment.NewLine + ". $scriptPath\\Change-OfficeChannel.ps1 -Channel " + ChangeChannelChannel.Text.Replace(" ", "") + " -RollBack $" + chkofficeProd.IsChecked.ToString().ToLower() + "";
                string stuff = "";
            }
        }

        private void PreviousButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (TransitionProductTabs(TransitionTabDirection.Back))
                {
                    this.TransitionTab(this, new TransitionTabEventArgs()
                    {
                        Direction = TransitionTabDirection.Back,
                        Index = 0,
                        UseIndex = true
                    });
                }
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
                if (TransitionProductTabs(TransitionTabDirection.Forward))
                {
                    this.TransitionTab(this, new TransitionTabEventArgs()
                    {
                        Direction = TransitionTabDirection.Forward,
                        Index = 6,
                        UseIndex = true
                    });
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }


        private bool TransitionProductTabs(TransitionTabDirection direction)
        {
            var currentIndex = MainTabControl.SelectedIndex;
            var tmpIndex = currentIndex;
            if (direction == TransitionTabDirection.Forward)
            {
                if (MainTabControl.SelectedIndex < MainTabControl.Items.Count - 1)
                {
                    do
                    {
                        tmpIndex++;
                        if (tmpIndex < MainTabControl.Items.Count)
                        {
                            var item = (TabItem)MainTabControl.Items[tmpIndex];
                            if (item == null || item.IsVisible) break;
                        }
                        else
                        {
                            return true;
                        }
                    } while (true);
                    MainTabControl.SelectedIndex = tmpIndex;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                if (MainTabControl.SelectedIndex > 0)
                {
                    do
                    {
                        tmpIndex--;
                        if (tmpIndex > 0)
                        {
                            var item = (TabItem)MainTabControl.Items[tmpIndex];
                            if (item == null || item.IsVisible) break;
                        }
                        else
                        {
                            return true;
                        }
                    } while (true);
                    MainTabControl.SelectedIndex = tmpIndex;
                }
                else
                {
                    return true;
                }
            }

            return false;
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

        private void ChangeChannelChannel_OnSelectionClosed(object sender, EventArgs e)
        {
            rollback_checked(sender, new RoutedEventArgs());
        }
    }

}

