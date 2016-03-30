using System;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using MahApps.Metro;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using MetroDemo.Events;
using MetroDemo.ExampleViews;
using MetroDemo.ExampleWindows;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.VisualBasic;

namespace MetroDemo
{
    public partial class MainWindow
    {
        private bool _shutdown;
        private int _cacheIndex = -1;

        public MainWindow()
        {
            GlobalObjects.ViewModel = new MainWindowViewModel(DialogCoordinator.Instance)
            {
                ConfigXmlParser = new OfficeInstallGenerator.ConfigXmlParser("<Configuration></Configuration>"),
                AllowMultipleDownloads = true,
                UseFolderShortNames = true
            };

            DataContext = GlobalObjects.ViewModel;
            

            InitializeComponent();

            ThemeManager.TransitionsEnabled = false;

            MainTabControl.SelectionChanged += MainTabControl_SelectionChanged;

            StartView.TransitionTab += TransitionTab;
            StartView.XmlImported += XmlImported;

            ProductView.TransitionTab += TransitionTab;
            UpdateView.TransitionTab += TransitionTab;
            DisplayView.TransitionTab += TransitionTab;
            GenerateView.TransitionTab += TransitionTab;

            GenerateView.InfoMessage += GenerateViewInfoMessage;
            GenerateView.ErrorMessage += GenerateView_ErrorMessage;

            DisplayView.InfoMessage += GenerateViewInfoMessage;
            DisplayView.ErrorMessage += GenerateView_ErrorMessage;

            ProductView.InfoMessage += GenerateViewInfoMessage;
            ProductView.ErrorMessage += GenerateView_ErrorMessage;

            UpdateView.InfoMessage += GenerateViewInfoMessage;
            UpdateView.ErrorMessage += GenerateView_ErrorMessage;

            StartView.InfoMessage += GenerateViewInfoMessage;
            StartView.ErrorMessage += GenerateView_ErrorMessage;
        }


        public string WindowWidth()
        {
            return ((Panel)Application.Current.MainWindow.Content).ToString(); 
        }

        private void RestartWorkflow(object sender, EventArgs eventArgs)
        {
            for (var i = 1; i < MainTabControl.Items.Count; i++)
            {
                var tabItem = (TabItem)MainTabControl.Items[i];
                //tabItem.IsEnabled = false;
            }

            ProductView.Reset();
            UpdateView.Reset();
            DisplayView.Reset();
        }

        private void XmlImported(object sender, EventArgs eventArgs)
        {
            try
            {
                GlobalObjects.ViewModel.PropertyChangeEventEnabled = false;
                ProductView.LoadXml();
                DisplayView.LoadXml();
                UpdateView.LoadXml();
            }
            catch (Exception ex)
            {
                ex.LogException();
            }
            finally
            {
                GlobalObjects.ViewModel.PropertyChangeEventEnabled = true;
            }
        }

        private void TransitionTab(object sender, Events.TransitionTabEventArgs e)
        {
            try
            {
                var newIndex = Convert.ToInt32(((dynamic)sender).Tag);

                if (e.Direction == TransitionTabDirection.Forward)
                {
                    MainTabControl.SelectedIndex = newIndex + 1;
                }
                else
                {
                    MainTabControl.SelectedIndex = newIndex - 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR:" + ex.Message);
            }
        }

        private async Task ShowMessageDialogAsync(string title, string message)
        {
            await Dispatcher.InvokeAsync(async () =>
            {
                var result = await this.ShowMessageAsync(title, message, 
                    MessageDialogStyle.Affirmative, new MetroDialogSettings()
                {
                    ColorScheme = MetroDialogColorScheme.Theme
                });
            });
        }

        private async Task ShowErrorDialogAsync(string title, string message)
        {
            await Dispatcher.InvokeAsync(async () =>
            {
                var result = await this.ShowMessageAsync(title, message,
                    MessageDialogStyle.Affirmative, new MetroDialogSettings()
                    {
                        ColorScheme = MetroDialogColorScheme.Error
                    });
            });
        }

        #region Events

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StartView.RestartWorkflow == null)
            {
                StartView.RestartWorkflow += RestartWorkflow;
            }

            e.Handled = false;

            if (GlobalObjects.ViewModel.BlockNavigation)
            {
                MainTabControl.SelectedIndex = _cacheIndex;
                return;
            }

            ThemeManager.TransitionsEnabled = MainTabControl.SelectedIndex != 4;
            ThemeManager.TransitionsEnabled = false;

            if (MainTabControl.SelectedIndex > -1)
            {
                ((TabItem)MainWindowTabs.Items[MainTabControl.SelectedIndex]).IsSelected = true;
                ((TabItem)MainWindowTabs.Items[MainTabControl.SelectedIndex]).IsEnabled = true;
            }

            if (_cacheIndex != MainTabControl.SelectedIndex)
            {
                if (!GlobalObjects.ViewModel.ResetXml)
                {
                    ProductView.UpdateXml();
                    DisplayView.UpdateXml();
                    UpdateView.UpdateXml();
                }
                GlobalObjects.ViewModel.ResetXml = false;

                _cacheIndex = MainWindowTabs.SelectedIndex;
            }
        }


        private void UIElement_OnIsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ((MetroTabItem)sender).IsEnabled = true;
        }
        
        private async void GenerateViewInfoMessage(object sender, MessageEventArgs e)
        {
            try
            {
                await ShowMessageDialogAsync(e.Title, e.Message);
            }
            catch (Exception ex)
            {
                ShowErrorDialogAsync("ERROR", ex.Message).ConfigureAwait(false);
            }
        }

        private async void GenerateView_ErrorMessage(object sender, MessageEventArgs e)
        {
            try
            {
                await ShowErrorDialogAsync(e.Title, e.Message);
            }
            catch (Exception ex)
            {
                ShowErrorDialogAsync("ERROR", ex.Message).ConfigureAwait(false);
            }
        }


        private void Nav_OnClick(object sender, RoutedEventArgs e)
        {


            if (OptionsFlyout.Width == 160)
            {
                var xtran = new Duration(TimeSpan.FromMilliseconds(100));
                var widthAnimation = new DoubleAnimation(45, xtran);
                OptionsFlyout.AreAnimationsEnabled = true;
                OptionsFlyout.BeginAnimation(WidthProperty, widthAnimation);
                OptionsFlyout.Width = 45;

                lblStart.Visibility = Visibility.Collapsed;
                lblOptions.Visibility = Visibility.Collapsed;
                lblDownload.Visibility = Visibility.Collapsed;
                lblProducts.Visibility = Visibility.Collapsed;
                lblGenerate.Visibility = Visibility.Collapsed;
                lblUpdate.Visibility = Visibility.Collapsed;
                lblAbout.Visibility = Visibility.Collapsed;

                var margin = ((Button)sender).Margin;
                margin.Left = -1;
                margin.Right = -10;
                ((Button) sender).Margin = margin;

                var mainMargin = ((MetroAnimatedSingleRowTabControl) MainWindowTabs).Margin;

                mainMargin.Left = 45;
                ((MetroAnimatedSingleRowTabControl) MainWindowTabs).Margin = mainMargin;
            }
            else
            {
                var xtran = new Duration(TimeSpan.FromMilliseconds(100));
                var widthAnimation = new DoubleAnimation(160, xtran);
                OptionsFlyout.AreAnimationsEnabled = true;
                OptionsFlyout.BeginAnimation(WidthProperty, widthAnimation);
                OptionsFlyout.Width = 160;

                lblStart.Visibility = Visibility.Visible;
                lblDownload.Visibility = Visibility.Visible;
                lblOptions.Visibility = Visibility.Visible;
                lblProducts.Visibility = Visibility.Visible;
                lblGenerate.Visibility = Visibility.Visible;
                lblUpdate.Visibility = Visibility.Visible;
                lblAbout.Visibility = Visibility.Visible;


                var margin = ((Button)sender).Margin;
                margin.Left = 100;
                ((Button)sender).Margin = margin;

                var mainMargin = ((MetroAnimatedSingleRowTabControl)MainWindowTabs).Margin;
                mainMargin.Left = 150;
                ((MetroAnimatedSingleRowTabControl)MainWindowTabs).Margin = mainMargin;
            }
           

        }

        #endregion

        #region Other
        public static readonly DependencyProperty ToggleFullScreenProperty =
            DependencyProperty.Register("ToggleFullScreen",
                                        typeof(bool),
                                        typeof(MainWindow),
                                        new PropertyMetadata(default(bool), ToggleFullScreenPropertyChangedCallback));

        private static void ToggleFullScreenPropertyChangedCallback(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
        {
            var metroWindow = (MetroWindow)dependencyObject;
            if (e.OldValue != e.NewValue)
            {
                var fullScreen = (bool)e.NewValue;
                if (fullScreen)
                {
                    metroWindow.UseNoneWindowStyle = true;
                    metroWindow.IgnoreTaskbarOnMaximize = true;
                    metroWindow.WindowState = WindowState.Maximized;
                }
                else
                {
                    metroWindow.UseNoneWindowStyle = false;
                    metroWindow.ShowTitleBar = true; // <-- this must be set to true
                    metroWindow.IgnoreTaskbarOnMaximize = false;
                    metroWindow.WindowState = WindowState.Normal;
                }
            }
        }

        public bool ToggleFullScreen
        {
            get { return (bool)GetValue(ToggleFullScreenProperty); }
            set { SetValue(ToggleFullScreenProperty, value); }
        }

        public static readonly DependencyProperty UseAccentForDialogsProperty =
            DependencyProperty.Register("UseAccentForDialogs",
                                        typeof(bool),
                                        typeof(MainWindow),
                                        new PropertyMetadata(default(bool), ToggleUseAccentForDialogsPropertyChangedCallback));

        private static void ToggleUseAccentForDialogsPropertyChangedCallback(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
        {
            var metroWindow = (MetroWindow)dependencyObject;
            if (e.OldValue != e.NewValue)
            {
                var useAccentForDialogs = (bool)e.NewValue;
                metroWindow.MetroDialogOptions.ColorScheme = useAccentForDialogs ? MetroDialogColorScheme.Accented : MetroDialogColorScheme.Theme;
            }
        }

        public bool UseAccentForDialogs
        {
            get { return (bool)GetValue(UseAccentForDialogsProperty); }
            set { SetValue(UseAccentForDialogsProperty, value); }
        }

        private async void CloseCustomDialog(object sender, RoutedEventArgs e)
        {
            var dialog = (BaseMetroDialog)this.Resources["CustomCloseDialogTest"];

            await this.HideMetroDialogAsync(dialog);
        }

        private async void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                e.Cancel = !_shutdown;
                if (_shutdown) return;

                var mySettings = new MetroDialogSettings()
                {
                    AffirmativeButtonText = "Quit",
                    NegativeButtonText = "Cancel",
                    AnimateShow = true,
                    AnimateHide = false
                };

                GenerateView.xmlBrowser.Visibility = Visibility.Hidden;

                var result = await this.ShowMessageAsync("Quit application?",
                    "Sure you want to quit application?",
                    MessageDialogStyle.AffirmativeAndNegative, mySettings);

                _shutdown = result == MessageDialogResult.Affirmative;

                if (_shutdown)
                {
                    Application.Current.Shutdown();
                }
                else
                {
                    GenerateView.xmlBrowser.Visibility = Visibility.Visible;
                }
            }
            catch { }
        }
        #endregion



    }
}
