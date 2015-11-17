using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MahApps.Metro;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using MetroDemo.Events;
using MetroDemo.ExampleViews;
using MetroDemo.ExampleWindows;

namespace MetroDemo
{
    public partial class MainWindow
    {
        private bool _shutdown;
        private readonly MainWindowViewModel _viewModel;

        public MainWindow()
        {
            _viewModel = new MainWindowViewModel(DialogCoordinator.Instance)
            {
                ConfigXmlParser = new OfficeInstallGenerator.ConfigXmlParser("<Configuration></Configuration>")
            };

            DataContext = _viewModel;

            InitializeComponent();

            ThemeManager.TransitionsEnabled = true;

            MainTabControl.SelectionChanged += MainTabControl_SelectionChanged;

            StartView.TransitionTab += TransitionTab;
            ProductView.TransitionTab += TransitionTab;
            UpdateView.TransitionTab += TransitionTab;
            DisplayView.TransitionTab += TransitionTab;

            StartView.ViewModel = _viewModel;
            ProductView.ViewModel = _viewModel;
            UpdateView.ViewModel = _viewModel;
            DisplayView.ViewModel = _viewModel;
            GenerateView.ViewModel = _viewModel;
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ThemeManager.TransitionsEnabled = MainTabControl.SelectedIndex != 4;

            ProductView.LoadXml();
            ProductView.UpdateXml();
            
            DisplayView.UpdateXml();
            
            UpdateView.UpdateXml();
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
            e.Cancel = !_shutdown;
            if (_shutdown) return;

            var mySettings = new MetroDialogSettings()
            {
                AffirmativeButtonText = "Quit",
                NegativeButtonText = "Cancel",
                AnimateShow = true,
                AnimateHide = false
            };

            var result = await this.ShowMessageAsync("Quit application?",
                "Sure you want to quit application?",
                MessageDialogStyle.AffirmativeAndNegative, mySettings);

            _shutdown = result == MessageDialogResult.Affirmative;

            if (_shutdown)
                Application.Current.Shutdown();
        }

        private void StartView_Loaded(object sender, RoutedEventArgs e)
        {

        }



    }
}
