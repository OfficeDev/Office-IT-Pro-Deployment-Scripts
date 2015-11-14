using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MahApps.Metro;
using MetroDemo;
using MetroDemo.Models;
using System.Windows.Input;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using OfficeInstallGenerator;

namespace MetroDemo
{
    public class AccentColorMenuData
    {
        public string Name { get; set; }
        public Brush BorderColorBrush { get; set; }
        public Brush ColorBrush { get; set; }

        private ICommand changeAccentCommand;

        public ICommand ChangeAccentCommand
        {
            get { return null; }
        }

        protected virtual void DoChangeTheme(object sender)
        {
            var theme = ThemeManager.DetectAppStyle(Application.Current);
            var accent = ThemeManager.GetAccent(this.Name);
            ThemeManager.ChangeAppStyle(Application.Current, accent, theme.Item1);
        }
    }

    public class AppThemeMenuData : AccentColorMenuData
    {
        protected override void DoChangeTheme(object sender)
        {
            var theme = ThemeManager.DetectAppStyle(Application.Current);
            var appTheme = ThemeManager.GetAppTheme(this.Name);
            ThemeManager.ChangeAppStyle(Application.Current, theme.Item2, appTheme);
        }
    }

    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private readonly IDialogCoordinator _dialogCoordinator;

        public MainWindowViewModel(IDialogCoordinator dialogCoordinator)
        {
            _dialogCoordinator = dialogCoordinator;

            Builds = new List<Build>()
            {
                new Build()
                {
                    Version = "16.0.4949.1003"
                }
            };

            MainProducts = new List<Product>()
            {
                new Product()
                {
                    DisplayName = "Office 365 ProPlus",
                    Id = "O365ProPlusRetail"
                },
                new Product()
                {
                    DisplayName = "Office 365 for Business",
                    Id = "O365BusinessRetail"
                }
            };

            AdditionalProducts = new List<Product>()
            {
                new Product()
                {
                    DisplayName = "Visio for Office 365",
                    Id = "VisioProRetail"
                },
                new Product()
                {
                    DisplayName = "Project for Office 365",
                    Id = "ProjectProRetail"
                }
            };

            ExcludeProducts = new List<ExcludeProduct>()
            {
                new ExcludeProduct()
                {
                    DisplayName = "Access"
                },
                new ExcludeProduct()
                {
                    DisplayName = "Excel"
                },
                new ExcludeProduct()
                {
                    DisplayName = "Groove"
                },
                new ExcludeProduct()
                {
                    DisplayName = "Lync"
                },
                new ExcludeProduct()
                {
                    DisplayName = "OneNote"
                },
                new ExcludeProduct()
                {
                    DisplayName = "Outlook"
                },
                new ExcludeProduct()
                {
                    DisplayName = "PowerPoint"
                },
                new ExcludeProduct()
                {
                    DisplayName = "Project"
                },
                new ExcludeProduct()
                {
                    DisplayName = "Publisher"
                },
                new ExcludeProduct()
                {
                    DisplayName = "Visio"
                },
                new ExcludeProduct()
                {
                    DisplayName = "Word"
                }
            };

            Languages = new List<Language>()
            {
                new Language { Id="en-us", Name="English" },
                new Language { Id="ar-sa", Name="Arabic" },
                new Language { Id="bg-bg", Name="Bulgarian" },
                new Language { Id="zh-cn", Name="Chinese - Simplified" },
                new Language { Id="zh-tw", Name="Chinese" },
                new Language { Id="hr-hr", Name="Croatian" },
                new Language { Id="cs-cz", Name="Czech" },
                new Language { Id="da-dk", Name="Danish" },
                new Language { Id="nl-nl", Name="Dutch" },
                new Language { Id="et-ee", Name="Estonian" },
                new Language { Id="fi-fi", Name="Finnish" },
                new Language { Id="fr-fr", Name="French" },
                new Language { Id="de-de", Name="German" },
                new Language { Id="el-gr", Name="Greek" },
                new Language { Id="he-il", Name="Hebrew" },
                new Language { Id="hi-in", Name="Hindi" },
                new Language { Id="hu-hu", Name="Hungarian" },
                new Language { Id="id-id", Name="Indonesian" },
                new Language { Id="it-it", Name="Italian" },
                new Language { Id="ja-jp", Name="Japanese" },
                new Language { Id="kk-kh", Name="Kazakh" },
                new Language { Id="ko-kr", Name="Korean" },
                new Language { Id="lv-lv", Name="Latvian" },
                new Language { Id="lt-lt", Name="Lithuanian" },
                new Language { Id="ms-my", Name="Malay" },
                new Language { Id="nb-no", Name="Norwegian - Bokml" },
                new Language { Id="pl-pl", Name="Polish" },
                new Language { Id="pt-br", Name="Portuguese - Brazil" },
                new Language { Id="pt-pt", Name="Portuguese - Portugal" },
                new Language { Id="ro-ro", Name="Romanian" },
                new Language { Id="ru-ru", Name="Russian" },
                new Language { Id="sr-latn-rs", Name="Serbian - Latin" },
                new Language { Id="sk-sk", Name="Slovak" },
                new Language { Id="sl-si", Name="Slovenian" },
                new Language { Id="es-es", Name="Spanish" },
                new Language { Id="sv-se", Name="Swedish" },
                new Language { Id="th-th", Name="Thai" },
                new Language { Id="tr-tr", Name="Turkish" },
                new Language { Id="uk-ua", Name="Ukrainian" }
            };

        }

        public List<Product> MainProducts { get; set; }

        public List<Product> AdditionalProducts { get; set; }

        public List<ExcludeProduct> ExcludeProducts { get; set; }

        public List<Language> Languages { get; set; }

        public List<Build> Builds { get; set; } 

        public ConfigXmlParser ConfigXmlParser { get; set; }


        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises the PropertyChanged event if needed.
        /// </summary>
        /// <param name="propertyName">The name of the property that changed.</param>
        protected virtual void RaisePropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public string Error { get { return string.Empty; } }


    }
}