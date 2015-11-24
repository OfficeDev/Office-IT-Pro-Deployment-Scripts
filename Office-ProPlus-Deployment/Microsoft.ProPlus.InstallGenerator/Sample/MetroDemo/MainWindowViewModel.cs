using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
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

    public static class GlobalObjects
    {
        public static MainWindowViewModel ViewModel { get; set; }
    }

    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private readonly IDialogCoordinator _dialogCoordinator;

        private List<Language> _selectedLanguages = null;

        

        public MainWindowViewModel(IDialogCoordinator dialogCoordinator)
        {
            _dialogCoordinator = dialogCoordinator;

            if (DefaultLanguage == null)
            {
                DefaultLanguage = new Language()
                {
                    Id = "en-us",
                    Name = "English",
                    Order = 1
                };
            }

            UseSameLanguagesForAllProducts = true;
            if (_selectedLanguages == null)
            {
                _selectedLanguages = new List<Language>()
                    {
                        DefaultLanguage,
                    };
            }

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

        public Language DefaultLanguage = null;

        public List<Product> MainProducts { get; set; }

        public List<Product> AdditionalProducts { get; set; }

        public List<ExcludeProduct> ExcludeProducts { get; set; }

        public List<Language> Languages { get; set; }



        public bool UseSameLanguagesForAllProducts { get; set; }

        public List<Build> Builds { get; set; } 

        public ConfigXmlParser ConfigXmlParser { get; set; }

        public string DefaultXml =
            "<Configuration></Configuration>";

        public bool ResetXml { get; set; }

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



        public List<Language> SelectedLanguages
        {
            get
            {
                if (_selectedLanguages == null)
                {
                    _selectedLanguages = new List<Language>()
                    {
                        DefaultLanguage,
                    };
                }

                return _selectedLanguages.OrderBy(l => l.Order).ToList();
            }
            set { _selectedLanguages = value; }
        }

        public List<Language> GetLanguages(string productId)
        {
            var languages = _selectedLanguages.Where(
                    l => l.ProductId == productId);

            if (!languages.Any())
            {
                foreach (var language in _selectedLanguages.Where(l => l.ProductId == null).OrderBy(l => l.Order).ToList())
                {
                    _selectedLanguages.Add(new Language()
                    {
                        Id = language.Id,
                        Name = language.Name,
                        Order = 1,
                        ProductId = productId
                    });  
                }

                languages = _selectedLanguages.Where(
                    l => l.ProductId == productId);
            }

            languages = FormatLanguage(languages.Distinct().OrderBy(l => l.Order).ToList());

            return languages.ToList();
        }

        public Language GetLanguage(string productId, string languageId)
        {
            var language = _selectedLanguages.FirstOrDefault(
                    l => l.ProductId == productId && l.Id == languageId);
            return language;
        }

        public void AddLanguages(string productId, List<Language> languages)
        {
            var currentLangs = _selectedLanguages.Where(
                                l => l.ProductId == productId).ToList();

            foreach (var language in currentLangs)
            {
                _selectedLanguages.Remove(language);
            }

            var order = 1;
            foreach (var language in languages.Where(l => l.Order != 1))
            {
                order ++;
                language.Order = order;
            }

            _selectedLanguages.AddRange(languages);
        }

        public void ChangePrimaryLanguage(string productId, Language primaryLanguage)
        {
            var languageItem = _selectedLanguages.FirstOrDefault(
                                l => l.ProductId == productId && l.Id == primaryLanguage.Id);
            var otherProductLanguages = _selectedLanguages.Where(
                l => l.ProductId == productId && l.Id != primaryLanguage.Id);

            if (languageItem != null)
            {
                languageItem.Order = 1; 
            }
            
            var order = 1;
            foreach (var productLanguage in otherProductLanguages)
            {
                order++;
                productLanguage.Order = order;
            }
        }

        public void RemoveLanguage(string productId, string languageId)
        {
            var currentLangs = _selectedLanguages.Where(
                 l => l.ProductId == productId && l.Id == languageId).ToList();

            foreach (var removelanguage in currentLangs)
            {
                _selectedLanguages.Remove(removelanguage);
            }
        }

        private List<Language> FormatLanguage(List<Language> languages)
        {
            if (languages == null) return new List<Language>();
            foreach (var language in languages)
            {
                language.Name = Regex.Replace(language.Name, @"\s\(Primary\)", "", RegexOptions.IgnoreCase);
            }
            if (languages.Any())
            {
                languages.FirstOrDefault().Name += " (Primary)";
            }
            return languages;
        }

    }

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
}