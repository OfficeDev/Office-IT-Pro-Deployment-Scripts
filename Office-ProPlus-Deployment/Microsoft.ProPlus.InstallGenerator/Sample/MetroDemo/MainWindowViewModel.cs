using System;
using System.Collections.Generic;
using System.IO;
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
using Micorosft.OfficeProPlus.ConfigurationXml;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using OfficeInstallGenerator;

namespace MetroDemo
{

    public static class GlobalObjects
    {
        public static MainWindowViewModel ViewModel { get; set; }

        public static string SetBranchFolderPath(string selectBranch, string folderPath)
        {
            foreach (var branch in ViewModel.Branches)
            {
                var branchName = branch.Branch.ToString();
                if (folderPath.ToLower().EndsWith(@"\" + branchName.ToLower()))
                {
                    folderPath = Regex.Replace(folderPath, @"\\" + branchName + "$", @"\" + selectBranch, RegexOptions.IgnoreCase);
                }
            }
            if (!folderPath.ToLower().EndsWith(@"\" + selectBranch.ToLower()))
            {
                folderPath += @"\" + selectBranch;
            }
            return folderPath;
        }

        public async static Task<bool> DirectoryExists(string path)
        {
            var task = Task.Run(() => Directory.Exists(path));
            return await Task.WhenAny(task, Task.Delay(1000)) == task && task.Result;
        }
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

            Branches = new List<OfficeBranch>()
            {
                new OfficeBranch()
                {
                    Branch = Branch.Current,
                    Name = "Current",
                    Id = "Current",
                    CurrentVersion = "16.0.6001.1038",
                    Versions = new List<Build>()
                    {
                        new Build() { Version = "16.0.6001.1038"},
                        new Build() { Version = "16.0.4229.1029"},
                        new Build() { Version = "16.0.4229.1024"}
                    }
                },
                new OfficeBranch()
                {
                    Branch = Branch.FirstReleaseCurrent,
                    Name = "First Release Current",
                    Id = "FirstReleaseCurrent",
                    CurrentVersion = "",
                    Versions = new List<Build>()
                    {

                    }
                },
                new OfficeBranch()
                {
                    Branch = Branch.Business,
                    Name = "Business",
                    Id = "Business",
                    CurrentVersion = "",
                    Versions = new List<Build>()
                    {
                    }
                },
                new OfficeBranch()
                {
                    Branch = Branch.FirstReleaseBusiness,
                    Name = "First Release Business",
                    Id = "FirstReleaseBusiness",
                    CurrentVersion = "16.0.6001.1038",
                    Versions = new List<Build>()
                    {
                        new Build() { Version = "16.0.6001.1038"},
                        new Build() { Version = "16.0.4229.1029"},
                        new Build() { Version = "16.0.4229.1024"}
                    }
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

        public List<OfficeBranch> Branches { get; set; }

        public bool UseSameLanguagesForAllProducts { get; set; }

        public List<Build> Builds { get; set; }

        private string _updatePath = "";
        public string UpdatePath
        {
            get { return _updatePath; }
            set
            {
                _updatePath = value;
                RaisePropertyChanged("UpdatePath");
            }
        }

        private string _selectedBranch = "";
        public string SelectedBranch
        {
            get { return _selectedBranch; }
            set
            {
                _selectedBranch = value;
                RaisePropertyChanged("SelectedBranch");
            }
        }

        public ConfigXmlParser ConfigXmlParser { get; set; }

        public string DefaultXml =
            "<Configuration></Configuration>";

        public bool ResetXml { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

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
            if (productId != null) productId = productId.ToLower();

            var languages = _selectedLanguages.Where(
                    l => l.ProductId == productId);

            if (!languages.Any())
            {
                languages = _selectedLanguages.Where(l => l.ProductId == "O365ProPlusRetail".ToLower());
            }

            if (!languages.Any())
            {
                languages = _selectedLanguages.Where(l => l.ProductId == "O365BusinessRetail".ToLower());
            }

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

            if (!languages.Any())
            {
                _selectedLanguages.Add(new Language()
                {
                    Id = DefaultLanguage.Id,
                    Name = DefaultLanguage.Name,
                    Order = 1,
                    ProductId = productId
                });  
            }

            foreach (var language in languages)
            {
                if (language.Name == null)
                {
                    var langLookup = this.Languages.FirstOrDefault(l => l.Id.ToLower() == language.Id.ToLower());
                    if (langLookup != null)
                    {
                        language.Name = langLookup.Name;
                    }
                }
            }

            languages = FormatLanguage(languages.Distinct().OrderBy(l => l.Order).ToList());

            return languages.ToList();
        }

        public Language GetLanguage(string productId, string languageId)
        {
            if (productId != null) productId = productId.ToLower();

            var language = _selectedLanguages.FirstOrDefault(
                    l => l.ProductId == productId && l.Id == languageId);
            return language;
        }

        public void AddLanguages(string productId, List<Language> languages)
        {
            if (productId != null) productId = productId.ToLower();

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

            foreach (var language in languages)
            {
                language.ProductId = language.ProductId != null ? language.ProductId.ToLower() : language.ProductId;
            }

            _selectedLanguages.AddRange(languages);
        }

        public void AddLanguage(string productId, Language language)
        {
            if (productId != null) productId = productId.ToLower();

            var currentLangs = _selectedLanguages.Where(
                                l => l.ProductId == productId).ToList();

            if (currentLangs.Any(l => l.Id.ToLower() == language.Id.ToLower()))
            {
                return;
            }
           
            if (currentLangs.Count > 0)
            {
                language.Order = 2;
            }

            language.ProductId = language.ProductId != null ? language.ProductId.ToLower() : language.ProductId;

            _selectedLanguages.Add(language);
        }

        public void ChangePrimaryLanguage(string productId, Language primaryLanguage)
        {
            if (productId != null) productId = productId.ToLower();

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
            if (productId != null) productId = productId.ToLower();

            var currentLangs = _selectedLanguages.Where(
                 l => l.ProductId == productId && l.Id == languageId).ToList();

            foreach (var removelanguage in currentLangs)
            {
                _selectedLanguages.Remove(removelanguage);
            }
        }

        public void ClearLanguages()
        {
            _selectedLanguages = new List<Language>();
        }


        public void ResetExcludedApps()
        {
            foreach (var excludedApp in ExcludeProducts)
            {
                excludedApp.Included = true;
            }
        }


        private List<Language> FormatLanguage(List<Language> languages)
        {
            if (languages == null) return new List<Language>();
            foreach (var language in languages)
            {
                if (language.Name != null)
                {
                    language.Name = Regex.Replace(language.Name, @"\s\(Primary\)", "", RegexOptions.IgnoreCase);
                }
                else
                {
                    var test = "";
                }
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