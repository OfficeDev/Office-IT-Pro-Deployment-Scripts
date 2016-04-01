using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ComponentModel;
using System.Globalization;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using System.Windows.Media;
using MahApps.Metro;
using MetroDemo;
using MetroDemo.Models;
using System.Windows.Input;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Extentions;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using OfficeInstallGenerator;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;

namespace MetroDemo
{

    public static class GlobalObjects
    {
        public static MainWindowViewModel ViewModel { get; set; }

        public static string SetBranchFolderPath(string selectBranch, string folderPath)
        {
            var newChannelName = selectBranch;
            if (ViewModel.UseFolderShortNames)
            {
                newChannelName = selectBranch.ConvertChannelToShortName();
            }

            var longFolderPath = "";
            var shortFolderPath = "";

            foreach (var branch in ViewModel.Branches)
            {
                var branchName = branch.Branch.ToString();

                if (folderPath.ToLower().EndsWith(@"\" + branchName.ToLower()))
                {
                    folderPath = Regex.Replace(folderPath, @"\\" + branchName + "$", @"\" + newChannelName, RegexOptions.IgnoreCase);
                    longFolderPath = Regex.Replace(folderPath, @"\\" + branchName + "$", @"\" + selectBranch, RegexOptions.IgnoreCase);
                }
                if (folderPath.ToLower().EndsWith(@"\" + branchName.ConvertChannelToShortName().ToLower()))
                {
                    folderPath = Regex.Replace(folderPath, @"\\" + branchName.ConvertChannelToShortName() + "$", @"\" + newChannelName, RegexOptions.IgnoreCase);
                    shortFolderPath = Regex.Replace(folderPath, @"\\" + branchName.ConvertChannelToShortName() + "$", @"\" + selectBranch.ConvertChannelToShortName(), RegexOptions.IgnoreCase);
                }
            }

            if (!folderPath.ToLower().EndsWith(@"\" + selectBranch.ConvertChannelToShortName().ToLower()) &&
                !folderPath.ToLower().EndsWith(@"\" + selectBranch.ToLower()))
            {
                longFolderPath = folderPath + @"\" + selectBranch;
                shortFolderPath = folderPath + @"\" + selectBranch.ConvertChannelToShortName();

                folderPath += @"\" + newChannelName;
            }

            if (ViewModel.UseFolderShortNames)
            {
                if (Directory.Exists(longFolderPath))
                {
                    Directory.Move(longFolderPath, shortFolderPath);
                }
            }
            else
            {
                if (Directory.Exists(shortFolderPath))
                {
                    Directory.Move(shortFolderPath, longFolderPath);
                }
            }

            return folderPath;
        }

        public async static Task<bool> DirectoryExists(string path)
        {
            var task = Task.Run(() => Directory.Exists(path));
            return await Task.WhenAny(task, Task.Delay(1000)) == task && task.Result;
        }

        public static string DefaultXml = "<Configuration><Updates Enabled=\"TRUE\"></Updates><Display Level=\"Full\" /></Configuration>";
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
                    NewName = "Current",
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
                    Branch = Branch.Business,
                    Name = "Deferred",
                    NewName = "Deferred",
                    Id = "Business",
                    CurrentVersion = "",
                    Versions = new List<Build>()
                    {
                       new Build() { Version = "16.0.6001.1068"},
                       new Build() { Version = "16.0.6001.1061"}
                    }
                },
                new OfficeBranch()
                {
                    Branch = Branch.FirstReleaseCurrent,
                    Name = "First Release Current",
                    NewName = "FirstReleaseCurrent",
                    Id = "FirstReleaseCurrent",
                    CurrentVersion = "",
                    Versions = new List<Build>()
                    {

                    }
                },
                new OfficeBranch()
                {
                    Branch = Branch.FirstReleaseBusiness,
                    Name = "First Release Deferred",
                    NewName = "FirstReleaseDeferred",
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

            Certificates = new List<Certificate>();

            SelectedCertificate = new Certificate();
             

        }

        public bool RunLocalConfigs { get; set; }

        public bool AllowMultipleDownloads { get; set; }

        public bool UseFolderShortNames { get; set; }

        public Certificate SelectedCertificate { get; set; }

        public Language DefaultLanguage = null;

        public List<Product> MainProducts { get; set; }

        public List<Product> AdditionalProducts { get; set; }

        public List<ExcludeProduct> ExcludeProducts { get; set; }

        public List<Language> Languages { get; set; }

        public List<Certificate> Certificates { get; set; } 

        public List<OfficeBranch> Branches { get; set; }

        public bool UseSameLanguagesForAllProducts { get; set; }

        public bool PropertyChangeEventEnabled { get; set; }

        private string _importFile = "";
        public string ImportFile
        {
            get { return _importFile; }
            set
            {
                _importFile = value;
                RaisePropertyChanged("ImportFile");
            }
        }

        private bool _silentInstall = false;
        public bool SilentInstall
        {
            get { return _silentInstall; }
            set
            {
                _silentInstall = value;
                RaisePropertyChanged("SilentInstall");
            }
        }

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

        public bool ResetXml { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void RaisePropertyChanged(string propertyName)
        {
            if (!PropertyChangeEventEnabled) return;
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public string Error { get { return string.Empty; } }

        public bool BlockNavigation { get; set; }

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

        public void SetProductLanguagesForAll(string productId)
        {
            _selectedLanguages = _selectedLanguages.Where(l => (l.ProductId == productId) ||
                            ((l.ProductId != null && productId != null) && l.ProductId.ToLower() == productId.ToLower())).ToList();
            foreach (var language in _selectedLanguages)
            {
                language.ProductId = null;
            }
        }

        public bool IsSigningCert(X509Certificate2 certificate)
        {
            
            foreach (X509Extension ext in certificate.Extensions)
            {
                if (ext.Oid.FriendlyName == "Enhanced Key Usage")
                {
                    var ku = ext as X509EnhancedKeyUsageExtension;
                    foreach (var eku in ku.EnhancedKeyUsages)
                    {
                        if (eku.FriendlyName == "Code Signing")
                        {
                            return true;
                        }
                    }  
                }
            }

            return false; 

        }
        public void SetCertificates()
        {
            try
            {
                Certificates.Clear();
                X509Store localStore = new X509Store(StoreLocation.CurrentUser);
                X509Store machineStore = new X509Store(StoreLocation.LocalMachine);

                localStore.Open(OpenFlags.ReadOnly);
                if (localStore.Certificates.Count > 0)
                {
                    foreach (var certificate in localStore.Certificates)
                    {
                        var cert = new Certificate();

                        Console.WriteLine(certificate.SubjectName.Name);

                        if(IsSigningCert(certificate))
                        {
                            if (String.IsNullOrEmpty(certificate.FriendlyName))
                            {

                                cert.FriendlyName = certificate.SubjectName.Name; 
                            }
                            else
                            {
                                cert.FriendlyName = certificate.FriendlyName;
                            }

                            cert.IssuerName = certificate.IssuerName.Name;
                            cert.ThumbPrint = certificate.Thumbprint;

                            Certificates.Add(cert);
                        }

                    }
                }

                machineStore.Open(OpenFlags.ReadOnly);
                if (machineStore.Certificates.Count > 0)
                {
                    foreach (var certificate in machineStore.Certificates)
                    {
                        Certificate cert = new Certificate();

                        if (IsSigningCert(certificate))
                        {
                            if (String.IsNullOrEmpty(certificate.FriendlyName))
                            {

                                cert.FriendlyName = certificate.SubjectName.Name;
                            }
                            else
                            {
                                cert.FriendlyName = certificate.FriendlyName;
                            }

                            cert.IssuerName = certificate.IssuerName.Name;
                            cert.ThumbPrint = certificate.Thumbprint;

                            Certificates.Add(cert);
                        }

                    }
                }
                localStore.Close();
                machineStore.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
                
        }

        public List<Language> GetLanguages(string productId)
        {
            if (productId != null) productId = productId.ToLower();

            var languages = _selectedLanguages.Where(l => (l.ProductId == productId) ||
                ((l.ProductId != null && productId != null) && l.ProductId.ToLower() == productId.ToLower()));
            

            if (!languages.Any())
            {
                var defaultLanguages = _selectedLanguages.Where(l => l.ProductId == null);
                if (!defaultLanguages.Any())
                {
                    defaultLanguages = _selectedLanguages.Where(l => l.ProductId == "O365ProPlusRetail".ToLower());
                }
                if (!defaultLanguages.Any())
                {
                    defaultLanguages = _selectedLanguages.Where(l => l.ProductId == "O365BusinessRetail".ToLower());
                }

                foreach (var language in defaultLanguages.OrderBy(l => l.Order).ToList())
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


        private IEnumerable<Language> FormatLanguage(List<Language> languages)
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