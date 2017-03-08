using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ComponentModel;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Media;
using MahApps.Metro;
using MetroDemo.Models;
using System.Windows.Input;
using MahApps.Metro.Controls.Dialogs;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Extentions;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using Newtonsoft.Json;
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

        public static async Task<bool> DirectoryExists(string path)
        {
            var task = Task.Run(() => Directory.Exists(path));
            return await Task.WhenAny(task, Task.Delay(1000)) == task && task.Result;
        }

        public static string DefaultXml = "<Configuration><Updates Enabled=\"TRUE\" /><Display Level=\"Full\" /><Property Name=\"PinIconsToTaskbar\" Value=\"TRUE\"/></Configuration>";

        public static string DefaultLanguagePackXml = "<Configuration><Add><Product ID=\"LanguagePack\"><Language ID=\"en-us\" /></Product></Add><Display Level=\"Full\" /></Configuration>";
    }

    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private readonly IDialogCoordinator _dialogCoordinator;
        private List<Language> _selectedLanguages = null;
        private List<Language> _removedLanguages = null;
        private List<RemoteComputer> _remoteComputers = null;
        private string _adminUsername = "";
        private string _adminPassword = "";
        private string _adminDomain = "";

        public MainWindowViewModel(IDialogCoordinator dialogCoordinator)
        {
            _dialogCoordinator = dialogCoordinator;

            ApplicationMode = ApplicationMode.InstallGenerator;

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

            _removedLanguages = new List<Language>();

            Builds = new List<Build>()
            {
                new Build()
                {
                    Version = "16.0.4949.1003"
                }
            };



            Channels = new List<Channel>()
            {
                new Channel()
                {

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
                    CurrentVersion = "16.0.7070.2028",
                    Versions = new List<Build>()
                    {
                        new Build() { Version = "16.0.7070.2028"},
                        new Build() { Version = "16.0.7070.2026"},
                        new Build() { Version = "16.0.7070.2022"},
                        new Build() { Version = "16.0.6965.2063"},
                        new Build() { Version = "16.0.6965.2058"},
                        new Build() { Version = "16.0.6965.2053"},
                        new Build() { Version = "16.0.6868.2067"},
                        new Build() { Version = "16.0.6868.2062"},
                        new Build() { Version = "16.0.6868.2060"},
                        new Build() { Version = "16.0.6769.2040"},
                        new Build() { Version = "16.0.6769.2017"},
                        new Build() { Version = "16.0.6769.2015"},
                        new Build() { Version = "16.0.6741.2021"},
                        new Build() { Version = "16.0.6741.2017"},
                        new Build() { Version = "16.0.6568.2036"},
                        new Build() { Version = "16.0.6568.2034"},
                        new Build() { Version = "16.0.6568.2025"},
                        new Build() { Version = "16.0.6366.2068"},
                        new Build() { Version = "16.0.6366.2062"},
                        new Build() { Version = "16.0.6366.2056"},
                        new Build() { Version = "16.0.6366.2036"},
                        new Build() { Version = "16.0.6001.1043"},
                        new Build() { Version = "16.0.6001.1038"},
                        new Build() { Version = "16.0.6001.1034"},
                        new Build() { Version = "16.0.4229.1029"},
                        new Build() { Version = "16.0.4229.1024"},
                    }
                },
                new OfficeBranch()
                {
                    Branch = Branch.Business,
                    Name = "Deferred",
                    NewName = "Deferred",
                    Id = "Business",
                    CurrentVersion = "16.0.6741.2056",
                    Versions = new List<Build>()
                    {
                        new Build() { Version = "16.0.6741.2056"},
                        new Build() { Version = "16.0.6001.1085"},
                        new Build() { Version = "16.0.6741.2048"},
                        new Build() { Version = "16.0.6001.1082"},
                        new Build() { Version = "16.0.6001.1078"},
                        new Build() { Version = "16.0.6001.1073"},
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
                    CurrentVersion = "16.0.7070.2030",
                    Versions = new List<Build>()
                    {
                        new Build() { Version = "16.0.7070.2030"},
                        new Build() { Version = "16.0.7070.2026"},
                        new Build() { Version = "16.0.7070.2022"},
                        new Build() { Version = "16.0.7070.2028"},
                        new Build() { Version = "16.0.7070.2020"},
                        new Build() { Version = "16.0.7070.2019"},
                        new Build() { Version = "16.0.6769.2015"},
                        new Build() { Version = "16.0.6769.2011"},
                        new Build() { Version = "16.0.6741.2017"},
                        new Build() { Version = "16.0.6741.2015"},
                        new Build() { Version = "16.0.6741.2014"},
                        new Build() { Version = "16.0.6568.2036"},
                        new Build() { Version = "16.0.6568.2025"},
                        new Build() { Version = "16.0.6568.2016"},
                        new Build() { Version = "16.0.6366.2062"},
                        new Build() { Version = "16.0.6366.2056"},
                        new Build() { Version = "16.0.6366.2047"},
                        new Build() { Version = "16.0.6366.2036"},
                        new Build() { Version = "16.0.6366.2025"},
                        new Build() { Version = "16.0.6228.1010"},
                        new Build() { Version = "16.0.6228.1007"},
                        new Build() { Version = "16.0.6228.1004"}
                    }
                },
                new OfficeBranch()
                {
                    Branch = Branch.FirstReleaseBusiness,
                    Name = "First Release Deferred",
                    NewName = "FirstReleaseDeferred",
                    Id = "FirstReleaseBusiness",
                    CurrentVersion = "16.0.6965.2069",
                    Versions = new List<Build>()
                    {
                        new Build() { Version = "16.0.6965.2069"},
                        new Build() { Version = "16.0.6965.2066"},
                        new Build() { Version = "16.0.6965.2063"},
                        new Build() { Version = "16.0.6965.2058"},
                        new Build() { Version = "16.0.6741.2047"},
                        new Build() { Version = "16.0.6741.2042"},
                        new Build() { Version = "16.0.6741.2037"},
                        new Build() { Version = "16.0.6741.2033"},
                        new Build() { Version = "16.0.6741.2026"},
                        new Build() { Version = "16.0.6741.2025"},
                        new Build() { Version = "16.0.6741.2021"},
                        new Build() { Version = "16.0.6741.2017"},
                        new Build() { Version = "16.0.6741.2015"},
                        new Build() { Version = "16.0.6741.2014"},
                        new Build() { Version = "16.0.6001.1061"},
                        new Build() { Version = "16.0.6001.1054"},
                        new Build() { Version = "16.0.6001.1043"},
                        new Build() { Version = "16.0.6001.1038"},
                        new Build() { Version = "16.0.6001.1034"},
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
                    Id = "O365ProPlusRetail",
                    ShortName = "Office 365 ProPlus"
                },
                new Product()
                {
                    DisplayName = "Office 365 for Business",
                    Id = "O365BusinessRetail",
                    ShortName = "Office 365 for Business"
                }
            };

            LanguagePackProducts = new List<Product>()
            {
                new Product()
                {
                    Id = "LanguagePack",
                    DisplayName = "Language Pack",
                    ShortName = "Language Pack"
                }
            };

            VisioProducts = new List<Product>()
            {
                new Product()
                {
                    DisplayName = "Visio for Office 365",
                    Id = "VisioProRetail",
                    ShortName = "Visio for Office 365"
                },
                new Product()
                {
                    DisplayName = "Visio for Office 365 Professional (Volume License)",
                    Id = "VisioProXVolume",
                    ShortName = "Visio for Office 365"
                },
                 new Product()
                {
                    DisplayName = "Visio for Office 365 Standard (Volume License)",
                    Id = "VisioStdXVolume",
                    ShortName = "Visio for Office 365"
                },
            };

            ProjectProducts = new List<Product>()
            {
                new Product()
                {
                    DisplayName = "Project for Office 365",
                    Id = "ProjectProRetail",
                    ShortName = "Project for Office 365"
                },
                new Product()
                {
                    DisplayName = "Project for Office 365 Professional(Volume License)",
                    Id = "ProjectProXVolume",
                    ShortName = "Project for Office 365"
                },
                new Product()
                {
                    DisplayName = "Project for Office 365 Standard (Volume License)",
                    Id = "ProjectStdXVolume",
                    ShortName = "Project for Office 365"
                }
            };

            SkypeProducts = new List<Product>()
            {
                new Product()
                {
                    DisplayName = "Skype For Business 2016",
                    Id = "SkypeforBusinessRetail",
                    ShortName = "Skype For Business 2016"
                },
                new Product()
                {
                    DisplayName = "Skype For Business Basic 2016",
                    Id = "SkypeforBusinessEntryRetail",
                    ShortName = "Skype For Business Basic 2016"
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
                },
                new ExcludeProduct()
                {
                    DisplayName = "OneDrive"
                }
            };

            Languages = new List<Language>()
            {
                new Language { Id="en-us", Name="English" },
                new Language { Id="MatchOS", Name="MatchOS" },
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
                new Language { Id="kk-kz", Name="Kazakh" },
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
                new Language { Id="uk-ua", Name="Ukrainian" },
                new Language { Id="vi-vn", Name="Vietnamese" }
            };

            Certificates = new List<Certificate>();

            SelectedCertificate = new Certificate();


        }

        public bool LocalConfig { get; set; }

        public string BranchesToJson
        {
            get
            {
                var json = JsonConvert.SerializeObject(Branches.ToArray());
                return json;
            }
        }

        public List<OfficeBranch> JsonToBranches(string json)
        {
            var updatedBranches = JsonConvert.DeserializeObject<List<UpdatedOfficeBranch>>(json);
            var branches = new List<OfficeBranch>();
            foreach (var updatedBranch in updatedBranches)
            {
                if (updatedBranch.Name.ToLower() != "extendeddeferred") {
                    var branch = new OfficeBranch();
                    if (updatedBranch.Name.ToLower() == "current")
                    {
                        branch.Branch = Branch.Current;
                        branch.Name = updatedBranch.Name;
                        branch.NewName = updatedBranch.Name;
                        branch.Id = updatedBranch.Name;
                    }
                    if (updatedBranch.Name.ToLower() == "deferred")
                    {
                        branch.Branch = Branch.Business;
                        branch.Name = updatedBranch.Name;
                        branch.NewName = updatedBranch.Name;
                        branch.Id = "Business";
                    }
                    if (updatedBranch.Name.ToLower() == "firstreleasedeferred")
                    {
                        branch.Branch = Branch.FirstReleaseBusiness;
                        branch.Name = "First Release Deferred";
                        branch.NewName = updatedBranch.Name;
                        branch.Id = "FirstReleaseBusiness";
                    }
                    if (updatedBranch.Name.ToLower() == "insidersslow")
                    {
                        branch.Branch = Branch.FirstReleaseCurrent;
                        branch.Name = "First Release Current";
                        branch.NewName = "FirstReleaseCurrent";
                        branch.Id = "FirstReleaseCurrent";
                    }
                    branch.Updated = false;
                    foreach (var update in updatedBranch.Updates)
                    {
                        if (update.Latest == true) { branch.CurrentVersion = update.LegacyVersion; }
                        var build = new Build();
                        build.NewBuild = update.Build;
                        build.NewVersion = update.Version;
                        build.Version = update.LegacyVersion;
                        if (branch.Versions == null)
                        {
                            branch.Versions = new List<Build>();
                        }
                        branch.Versions.Add(build);
                    }
                    branches.Add(branch);
                }
            }
            return branches;
        }

        public ApplicationMode ApplicationMode { get; set; }

        public bool AllowMultipleDownloads { get; set; }

        public bool UseFolderShortNames { get; set; }


        public void SetCredentials(string uName, string password, string domain)
        {
            _adminUsername = uName;
            _adminPassword = password;
            _adminDomain = domain;
        }

        public string GetUsername()
        {
            return _adminUsername;
        }

        public string GetPassword()
        {
            return _adminPassword;
        }

        public string GetDomain()
        {
            return _adminDomain;
        }
        public List<RemoteComputer> RemoteConnectionInfo(string connectionInfo = null)
        {
            if (_remoteComputers == null)
            {
                _remoteComputers = new List<RemoteComputer>();
            }

            if (!string.IsNullOrEmpty(connectionInfo))
            {
                _remoteComputers = new List<RemoteComputer>();

                var lineSplit = Microsoft.VisualBasic.Strings.Split(connectionInfo, Environment.NewLine);
                foreach (var line in lineSplit)
                {
                    var computerName = "";
                    string userDomain = null;
                    string userName = null;
                    string userPassword = null;

                    var remoteComputer = new RemoteComputer();

                    var splitChar = ' ';
                    if (line.Contains((char)9))
                    {
                        splitChar = (char)9;
                    }

                    var lineInfo = line.Split(splitChar);
                    computerName = lineInfo[0];
                    if (lineInfo.Length > 1)
                    {
                        userName = lineInfo[1];
                        if (userName.Contains(@"\"))
                        {
                            userDomain = userName.Split('\\')[0];
                            userName = userName.Split('\\')[1];
                        }
                    }
                    if (lineInfo.Length > 2)
                    {
                        userPassword = lineInfo[1];
                    }

                    remoteComputer.UserName = userName;
                    remoteComputer.Domain = userDomain;
                    remoteComputer.Name = computerName;
                    remoteComputer.Password = userPassword;

                    _remoteComputers.Add(remoteComputer);
                }
            }

            return _remoteComputers;
        }

        //public string RemoteLoggingPath { get; set; }




        public Certificate SelectedCertificate { get; set; }

        public Language DefaultLanguage = null;

        public List<RemoteMachine> RemoteMachines { get; set; }

        public List<Channel> Channels { get; set; }

        public List<Product> MainProducts { get; set; }

        public List<Product> LanguagePackProducts { get; set; }

        public List<Product> VisioProducts { get; set; }

        public List<Product> ProjectProducts { get; set; }

        public List<Product> SkypeProducts { get; set; }

        public List<ExcludeProduct> ExcludeProducts { get; set; }

        public List<Language> Languages { get; set; }

        public List<Certificate> Certificates { get; set; }

        public List<OfficeBranch> Branches { get; set; }

        public bool UseSameLanguagesForAllProducts { get; set; }

        public bool PropertyChangeEventEnabled { get; set; }

        public string DownloadFolderPath { get; set; }

        private string _remoteLoggingPath = "";
        public string RemoteLoggingPath
        {
            get { return _remoteLoggingPath; }
            set
            {
                _remoteLoggingPath = value;
                RaisePropertyChanged("RemoteLoggingPath");
            }
        }

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

        public string newVersion { get; set; }

        public string newChannel { get; set; }

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
                var localStore = new X509Store(StoreLocation.CurrentUser);
                var machineStore = new X509Store(StoreLocation.LocalMachine);

                localStore.Open(OpenFlags.ReadOnly);
                if (localStore.Certificates.Count > 0)
                {
                    foreach (var certificate in localStore.Certificates)
                    {
                        var cert = new Certificate();
                        if (IsSigningCert(certificate))
                        {
                            if (string.IsNullOrEmpty(certificate.FriendlyName))
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
                        var cert = new Certificate();

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

        public List<Language> GetRemovedLanguages()
        {
            return _removedLanguages.ToList();
        }

        public List<Language> GetRemovedLanguages(string productId)
        {
            if (productId != null) productId = productId.ToLower();

            var languages = _removedLanguages.Where(l => (l.ProductId == productId) ||
                ((l.ProductId != null && productId != null) && l.ProductId.ToLower() == productId.ToLower()));

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
                order++;
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

        public void AddRemovedLanguage(string productId, Language language)
        {
            if (productId != null) productId = productId.ToLower();

            var currentLangs = _removedLanguages.Where(l => l.ProductId == productId).ToList();

            if (currentLangs.Any(l => l.Id.ToLower() == language.Id.ToLower()))
            {
                return;
            }

            language.ProductId = language.ProductId != null ? language.ProductId.ToLower() : language.ProductId;

            _removedLanguages.Add(language);
        }

        public void RemoveAddRemovedLanguage(string productId, Language language)
        {
            if (productId != null) productId = productId.ToLower();

            var currentLang = _removedLanguages.FirstOrDefault(l => l.ProductId == productId);
            if (currentLang != null)
            {
                _removedLanguages.Remove(currentLang);
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