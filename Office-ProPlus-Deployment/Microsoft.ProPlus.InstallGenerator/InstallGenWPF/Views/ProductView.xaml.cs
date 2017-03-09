using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using MahApps.Metro.Controls;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using MetroDemo.Models;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using OfficeInstallGenerator.Model;
using System.Xml;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class ProductView : UserControl
    {
        private LanguagesDialog languagesDialog = null;
        private CancellationTokenSource _tokenSource = new CancellationTokenSource();

        public event TransitionTabEventHandler TransitionTab;
        public event MessageEventHandler InfoMessage;
        public event MessageEventHandler ErrorMessage;

        //private readonly  _versionDownloader = null;
        private Task _downloadTask = null;
        private int _cachedIndex = 0;
        private bool _blockUpdate = false;
        private bool chBxMainProductFirstInitialize = true;
        private List<Build> ProductVersionsAndBuilds;
        

        public ProductView()
        {
            InitializeComponent();
            
        }

        private void ProductView_Loaded(object sender, RoutedEventArgs e)             
        {
            try
            {
                // LoadExcludedProducts();
                cbProject.IsEnabled = false;
                cbVisio.IsEnabled = false;
                CbSkype.IsEnabled = false;     
                //ProductVersionsAndBuilds =       
                if (MainTabControl == null) return;
                MainTabControl.SelectedIndex = 0;

                if (ProductTab.Visibility == Visibility.Collapsed)
                {
                    MainTabControl.SelectedIndex = 1;
                }

                if (GlobalObjects.ViewModel == null) return;
                LanguageList.ItemsSource = GlobalObjects.ViewModel.GetLanguages(null);

                GlobalObjects.ViewModel.PropertyChangeEventEnabled = false;
                LanguageUnique.SelectionChanged -= LanguageUnique_OnSelectionChanged;
                LoadXml();
                LanguageUnique.SelectionChanged += LanguageUnique_OnSelectionChanged;
                GlobalObjects.ViewModel.PropertyChangeEventEnabled = true;

                var currentIndex = ProductBranch.SelectedIndex;
                ProductBranch.ItemsSource = GlobalObjects.ViewModel.Branches;
                if (currentIndex == -1) currentIndex = 0;
                ProductBranch.SelectedIndex = currentIndex;

                if (GlobalObjects.ViewModel.ApplicationMode == ApplicationMode.LanguagePack)
                {
                    MainProducts.ItemsSource = GlobalObjects.ViewModel.LanguagePackProducts;
                    MainProducts.SelectedIndex = 0;

                    ProductsSelectionChanged();

                    UseLangForAllLabel.Visibility = Visibility.Collapsed;
                    UseLangForAllProducts.Visibility = Visibility.Collapsed;
                }
                else
                {
                    MainProducts.ItemsSource = GlobalObjects.ViewModel.MainProducts;
                    MainProducts.SelectedIndex = 0;

                    UseLangForAllLabel.Visibility = Visibility.Visible;
                    UseLangForAllProducts.Visibility = Visibility.Visible;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LogAnaylytics(string path, string pageName)
        {
            try
            {
                GoogleAnalytics.Log(path, pageName);
            }
            catch { }
        }

        private void LanguageChange()
        {
            var languages = GlobalObjects.ViewModel.GetLanguages(GetSelectedProduct());

            LanguageList.ItemsSource = null;
            LanguageList.ItemsSource = languages;
        }

        private void RemoveSelectedLanguage()
        {
            var selectProductId = GetSelectedProduct();
            var installOffice = new InstallOffice();

            var currentItems = (List<Language>)LanguageList.ItemsSource ?? new List<Language>();
            foreach (Language language in LanguageList.SelectedItems)
            {
                if (currentItems.Contains(language))
                {
                    currentItems.Remove(language);
                }

                GlobalObjects.ViewModel.RemoveLanguage(selectProductId, language.Id);

                if (!GlobalObjects.ViewModel.LocalConfig) continue;

                var productId = language.ProductId;
                if (string.IsNullOrEmpty(productId)) productId = "O365ProPlusRetail";

                var languageInstalled = installOffice.ProPlusLanguageInstalled(productId, language.Id);
                if (languageInstalled)
                {
                    GlobalObjects.ViewModel.AddRemovedLanguage(productId, language);
                }
            }

            LanguageList.ItemsSource = null;
            LanguageList.ItemsSource = GlobalObjects.ViewModel.GetLanguages(selectProductId);

            UpdateXml();
        }

        private void LaunchLanguageDialog()
        {
            try
            {
                if (languagesDialog == null)
                {
                    var currentItems1 = (List<Language>)LanguageList.ItemsSource ?? new List<Language>();

                    var languageList = GlobalObjects.ViewModel.Languages.ToList();
                    foreach (var language in currentItems1)
                    {
                        languageList.Remove(language);
                    }

                    languagesDialog = new LanguagesDialog
                    {
                        LanguageSource = languageList
                    };
                    languagesDialog.Closed += (o, args) =>
                    {
                        languagesDialog = null;
                    };
                    languagesDialog.Closing += (o, args) =>
                    {
                        var currentItems2 = (List<Language>)LanguageList.ItemsSource ?? new List<Language>();

                        if (languagesDialog.SelectedItems?.Count > 0)
                        {
                            currentItems2.AddRange(languagesDialog.SelectedItems);
                        }

                        var selectedLangs = FormatLanguage(currentItems2.Distinct().ToList()).ToList();

                        var selectProductId = GetSelectedProduct();

                        foreach (var languages in selectedLangs)
                        {
                            languages.ProductId = selectProductId;
                        }

                        GlobalObjects.ViewModel.AddLanguages(selectProductId, selectedLangs);

                        foreach (var language in selectedLangs)
                        {
                            GlobalObjects.ViewModel.RemoveAddRemovedLanguage(selectProductId, language); 
                        }
                        
                        LanguageList.ItemsSource = null;
                        LanguageList.ItemsSource = selectedLangs;

                        UpdateXml();
                    };
                }
                languagesDialog.Launch();

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }
        
        private void ChangePrimaryLanguage()
        {
            var currentItems = (List<Language>)LanguageList.ItemsSource ?? new List<Language>();
            if (currentItems.Count <= 0) return;
            if (LanguageList.SelectedItems.Count != 1) return;

            var selectedLanguage = LanguageList.SelectedItems.Cast<Language>().FirstOrDefault();

            var selectProductId = GetSelectedProduct();

            GlobalObjects.ViewModel.ChangePrimaryLanguage(selectProductId, selectedLanguage);

            LanguageList.ItemsSource = null;
            LanguageList.ItemsSource = GlobalObjects.ViewModel.GetLanguages(selectProductId);
        }

        public void Reset()
        {
            try
            {
                _blockUpdate = true;
                cbVisio.SelectedIndex = 0;
                cbProject.SelectedIndex = 0;
                CbSkype.SelectedIndex = 0;
                chkVisio.IsChecked = false;
                chkProject.IsChecked = false;                
                if (chBxMainProductFirstInitialize)
                {
                    chkofficeProd.IsChecked = true;
                    chBxMainProductFirstInitialize = false;
                }
                

                MainProducts.SelectedIndex = 0;
                ProductEdition32Bit.IsChecked = true;
                ProductEdition64Bit.IsChecked = false;
                ProductBranch.SelectedIndex = 0;
                ProductVersion.Text = "";
                ProductUpdateSource.Text = "";
                ProductDownloadSource.Text = "";

                UseLangForAllProducts.IsChecked = true;

                GlobalObjects.ViewModel.ClearLanguages();

                LoadExcludedProducts();
            }
            finally
            {
                _blockUpdate = false;
            }
        }

        public void LoadXml()
        {
            var languages = new List<Language>
            {
                GlobalObjects.ViewModel.DefaultLanguage
            };

            Reset();

            try
            {
                _blockUpdate = true;
                var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
                if (configXml.Add != null)
                {
                    if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office32Bit)
                    {
                        ProductEdition32Bit.IsChecked = true;
                        ProductEdition64Bit.IsChecked = false;
                    }
                    if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office64Bit)
                    {
                        ProductEdition32Bit.IsChecked = false;
                        ProductEdition64Bit.IsChecked = true;
                    }

                    ProductVersion.Text = configXml.Add.Version?.ToString() ?? "";
                    ProductUpdateSource.Text = configXml.Add.SourcePath?.ToString() ?? "";
                    ProductDownloadSource.Text = configXml.Add.DownloadPath?.ToString() ?? "";

                    var branchIndex = 0;
                                           
                        foreach (OfficeBranch branchItem in ProductBranch.Items)
                        {
                            if (branchItem.Id?.ToUpper() == configXml.Add.Branch.ToString().ToUpper())
                            {
                                ProductBranch.SelectedIndex = branchIndex;
                                break;
                            }
                            branchIndex++;
                        }
                    

                    branchIndex = 0;
                    foreach (OfficeBranch branchItem in ProductBranch.Items)
                    {
                        if (branchItem.NewName?.ToUpper() == configXml.Add.ODTChannel.ToString().ToUpper())
                        {
                            ProductBranch.SelectedIndex = branchIndex;
                            break;
                        }
                        branchIndex++;
                    }

                    if (configXml.Add.Products != null && configXml.Add.Products.Count > 0)
                    {
                        LanguageList.ItemsSource = null;

                        GlobalObjects.ViewModel.ClearLanguages();

                        var n = 0;
                        foreach (var product in configXml.Add.Products)
                        {
                            var index = 0;
                            foreach (Product item in MainProducts.Items)
                            {
                                if (item.Id.ToUpper() == product.ID.ToUpper())
                                {
                                    break;
                                }
                                index++;
                            }

                            MainProducts.SelectedIndex = index;

                            foreach (Product item in cbVisio.Items)
                            {
                                if (item.Id.ToUpper() != product.ID.ToUpper()) continue;
                                chkVisio.IsChecked = true;
                                cbVisio.IsEnabled = true;
                                cbVisio.SelectedItem = item;
                                break;
                            }

                            foreach (Product item in cbProject.Items)
                            {
                                if (item.Id.ToUpper() != product.ID.ToUpper()) continue;
                                chkProject.IsChecked = true;
                                cbProject.IsEnabled = true;
                                cbProject.SelectedItem = item;
                                break;
                            }

                            foreach (Product item in CbSkype.Items)
                            {
                                if (item.Id.ToUpper() != product.ID.ToUpper()) continue;
                                ChkSkype.IsChecked = true;
                                CbSkype.IsEnabled = true;
                                CbSkype.SelectedItem = item;
                                break;
                            }

                            if (product.Languages != null)
                            {
                                if (n == 0) languages.Clear();

                                var useSameLangs = configXml.Add.IsLanguagesSameForAllProducts();
                                GlobalObjects.ViewModel.UseSameLanguagesForAllProducts = useSameLangs;

                                var order = 1;
                                foreach (var language in product.Languages)
                                {
                                    var languageLookup = GlobalObjects.ViewModel.Languages.FirstOrDefault(
                                        l => l.Id.ToLower() == language.ID.ToLower());
                                    if (languageLookup == null) continue;
                                    string productId = null;

                                    if (!useSameLangs)
                                    {
                                        productId = product.ID;
                                    }

                                    var newLanguage = new Language()
                                    {
                                        Id = languageLookup.Id,
                                        Name = languageLookup.Name,
                                        Order = order,
                                        ProductId = productId
                                    };

                                    GlobalObjects.ViewModel.AddLanguage(productId, newLanguage);

                                    if (n == 0) languages.Add(newLanguage);
                                    order++;
                                }

                                n++;

                                UseLangForAllProducts.IsChecked = useSameLangs;
                            }

                            if (product.ExcludeApps != null)
                            {
                                foreach (var excludedApp in product.ExcludeApps)
                                {
                                    var vmExcludedApp =
                                        GlobalObjects.ViewModel.ExcludeProducts.FirstOrDefault(
                                            e => e.DisplayName.ToLower() == excludedApp.ID.ToLower());
                                    if (vmExcludedApp != null)
                                    {
                                        vmExcludedApp.Included = false;
                                    }
                                }
                            }

                            LoadExcludedProducts();
                        }
                    }
                    else
                    {
                        MainProducts.SelectedIndex = 0;
                    }
                }
                else
                {
                    MainProducts.SelectedIndex = 0;
                    ProductEdition32Bit.IsChecked = true;
                    ProductEdition64Bit.IsChecked = false;
                    ProductBranch.SelectedIndex = 0;
                    ProductVersion.Text = "";
                }

                var distictList = languages.Distinct().ToList();
                LanguageList.ItemsSource = FormatLanguage(distictList);

                LanguageChange();
                
            }
            finally
            {
                _blockUpdate = false;
            }
        }

        public void UpdateXml()
        {
            if (_blockUpdate) return;

            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Add == null)
            {
                configXml.Add = new ODTAdd();
            }            

            var languages = GlobalObjects.ViewModel.GetRemovedLanguages();
            if (languages.Count > 0)
            {
                if (configXml.Remove == null)
                {
                    configXml.Remove = new ODTRemove
                    {
                        Products = new List<ODTProduct>()
                    };

                    foreach (var productId in languages.Select(p => p.ProductId).ToList().Distinct())
                    {
                        var tmpProdId = productId;
                        if (productId == null)
                        {
                            tmpProdId = "O365ProPlusRetail";
                        }

                        var pLanguages = languages.Where(l => l.ProductId == productId);
                        var odtProduct = new ODTProduct()
                        {
                            ID = tmpProdId
                        };
                        if (odtProduct.Languages == null) odtProduct.Languages = new List<ODTLanguage>();

                        foreach (var language in pLanguages)
                        {
                            odtProduct.Languages.Add(new ODTLanguage()
                            {
                                ID = language.Id
                            });
                        }

                        configXml.Remove.Products.Add(odtProduct);

                    }
                }
            }

            if (ProductEdition32Bit.IsChecked.HasValue && ProductEdition32Bit.IsChecked.Value)
            {
                configXml.Add.OfficeClientEdition = OfficeClientEdition.Office32Bit;
            }

            if (ProductEdition64Bit.IsChecked.HasValue && ProductEdition64Bit.IsChecked.Value)
            {
                configXml.Add.OfficeClientEdition = OfficeClientEdition.Office64Bit;
            }

            if (ProductBranch.SelectedItem != null)
            {
                var selectedItem = (OfficeBranch)ProductBranch.SelectedItem;
                //configXml.Add.Branch = selectedItem.Branch;
                switch (selectedItem.Branch)
                {
                    case Branch.Business:
                        configXml.Add.ODTChannel = ODTChannel.Deferred;
                        break;
                    case Branch.Current:
                        configXml.Add.ODTChannel = ODTChannel.Current;
                        break;
                    case Branch.FirstReleaseBusiness:
                        configXml.Add.ODTChannel = ODTChannel.FirstReleaseDeferred;
                        break;
                    case Branch.FirstReleaseCurrent:
                        configXml.Add.ODTChannel = ODTChannel.FirstReleaseCurrent;
                        break;
                }
            }

            if (configXml.Add.Products == null)
            {
                configXml.Add.Products = new List<ODTProduct>();   
            }

            var versionText = "";
            if (ProductVersion.SelectedIndex > -1)
            {
                var version = (Build) ProductVersion.SelectedValue;
                versionText = version.Version;
            }
            else
            {
                versionText = ProductVersion.Text;
            }

            try
            {
                if (!string.IsNullOrEmpty(versionText))
                {
                    Version productVersion = null;
                    Version.TryParse(versionText, out productVersion);
                    configXml.Add.Version = productVersion;
                }
                else
                {
                    configXml.Add.Version = null;
                }
            }
            catch { }

            configXml.Add.SourcePath = ProductUpdateSource.Text.Length > 0 ? ProductUpdateSource.Text : null;
            configXml.Add.DownloadPath = ProductDownloadSource.Text.Length > 0 ? ProductDownloadSource.Text : null;
            
            var mainProduct = (Product) MainProducts.SelectedItem;
            if (mainProduct != null)
            {
                configXml.Add.Products.Clear();

                var existingProduct = new ODTProduct()
                {
                    ID = mainProduct.Id
                };
                if (chkofficeProd.IsChecked.HasValue && chkofficeProd.IsChecked.Value)
                {
                    configXml.Add.Products.Add(existingProduct);
                }

                if (chkVisio.IsChecked.HasValue && chkVisio.IsChecked.Value)
                {
                    var addProduct = (Product) cbVisio.SelectedItem;
                    var additionalProduct = new ODTProduct()
                    {
                        ID = addProduct.Id
                    };
                    configXml.Add.Products.Add(additionalProduct);
                }

                if (chkProject.IsChecked.HasValue && chkProject.IsChecked.Value)
                {
                    var addProduct = (Product)cbProject.SelectedItem;
                    var additionalProduct = new ODTProduct()
                    {
                        ID = addProduct.Id
                    };
                    configXml.Add.Products.Add(additionalProduct);
                }

                if (ChkSkype.IsChecked.HasValue && ChkSkype.IsChecked.Value)
                {
                    var addProduct = (Product)CbSkype.SelectedItem;
                    var additionalProduct = new ODTProduct()
                    {
                        ID = addProduct.Id
                    };
                    configXml.Add.Products.Add(additionalProduct);
                }

                if (existingProduct.Languages == null)
                {
                    existingProduct.Languages = new List<ODTLanguage>();
                }

                foreach (var product in configXml.Add.Products)
                {
                    product.Languages = new List<ODTLanguage>();

                    ODTProduct removeProduct = null;
                    if (configXml.Remove != null)
                    {
                       removeProduct = configXml.Remove.Products.FirstOrDefault(p => p.ID == product.ID);
                    }

                    var productLanguages = GlobalObjects.ViewModel.GetLanguages(GlobalObjects.ViewModel.UseSameLanguagesForAllProducts ? null : product.ID);

                    foreach (var language in productLanguages)
                    {
                        var rmLang = removeProduct?.Languages.FirstOrDefault(l => l.ID.ToLower() == language.Id.ToLower());
                        if (rmLang != null)
                        {
                            removeProduct.Languages.Remove(rmLang);
                            if (removeProduct.Languages.Count == 0)
                            {
                                configXml.Remove.Products.Remove(removeProduct);
                            }
                            if (configXml.Remove.Products.Count == 0)
                            {
                                configXml.Remove = null;
                            }
                        }

                        product.Languages.Add(new ODTLanguage()
                        {
                            ID = language.Id
                        });
                    }

                    var removedLanguages = GlobalObjects.ViewModel.GetRemovedLanguages(product.ID);
                    foreach (var language in removedLanguages)
                    {
                        product.Languages.Add(new ODTLanguage()
                        {
                            ID = language.Id
                        });
                    }
                }

                var excludedApps = ExcludeProducts();

                foreach (var excludedApp in excludedApps)
                {
                    if (existingProduct.ExcludeApps == null)
                    {
                        existingProduct.ExcludeApps = new List<ODTExcludeApp>();
                    }

                    existingProduct.ExcludeApps.Add(new ODTExcludeApp()
                    {
                        ID = excludedApp.DisplayName
                    });
                }

            }   
        }

        public void ChangeBranch(string branchName)
        {
            if (branchName == null) return;
            var selectIndex = 0;
            for (var i = 0; i < ProductBranch.Items.Count; i++)
            {
                var item = (OfficeBranch)ProductBranch.Items[i];
                if (item == null) continue;
                if (item.NewName.ToLower() != branchName.ToLower()) continue;
                selectIndex = i;
                break;
            }

            ProductBranch.SelectedIndex = selectIndex;
        }

        public async Task UpdateVersions()
        {
            var branch = (OfficeBranch)ProductBranch.SelectedItem;
            if (branch == null) return;
            var selectVersion = new OfficeBranch();
            selectVersion.Versions = new List<Build>();
            var selectBuilds = new OfficeBranch();
            selectBuilds.Versions = new List<Build>();
            selectVersion.Versions.AddRange(branch.Versions.GroupBy(v => v.NewVersion).Select(v => v.First()).ToList());
            selectBuilds.Versions.AddRange(branch.Versions.Where(v => v.NewVersion == branch.Versions[0].NewVersion));
            ProductVersion.ItemsSource = branch.Versions;//.Select(v => v.NewVersion);
            
            NewVersion.ItemsSource = selectVersion.Versions;//.Select(b => b.NewBuild);
            ProductBuild.ItemsSource = selectBuilds.Versions;
            ProductVersion.SetValue(TextBoxHelper.WatermarkProperty, branch.CurrentVersion);
            NewVersion.SetValue(TextBoxHelper.WatermarkProperty, branch.Versions[0].NewVersion);
            ProductBuild.SetValue(TextBoxHelper.WatermarkProperty, branch.Versions[0].NewBuild);

            var officeEdition = OfficeEdition.Office32Bit;
            if (ProductEdition64Bit.IsChecked.HasValue && ProductEdition64Bit.IsChecked.Value)
            {
                officeEdition = OfficeEdition.Office64Bit;
            }

            await GetBranchVersion(branch, officeEdition);
        }

        private void LoadExcludedProducts()
        {
            ExcludedApps1.ItemsSource = null;
            ExcludedApps2.ItemsSource = null;

            if (GlobalObjects.ViewModel == null) return;
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser;

            foreach (var excludeApp in GlobalObjects.ViewModel.ExcludeProducts)
            {
                var appIncluded = true;

                if (configXml.ConfigurationXml.Add != null && configXml.ConfigurationXml.Add.Products != null)
                {
                    foreach (var product in configXml.ConfigurationXml.Add.Products)
                    {
                        if (product.ExcludeApps == null) continue;

                        foreach (var e in product.ExcludeApps)
                        {
                            if (e.ID.ToUpper() == excludeApp.DisplayName.ToUpper())
                            {
                                appIncluded = false;
                            }
                        }
                    }
                }

                excludeApp.Included = appIncluded;
            }

            if (GlobalObjects.ViewModel.ExcludeProducts != null)
            {
                var splitCount = Convert.ToInt32(Math.Round((double) GlobalObjects.ViewModel.ExcludeProducts.Count/2, 0));
                ExcludedApps1.ItemsSource = GlobalObjects.ViewModel.ExcludeProducts.Take(splitCount).ToList();
                ExcludedApps2.ItemsSource = GlobalObjects.ViewModel.ExcludeProducts.Skip(splitCount).ToList();
            }
        }

        private IEnumerable<ExcludeProduct> ExcludeProducts()
        {
            if (ExcludedApps1.Items != null && ExcludedApps2.Items != null)
            {
                var excludedProducts = ExcludedApps1.Items.Cast<ExcludeProduct>().ToList();
                excludedProducts.AddRange(ExcludedApps2.Items.Cast<ExcludeProduct>().ToList());
                return excludedProducts.Where(e => e.Status == ExcludedStatus.Excluded).ToList();
            }
            return new List<ExcludeProduct>();
        }

        private IEnumerable<Language> FormatLanguage(List<Language> languages)
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
            return languages.ToList();
        }
        
        private string GetSelectedProduct()
        {
            string selectedProductId = null;
            if (!LanguageUnique.IsEnabled) return null;
            var selectProduct = (Product)LanguageUnique.SelectedItem;
            if (selectProduct != null)
            {
                selectedProductId = selectProduct.Id;
            }
            return selectedProductId;
        }

        private void ProductsSelectionChanged()
        {                        

            LanguageUnique.ItemsSource = null;

            var products = new List<Product>();

            if (chkofficeProd.IsChecked.HasValue && chkofficeProd.IsChecked.Value)
            {
                foreach (Product product in MainProducts.SelectedItems)
                {
                    products.Add(product);
                }
            }
            if (chkVisio.IsChecked.HasValue && chkVisio.IsChecked.Value)
            {
                var visioProduct = (Product)cbVisio.SelectedItem;
                products.Add(visioProduct);
            }

            if (chkProject.IsChecked.HasValue && chkProject.IsChecked.Value)
            {
                var projetProduct = (Product)cbProject.SelectedItem;
                products.Add(projetProduct);
            }

            if (ChkSkype.IsChecked.HasValue && ChkSkype.IsChecked.Value)
            {
                var skypeProduct = (Product)CbSkype.SelectedItem;
                products.Add(skypeProduct);
            }

            LanguageUnique.DisplayMemberPath = "ShortName";
            LanguageUnique.ItemsSource = products;
            if (products.Count > 0)
            {
                LanguageUnique.SelectedIndex = 0;
            }

            UpdateXml();
        }

        private async Task GetBranchVersion(OfficeBranch branch, OfficeEdition officeEdition)
        {
            try
            {
                var modelBranch = GlobalObjects.ViewModel.Branches.FirstOrDefault(b =>
                    b.Branch.ToString().ToLower() == branch.Branch.ToString().ToLower());
                if (modelBranch == null) return;

                ProductVersion.SetValue(TextBoxHelper.WatermarkProperty, modelBranch.CurrentVersion);
            }
            catch (Exception)
            {

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
                        tmpIndex ++;
                        if (tmpIndex < MainTabControl.Items.Count)
                        {
                            var item = (TabItem) MainTabControl.Items[tmpIndex];
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

        private void SetTabStatus(bool enabled)
        {
            Dispatcher.Invoke(() =>
            {
                ProductTab.IsEnabled = enabled;
                LanguagesTab.IsEnabled = enabled;
                OptionalTab.IsEnabled = enabled;
                ExcludedTab.IsEnabled = enabled;
            });
        }

        #region "Events"

        private void CbVisio_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                ProductsSelectionChanged();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void CbProject_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                ProductsSelectionChanged();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void ChkProject_OnChecked(object sender, RoutedEventArgs e)
        {
            try
            {
                cbProject.IsEnabled = (chkProject.IsChecked.HasValue && chkProject.IsChecked.Value);
                ProductsSelectionChanged();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void ChkVisio_OnChecked(object sender, RoutedEventArgs e)
        {
            try
            {
                cbVisio.IsEnabled = (chkVisio.IsChecked.HasValue && chkVisio.IsChecked.Value);                
                ProductsSelectionChanged();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void chkofficeProd_Checked(object sender, RoutedEventArgs e)
        {
            try
            {                
                MainProducts.IsEnabled = (chkofficeProd.IsChecked.HasValue && chkofficeProd.IsChecked.Value);
                ProductsSelectionChanged();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void MainTabControl_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (GlobalObjects.ViewModel.BlockNavigation)
                {
                    MainTabControl.SelectedIndex = _cachedIndex;
                    return;
                }

                UpdateXml();                
                switch (MainTabControl.SelectedIndex)
                {
                    case 0:
                        LogAnaylytics("/ProductView", "Products");
                        break;
                    case 1:
                        LogAnaylytics("/ProductView", "Languages");
                        break;
                    case 2:
                        LogAnaylytics("/ProductView", "Optional");
                        break;
                    case 3:
                        LogAnaylytics("/ProductView", "Excluded");
                        break;
                }

                _cachedIndex = MainTabControl.SelectedIndex;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }
        
        private async void ProductBranch_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (ProductBranch.SelectedItem != null)
                {
                    var branch = (OfficeBranch) ProductBranch.SelectedItem;
                    GlobalObjects.ViewModel.SelectedBranch = branch.Branch.ToString();

                    if (ProductUpdateSource.Text.Length > 0)
                    {
                        var otherFolder = GlobalObjects.SetBranchFolderPath(branch.Branch.ToString(), ProductUpdateSource.Text);
                        if (await GlobalObjects.DirectoryExists(otherFolder))
                        {
                            ProductUpdateSource.Text = GlobalObjects.SetBranchFolderPath(branch.Branch.ToString(), ProductUpdateSource.Text);
                        }
                    }
                }

                await UpdateVersions();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void OpenFolderButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var folderPath = ProductUpdateSource.Text.Trim();
                if (string.IsNullOrEmpty(folderPath)) return;

                if (await GlobalObjects.DirectoryExists(folderPath))
                {
                    Process.Start("explorer", folderPath);
                }
                else
                {
                    MessageBox.Show("Directory path does not exist.");
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void OpenLoggingFolderButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var folderPath = RemoteLoggingPath.Text.Trim();
                if (string.IsNullOrEmpty(folderPath)) return;

                if (await GlobalObjects.DirectoryExists(folderPath))
                {
                    Process.Start("explorer", folderPath);
                }
                else
                {
                    MessageBox.Show("Directory path does not exist.");
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void BuildFilePath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                var enabled = false;
                var openFolderEnabled = false;
                if (ProductUpdateSource.Text.Trim().Length > 0)
                {
                    var match = Regex.Match(ProductUpdateSource.Text, @"^\w:\\|\\\\.*\\..*");
                    if (match.Success)
                    {
                        enabled = true;
                        var folderExists = await GlobalObjects.DirectoryExists(ProductUpdateSource.Text);
                        if (!folderExists)
                        {
                            folderExists = await GlobalObjects.DirectoryExists(ProductUpdateSource.Text);
                        }

                        openFolderEnabled = folderExists;  
                    }
                }

                OpenFolderButton.IsEnabled = openFolderEnabled;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void RemoteLoggingPath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                var enabled = false;
                var openFolderEnabled = false;
                if (RemoteLoggingPath.Text.Trim().Length > 0)
                {
                    var match = Regex.Match(RemoteLoggingPath.Text, @"^\w:\\|\\\\.*\\..*");
                    if (match.Success)
                    {
                        enabled = true;
                        var folderExists = await GlobalObjects.DirectoryExists(RemoteLoggingPath.Text);
                        GlobalObjects.ViewModel.RemoteLoggingPath = RemoteLoggingPath.Text;
                        if (!folderExists)
                        {
                            folderExists = await GlobalObjects.DirectoryExists(RemoteLoggingPath.Text);
                            GlobalObjects.ViewModel.RemoteLoggingPath = "";
                        }


                        openFolderEnabled = folderExists;
                    }
                }

                OpenLoggingFolderButton.IsEnabled = openFolderEnabled;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void RemotePath_Click(object sender, RoutedEventArgs e)
        {
        }

        private void UpdatePath_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg1 = new Ionic.Utils.FolderBrowserDialogEx
                {
                    Description = "Select a folder:",
                    ShowNewFolderButton = true,
                    ShowEditBox = true,
                    SelectedPath = ProductUpdateSource.Text,
                    ShowFullPathInEditBox = true,
                    RootFolder = System.Environment.SpecialFolder.MyComputer
                };
                //dlg1.NewStyle = false;

                // Show the FolderBrowserDialog.
                var result = dlg1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    ProductUpdateSource.Text = dlg1.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void RemoteLoggingPath_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg1 = new Ionic.Utils.FolderBrowserDialogEx
                {
                    Description = "Select a folder:",
                    ShowNewFolderButton = true,
                    ShowEditBox = true,
                    SelectedPath = ProductUpdateSource.Text,
                    ShowFullPathInEditBox = true,
                    RootFolder = System.Environment.SpecialFolder.MyComputer
                };
                //dlg1.NewStyle = false;

                // Show the FolderBrowserDialog.
                var result = dlg1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    RemoteLoggingPath.Text = dlg1.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LanguageUnique_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                LanguageChange();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void UseLangForAllProducts_OnChecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LanguageUnique == null) return;

                if (UseLangForAllProducts.IsChecked.HasValue)
                {
                    GlobalObjects.ViewModel.UseSameLanguagesForAllProducts = UseLangForAllProducts.IsChecked.Value;

                    if (GlobalObjects.ViewModel.UseSameLanguagesForAllProducts)
                    {
                        GlobalObjects.ViewModel.SetProductLanguagesForAll(GetSelectedProduct());
                    }

                    LanguageUnique.IsEnabled = !(UseLangForAllProducts.IsChecked.Value);
                }
                else
                {
                    GlobalObjects.ViewModel.UseSameLanguagesForAllProducts = false;
                    LanguageUnique.IsEnabled = true;
                }

                LanguageChange();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void Products_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                ProductsSelectionChanged();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void ToggleSwitch_OnIsCheckedChanged(object sender, EventArgs e)
        {
            try
            {
                var toggleSwitch = (ToggleSwitch) sender;
                if (toggleSwitch != null)
                {
                    var context = (ExcludeProduct) toggleSwitch.DataContext;
                    if (context != null)
                    {
                        if (toggleSwitch.IsChecked.HasValue)
                        {
                            context.Included = toggleSwitch.IsChecked.Value;

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void AddLanguageButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LaunchLanguageDialog();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void RemoveLanguageButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                RemoveSelectedLanguage();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void SetPrimaryButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                ChangePrimaryLanguage();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void Selector_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (BranchChanged == null) return;
            var selectedBranch = ProductBranch.SelectedItem.ToString();
            this.BranchChanged(this, new BranchChangedEventArgs()
            {
                BranchName = selectedBranch
            });
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateXml();

                if (TransitionProductTabs(TransitionTabDirection.Forward))
                {
                    this.TransitionTab?.Invoke(this, new TransitionTabEventArgs()
                    {
                        Direction = TransitionTabDirection.Forward,
                        Index = 1
                    });
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void PreviousButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateXml();

                if (TransitionProductTabs(TransitionTabDirection.Back))
                {
                    this.TransitionTab(this, new TransitionTabEventArgs()
                    {
                        Direction = TransitionTabDirection.Back,
                        Index = 1
                    });
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }
        

        public BranchChangedEventHandler BranchChanged { get; set; }

        #endregion

        #region "Info"

        private void ProductInfo_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic) sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private InformationDialog informationDialog = null;

        private void LaunchInformationDialog(string sourceName)
        {
            try
            {
                if (informationDialog == null)
                {

                    informationDialog = new InformationDialog
                    {
                        Height = 500,
                        Width = 400
                    };
                    informationDialog.Closed += (o, args) =>
                    {
                        informationDialog = null;
                    };
                    informationDialog.Closing += (o, args) =>
                    {

                    };
                }
                
                informationDialog.Height = 500;
                informationDialog.Width = 400;

                var filePath = AppDomain.CurrentDomain.BaseDirectory + @"HelpFiles\" + sourceName + ".html";
                var helpFile = File.ReadAllText(filePath);

                informationDialog.HelpInfo.NavigateToString(helpFile);
                informationDialog.Launch();

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }






        #endregion

        private void chkofficeProd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if ((chkVisio.IsChecked.HasValue && chkVisio.IsChecked.Value == false) && (chkProject.IsChecked.HasValue && chkProject.IsChecked.Value == false) && (ChkSkype.IsChecked.HasValue && chkProject.IsChecked.Value == false) && (chkofficeProd.IsChecked.HasValue && chkofficeProd.IsChecked.Value == false))
                {

                    GlobalObjects.ViewModel.BlockNavigation = true;
                    throw new Exception("At least one product must be selected.");

                }
                else if (GlobalObjects.ViewModel.BlockNavigation)
                {
                    GlobalObjects.ViewModel.BlockNavigation = false;
                }
            }
            catch(Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void ChkSkype_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                CbSkype.IsEnabled = (ChkSkype.IsChecked.HasValue && ChkSkype.IsChecked.Value);
                ProductsSelectionChanged();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void CbSkype_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                ProductsSelectionChanged();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void ProductDownloadSource_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                var enabled = false;
                var openFolderEnabled = false;
                if (ProductDownloadSource.Text.Trim().Length > 0)
                {
                    var match = Regex.Match(ProductDownloadSource.Text, @"^\w:\\|\\\\.*\\..*");
                    if (match.Success)
                    {
                        enabled = true;
                        var folderExists = await GlobalObjects.DirectoryExists(ProductDownloadSource.Text);
                        if (!folderExists)
                        {
                            folderExists = await GlobalObjects.DirectoryExists(ProductDownloadSource.Text);
                        }

                        openFolderEnabled = folderExists;
                    }
                }

                OpenDownloadFolderButton.IsEnabled = openFolderEnabled;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private async void OpenDownloadFolderButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var folderPath = ProductDownloadSource.Text.Trim();
                if (string.IsNullOrEmpty(folderPath)) return;

                if (await GlobalObjects.DirectoryExists(folderPath))
                {
                    Process.Start("explorer", folderPath);
                }
                else
                {
                    MessageBox.Show("Directory path does not exist.");
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void DownloadPath_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg1 = new Ionic.Utils.FolderBrowserDialogEx
                {
                    Description = "Select a folder:",
                    ShowNewFolderButton = true,
                    ShowEditBox = true,
                    SelectedPath = ProductDownloadSource.Text,
                    ShowFullPathInEditBox = true,
                    RootFolder = System.Environment.SpecialFolder.MyComputer
                };
                //dlg1.NewStyle = false;

                // Show the FolderBrowserDialog.
                var result = dlg1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    ProductDownloadSource.Text = dlg1.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void NewVersion_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.NewVersion.IsDropDownOpen)
            {
                var branch = (OfficeBranch)ProductBranch.SelectedItem;
                if (branch == null) return;
                var selectBuilds = new OfficeBranch();
                if (NewVersion.SelectedValue == null)
                {
                    var selectingVersion = new OfficeBranch();
                    selectingVersion.Versions = new List<Build>();
                    selectingVersion.Versions.AddRange(branch.Versions.GroupBy(v => v.NewVersion).Select(v => v.First()).ToList());
                    NewVersion.SelectedItem = branch.Versions[0];

                }
                selectBuilds.Versions = new List<Build>();
                selectBuilds.Versions.AddRange(branch.Versions.Where(v => v.NewVersion == ((Build)NewVersion.SelectedValue).NewVersion));
                ProductBuild.ItemsSource = selectBuilds.Versions;
                ProductVersion.SelectedItem = null;
                ProductBuild.SetValue(TextBoxHelper.WatermarkProperty, selectBuilds.Versions[0].NewBuild);
                ProductVersion.SetValue(TextBoxHelper.WatermarkProperty, selectBuilds.Versions[0].Version);
            }
        }

        private void ProductBuild_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var branch = (OfficeBranch)ProductBranch.SelectedItem;
            if (branch == null || ProductBuild.SelectedValue == null) return;
            if(NewVersion.SelectedValue == null)
            {
                    var selectingVersion = new OfficeBranch();
                    selectingVersion.Versions = new List<Build>();
                    selectingVersion.Versions.AddRange(branch.Versions.GroupBy(v => v.NewVersion).Select(v => v.First()).ToList());
                    NewVersion.SelectedItem = branch.Versions[0];
                
            }
                var selectVersion = new OfficeBranch();
            selectVersion.Versions = new List<Build>();
            var something = ProductVersion.ItemsSource;//selectVersion.Versions.Where(v => v.NewVersion == ((Build)ProductBuild.SelectedValue).NewBuild);
            foreach (Build prod in ProductVersion.Items)
            {
                if (prod.NewBuild == ((Build)ProductBuild.SelectedValue).NewBuild && ((Build)NewVersion.SelectedValue).NewVersion == prod.NewVersion)
                {
                    ProductVersion.SelectedItem = prod;
                    break;
                }
            }
            //ProductVersion.SelectedIndex = branch.Versions.Select(v => v.NewBuild == ((Build)ProductBuild.SelectedValue).NewBuild);
        }
    }
}

