using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using MetroDemo.Models;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using OfficeInstallGenerator.Model;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class ProductView : UserControl
    {
        private LanguagesDialog languagesDialog = null;
        public event TransitionTabEventHandler TransitionTab;


        public ProductView()
        {
            InitializeComponent();
        }

        private void ProductView_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var splitCount = Convert.ToInt32(Math.Round((double)GlobalObjects.ViewModel.ExcludeProducts.Count / 2, 0));
                ExcludedApps1.ItemsSource = GlobalObjects.ViewModel.ExcludeProducts.Take(splitCount).ToList();
                ExcludedApps2.ItemsSource = GlobalObjects.ViewModel.ExcludeProducts.Skip(splitCount).ToList();

                MainTabControl.SelectedIndex = 0;

                LanguageList.ItemsSource = GlobalObjects.ViewModel.GetLanguages(null);

                LoadXml();
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
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

                        if (languagesDialog.SelectedItems != null)
                        {
                            if (languagesDialog.SelectedItems.Count > 0)
                            {
                                currentItems2.AddRange(languagesDialog.SelectedItems);
                            }
                        }

                        var selectedLangs = FormatLanguage(currentItems2.Distinct().ToList()).ToList();

                        string selectProductId = null;
                        if (LanguageUnique.IsEnabled)
                        {
                            selectProductId = ((Product)LanguageUnique.SelectedItem).Id;
                        }

                        foreach (var languages in selectedLangs)
                        {
                            languages.ProductId = selectProductId;
                        }

                        GlobalObjects.ViewModel.AddLanguages(selectProductId, selectedLangs);

                        LanguageList.ItemsSource = null;
                        LanguageList.ItemsSource = selectedLangs;

                    };
                }
                languagesDialog.Launch();

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void LanguageChange()
        {
            LanguageList.ItemsSource = null;

            string selectedProductId = null;
            if (LanguageUnique.IsEnabled)
            {
                var selectProduct = (Product) LanguageUnique.SelectedItem;
                if (selectProduct != null)
                {
                    selectedProductId = selectProduct.Id;
                }
            }

            var languages = GlobalObjects.ViewModel.GetLanguages(selectedProductId);
            LanguageList.ItemsSource = languages;
        }

        private void RemoveSelectedLanguage()
        {
            string selectProductId = null;
            if (LanguageUnique.IsEnabled)
            {
                selectProductId = ((Product)LanguageUnique.SelectedItem).Id;
            }

            var currentItems = (List<Language>)LanguageList.ItemsSource ?? new List<Language>();
            foreach (Language language in LanguageList.SelectedItems)
            {
                if (currentItems.Contains(language))
                {
                    currentItems.Remove(language);
                }



                GlobalObjects.ViewModel.RemoveLanguage(selectProductId, language.Id);
            }
            LanguageList.ItemsSource = null;
            LanguageList.ItemsSource = GlobalObjects.ViewModel.GetLanguages(selectProductId);
        }

        private void ChangePrimaryLanguage()
        {
            var currentItems = (List<Language>)LanguageList.ItemsSource ?? new List<Language>();
            if (currentItems.Count <= 0) return;
            if (LanguageList.SelectedItems.Count != 1) return;

            var selectedLanguage = LanguageList.SelectedItems.Cast<Language>().FirstOrDefault();

            string selectProductId = null;
            if (LanguageUnique.IsEnabled)
            {
                selectProductId = ((Product)LanguageUnique.SelectedItem).Id;
            }

            GlobalObjects.ViewModel.ChangePrimaryLanguage(selectProductId, selectedLanguage);

            LanguageList.ItemsSource = null;
            LanguageList.ItemsSource = GlobalObjects.ViewModel.GetLanguages(selectProductId);
        }

        public void LoadXml()
        {
            var languages = new List<Language>
                {
                    new Language()
                    {
                        Id = "en-us",
                        Name = "English"
                    }
                };

            AdditionalProducts.SelectedItems.Clear();

            MainProducts.SelectedIndex = 0;
            ProductEdition32Bit.IsChecked = true;
            ProductEdition64Bit.IsChecked = false;
            ProductBranch.SelectedIndex = 0;

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

                ProductVersion.Text = configXml.Add.Version != null ? configXml.Add.Version.ToString() : "";
                ProductUpdateSource.Text = configXml.Add.SourcePath != null ? configXml.Add.SourcePath.ToString() : "";

                var branchIndex = 0;
                foreach (ComboBoxItem branchItem in ProductBranch.Items)
                {
                    if (branchItem.Tag.ToString().ToUpper() == configXml.Add.Branch.ToString().ToUpper())
                    {
                        ProductBranch.SelectedIndex = branchIndex;
                        break;
                    }
                    branchIndex++;
                }

                if (configXml.Add.Products != null && configXml.Add.Products.Count > 0)
                {
                    languages = new List<Language>();
                    foreach (var product in configXml.Add.Products)
                    {
                        var index = 0;
                        foreach (Product item in MainProducts.Items)
                        {
                            if (item.Id.ToUpper() == product.ID.ToUpper())
                            {
                                break;
                            }
                            index ++;
                        }

                        MainProducts.SelectedIndex = index;

                        foreach (Product item in AdditionalProducts.Items)
                        {
                            if (item.Id.ToUpper() != product.ID.ToUpper()) continue;
                            AdditionalProducts.SelectedItems.Add(item);
                            break;
                        }

                        if (product.Languages != null)
                        {
                            foreach (var language in product.Languages)
                            {
                                var languageLookup =
                                    GlobalObjects.ViewModel.Languages.FirstOrDefault(l => l.Id.ToLower() == language.ID.ToLower());
                                if (languageLookup != null)
                                {
                                    languages.Add(new Language()
                                    {
                                        Id = languageLookup.Id,
                                        Name = languageLookup.Name
                                    });
                                }
                            }
                        }


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

            }

            LanguageList.ItemsSource = null;
            var distictList = languages.Distinct().ToList();
            LanguageList.ItemsSource = FormatLanguage(distictList);
        }

        public void UpdateXml()
        {
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Add == null)
            {
                configXml.Add = new ODTAdd();
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
                var selectedItem = (ComboBoxItem) ProductBranch.SelectedItem;
                configXml.Add.Branch = (Branch) Enum.Parse(typeof(Branch), selectedItem.Tag.ToString());
            }

            if (configXml.Add.Products == null)
            {
                configXml.Add.Products = new List<ODTProduct>();   
            }

            var mainProduct = (Product) MainProducts.SelectedItem;
            if (mainProduct != null)
            {
                configXml.Add.Products.Clear();

                var existingProduct = new ODTProduct()
                {
                    ID = mainProduct.Id
                };

                configXml.Add.Products.Add(existingProduct);

                foreach (Product addProduct in AdditionalProducts.SelectedItems)
                {
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

                    foreach (Language language in LanguageList.Items)
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
                    if (existingProduct.ExcludedApps == null)
                    {
                        existingProduct.ExcludedApps = new List<ODTExcludedApp>();
                    }

                    existingProduct.ExcludedApps.Add(new ODTExcludedApp()
                    {
                        ID = excludedApp.Id
                    });
                }

            }

            var xml = GlobalObjects.ViewModel.ConfigXmlParser.Xml;
            if (xml != null)
            {

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


        #region "Events"

        private void LanguageUnique_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                LanguageChange();
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
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
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void Products_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                LanguageUnique.ItemsSource = null;

                var products = new List<Product>();

                foreach (Product product in MainProducts.SelectedItems)
                {
                    products.Add(product);
                }

                foreach (Product product in AdditionalProducts.SelectedItems)
                {
                    products.Add(product);
                }

                LanguageUnique.DisplayMemberPath = "DisplayName";
                LanguageUnique.ItemsSource = products;
                if (products.Count > 0)
                {
                    LanguageUnique.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void MainTabControl_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
 
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void ToggleSwitch_OnIsCheckedChanged(object sender, EventArgs e)
        {
            try
            {
                var toggleSwitch = (ToggleSwitch) sender;

                if (toggleSwitch != null)
                {
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
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
                MessageBox.Show("ERROR: " + ex.Message);
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
                MessageBox.Show("ERROR: " + ex.Message);
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
                MessageBox.Show("ERROR: " + ex.Message);
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

                this.TransitionTab(this, new TransitionTabEventArgs()
                {
                    Direction = TransitionTabDirection.Forward,
                    Index = 1
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void PreviousButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateXml();

                this.TransitionTab(this, new TransitionTabEventArgs()
                {
                    Direction = TransitionTabDirection.Back,
                    Index = 1
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }
        
        public BranchChangedEventHandler BranchChanged { get; set; }

        #endregion


    }
}
