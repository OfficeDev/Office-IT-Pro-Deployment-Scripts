using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MetroDemo;
using MetroDemo.Events;
using MetroDemo.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config
{
    /// <summary>
    /// Interaction logic for ProductsLanguagesView.xaml
    /// </summary>
    public partial class ProductsLanguagesView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;

        public ProductsLanguagesView()
        {
            InitializeComponent();
        }


        #region events
        private void ProductsLanguagesView_OnLoaded(object sender, RoutedEventArgs e)
        {
           
        }

        private void ChannelVersionPage_OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ToggleNext();
        }

        private void CbProducts_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cbProducts.Text = null;
        }
        private void CbLanguages_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cbLanguages.Text = null;
        }

        private void IncludeProductsToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox)sender;
            var selectedProduct = checkbox.DataContext as Product;

            GlobalObjects.ViewModel.SccmConfiguration.Products.Add(selectedProduct);

            UpdateProductText();
            ToggleNext();
        }

        private void IncludeProductsToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox)sender;
            var unSelectedProduct = checkbox.DataContext as Product;

            foreach (var product in GlobalObjects.ViewModel.SccmConfiguration.Products)
            {
                if (product.DisplayName == unSelectedProduct.DisplayName)
                {
                    GlobalObjects.ViewModel.SccmConfiguration.Products.Remove(product);
                    break;
                }
            }

            UpdateProductText();
            ToggleNext();
        }

        private void CbExcludedProducts_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cbExcludedProducts.Text = null;
        }

        private void ExludeProductsToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox) sender;
            var selectedProduct = checkbox.DataContext as ExcludeProduct; 

            GlobalObjects.ViewModel.SccmConfiguration.ExcludedProducts.Add(selectedProduct);

            UpdateExlcudeProductText();
        }

        private void ExcludeProductsToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox) sender;
            var unSelectedProduct = checkbox.DataContext as ExcludeProduct;

            foreach (var product in GlobalObjects.ViewModel.SccmConfiguration.ExcludedProducts)
            {
                if (product.DisplayName == unSelectedProduct.DisplayName)
                {
                    GlobalObjects.ViewModel.SccmConfiguration.ExcludedProducts.Remove(product);
                    break;
                }
            }

            UpdateExlcudeProductText();
        }

        private void LanguageToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox) sender;
            var selectedLanguage = checkbox.DataContext as Language; 

            GlobalObjects.ViewModel.SccmConfiguration.Languages.Add(selectedLanguage);

            UpdateLanguagesText();
            ToggleNext();
        }

        private void LanguageToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox)sender;
            var unSelectedLanguage = checkbox.DataContext as Language;

            foreach (var language in GlobalObjects.ViewModel.SccmConfiguration.Languages)
            {
                if (language.Id == unSelectedLanguage.Id)
                {
                    GlobalObjects.ViewModel.SccmConfiguration.Languages.Remove(language);
                    break;
                }
            }

            UpdateLanguagesText();
            ToggleNext();
        }

        #endregion

        #region helpers    
        private void UpdateProductText()
        {
            tbSelectedProducts.Text = "Selected: ";
            GlobalObjects.ViewModel.SccmConfiguration.Products.ForEach(p =>
            {
                tbSelectedProducts.Text += p.DisplayName + ", ";
            });
        }

        private void UpdateExlcudeProductText()
        {
            tbExcludedProducts.Text = "Selected: ";
            GlobalObjects.ViewModel.SccmConfiguration.ExcludedProducts.ForEach(p =>
            {
                tbExcludedProducts.Text += p.DisplayName + ", ";
            });
        }

        private void UpdateLanguagesText()
        {
            tbLanguages.Text = "Selected: ";
            GlobalObjects.ViewModel.SccmConfiguration.Languages.ForEach(l =>
            {
                tbLanguages.Text += l.Id + ", ";
            });
        }

        private void ToggleNext()
        {
            var SccmConfig = GlobalObjects.ViewModel.SccmConfiguration;

            if (SccmConfig.Products.Count > 0 && SccmConfig.Languages.Count > 0)
            {
                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = true
                });
            }
            else
            {
                ToggleNextButton?.Invoke(this, new ToggleEventArgs()
                {
                    Enabled = false
                });
            }
        }

        #endregion
    }
}
