using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using MetroDemo.ExampleWindows;
using MetroDemo.Models;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGenerator.Models;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config
{
    /// <summary>
    /// Interaction logic for ProductsLanguagesView.xaml
    /// </summary>
    public partial class ProductsLanguagesView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;
        private SccmAddLanguages AddlanguagesDialog = null;
        public event MessageEventHandler ErrorMessage;


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

        private void IncludeProductsToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox)sender;
            var selectedProduct = checkbox.DataContext as Product;

            GlobalObjects.ViewModel.SccmConfiguration.Products.Add(selectedProduct);

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

            ToggleNext();
        }

        private void ExludeProductsToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox) sender;
            var selectedProduct = checkbox.DataContext as ExcludeProduct; 

            GlobalObjects.ViewModel.SccmConfiguration.ExcludedProducts.Add(selectedProduct);

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

        }

        private void LanguageToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var checkbox = (CheckBox) sender;
            var selectedLanguage = checkbox.DataContext as Language; 

            GlobalObjects.ViewModel.SccmConfiguration.Languages.Add(selectedLanguage);

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

            ToggleNext();
        }

        #endregion

        #region helpers    
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


        private void LaunchLanguageDialog()
        {
            try
            {
                if (AddlanguagesDialog == null)
                {
                    var currentItems1 = (ObservableCollection<Language>)LanguageList.ItemsSource ?? new ObservableCollection<Language>();

                    var languageList = GlobalObjects.ViewModel.Languages.ToList();
                    foreach (var language in currentItems1)
                    {
                        languageList.Remove(language);
                    }

                    AddlanguagesDialog = new SccmAddLanguages
                    {
                        LanguageSource = languageList
                    };
                    AddlanguagesDialog.Closed += (o, args) =>
                    {
                        AddlanguagesDialog = null;
                    };
                    AddlanguagesDialog.Closing += (o, args) =>
                    {

                        var selectedLanguages = AddlanguagesDialog.SelectedItems;

                        selectedLanguages.ForEach(l =>
                        {
                            if (GlobalObjects.ViewModel.SccmConfiguration.Languages.IndexOf(l) == -1)
                            {
                                GlobalObjects.ViewModel.SccmConfiguration.Languages.Add(l);
                            }
                        });
                    };
                }
                AddlanguagesDialog.Launch();

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
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

        #endregion

        private void BAddProducts_OnClick(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void BExcludeApps_OnClick(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void BRemoveProduct_OnClick(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void BAddLanguage_OnClick(object sender, RoutedEventArgs e)
        {
            LaunchLanguageDialog();
        }

        private void BRemoveLanguage_OnClick(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }


    }
}
