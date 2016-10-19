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
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Views.CM_Config
{
    /// <summary>
    /// Interaction logic for ProductsLanguagesView.xaml
    /// </summary>
    public partial class ProductsLanguagesView : UserControl
    {
        public event ToggleNextEventHandler ToggleNextButton;
        private CMAddLanguages AddlanguagesDialog = null;
        private CMRemoveLanguages RemovelanguagesDialog = null; 
        private CMAddProducts AddproductsDialog = null;
        private CMRemoveProducts RemoveproductsDialog = null;
        private CMExcludeProducts ExcludeProductsDialog = null;

        public CmProgram CurrentCmProgram = GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1]; 

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

            var grid = (Grid)sender;

            if (grid.Visibility == Visibility.Visible)
            {
                CurrentCmProgram =
                    GlobalObjects.ViewModel.CmPackage.Programs[GlobalObjects.ViewModel.CmPackage.Programs.Count - 1];
            }

            ToggleNext();
        }
        #endregion

        #region helpers    
        private void ToggleNext()
        {
            var CMConfig = GlobalObjects.ViewModel.CmPackage;

            if (CurrentCmProgram.Languages.Count > 0 && CurrentCmProgram.Products.Count > 0)
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
                    var languageList = GlobalObjects.ViewModel.Languages.ToList();

                    AddlanguagesDialog = new CMAddLanguages
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
                            if (CurrentCmProgram.Languages.IndexOf(l) == -1)
                            {
                                CurrentCmProgram.Languages.Add(l);
                            }
                        });
                        AddlanguagesDialog = null;
                        LanguageList.ItemsSource = null;
                        LanguageList.ItemsSource = CurrentCmProgram.Languages;
                        ToggleNext();
                    };
                }
                AddlanguagesDialog.Launch();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LaunchAddProductDialog()
        {
            try
            {
                if (AddproductsDialog == null)
                {
                    var productList = GlobalObjects.ViewModel.AllProductsNoExclude.ToList();


                    AddproductsDialog = new CMAddProducts
                    {
                        ProductSource = productList
                    };
                    AddproductsDialog.Closed += (o, args) =>
                    {
                        AddproductsDialog = null;
                    };
                    AddproductsDialog.Closing += (o, args) =>
                    {
                        var selectedProducts = AddproductsDialog.SelectedItems;

                        selectedProducts.ForEach(p =>
                        {                   
                            if (CurrentCmProgram.Products.IndexOf(p) == -1)
                            {
                                p.ProductAction = ProductAction.Install;
                                CurrentCmProgram.Products.Add(p);
                            }
                        });
                        AddproductsDialog = null;
                        ProductList.ItemsSource = null; 
                        ProductList.ItemsSource = CurrentCmProgram.Products; 
                        ToggleNext();
                    };
                }
                AddproductsDialog.Launch();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LaunchRemoveProductDialog()
        {
            try
            {
                if (RemoveproductsDialog == null)
                {
                    var productList = CurrentCmProgram.Products.ToList();

                    RemoveproductsDialog = new CMRemoveProducts
                    {
                        ProductSource = productList
                    };
                    RemoveproductsDialog.Closed += (o, args) =>
                    {
                        RemoveproductsDialog = null;
                    };
                    RemoveproductsDialog.Closing += (o, args) =>
                    {
                        var selectedProducts = RemoveproductsDialog.SelectedItems;

                        selectedProducts.ForEach(p =>
                        {
                            if(CurrentCmProgram.Products.IndexOf(p) > -1)
                                CurrentCmProgram.Products.Remove(p);
                        });
                        ProductList.ItemsSource = null;
                        ProductList.ItemsSource = CurrentCmProgram.Products;
                        RemoveproductsDialog = null;
                        ToggleNext();
                    };
                }
                RemoveproductsDialog.Launch();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LauncheRemoveLanguagesDialog()
        {
            try
            {
                if (RemovelanguagesDialog == null)
                {
                    var languageList = CurrentCmProgram.Languages.ToList();

                    RemovelanguagesDialog = new CMRemoveLanguages()
                    {
                        LanguageSource = languageList
                    };
                    RemovelanguagesDialog.Closed += (o, args) =>
                    {
                        RemovelanguagesDialog = null;
                    };
                    RemovelanguagesDialog.Closing += (o, args) =>
                    {

                        var selectedLanguages = RemovelanguagesDialog.SelectedItems;

                        selectedLanguages.ForEach(l =>
                        {
                            if (CurrentCmProgram.Languages.IndexOf(l) > -1)
                            {
                                CurrentCmProgram.Languages.Remove(l);
                            }
                        });
                        RemovelanguagesDialog = null;
                        LanguageList.ItemsSource = null;
                        LanguageList.ItemsSource = CurrentCmProgram.Languages;
                        ToggleNext();
                    };
                }
                RemovelanguagesDialog.Launch();

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LaunchExcludeProductsDialog()
        {
            try
            {
                if (ExcludeProductsDialog == null)
                {
                    var productList = GlobalObjects.ViewModel.ExcludeProducts.ToList();


                    ExcludeProductsDialog = new CMExcludeProducts
                    {
                        ProductSource = productList
                    };

                    ExcludeProductsDialog.Closed += (o, args) =>
                    {
                        ExcludeProductsDialog = null;
                    };

                    ExcludeProductsDialog.Closing += (o, args) =>
                    {
                        var selectedProducts = ExcludeProductsDialog.SelectedItems;

                        selectedProducts.ForEach(p =>
                        {
                            var tempProduct = new Product();
                            tempProduct.DisplayName = p.DisplayName;
                            tempProduct.Id = p.DisplayName;
                            tempProduct.ProductAction = ProductAction.Exclude;

                            if (CurrentCmProgram.Products.IndexOf(tempProduct) == -1)
                            {
                                CurrentCmProgram.Products.Add(tempProduct);
                            }
                        });
                        ExcludeProductsDialog = null;
                        ProductList.ItemsSource = null;
                        ProductList.ItemsSource = CurrentCmProgram.Products;
                        ToggleNext();
                    };
                }
                ExcludeProductsDialog.Launch();

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
            LaunchAddProductDialog();
        }

        private void BExcludeApps_OnClick(object sender, RoutedEventArgs e)
        {
            LaunchExcludeProductsDialog();
        }

        private void BRemoveProduct_OnClick(object sender, RoutedEventArgs e)
        {
            LaunchRemoveProductDialog();
        }

        private void BAddLanguage_OnClick(object sender, RoutedEventArgs e)
        {
            LaunchLanguageDialog();
        }

        private void BRemoveLanguage_OnClick(object sender, RoutedEventArgs e)
        {
            LauncheRemoveLanguagesDialog();
        }

    
    }
}
