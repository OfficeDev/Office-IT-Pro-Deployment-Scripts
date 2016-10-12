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
        private SccmRemoveLanguages RemovelanguagesDialog = null; 
        private SccmAddProducts AddproductsDialog = null;
        private SccmRemoveProducts RemoveproductsDialog = null;
        private SccmExcludeProducts ExcludeProductsDialog = null; 

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
        #endregion

        #region helpers    
        private void ToggleNext()
        {
            var SccmConfig = GlobalObjects.ViewModel.SccmConfiguration;

            if (SccmConfig.Languages.Count > 0 && SccmConfig.Products.Count > 0)
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
                        AddlanguagesDialog = null;
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


                    AddproductsDialog = new SccmAddProducts
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
                            if (GlobalObjects.ViewModel.SccmConfiguration.Products.IndexOf(p) == -1)
                            {
                                p.ProductAction = ProductAction.Install;
                                GlobalObjects.ViewModel.SccmConfiguration.Products.Add(p);
                            }
                        });
                        AddproductsDialog = null;
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
                    var productList = GlobalObjects.ViewModel.SccmConfiguration.Products.ToList();

                    RemoveproductsDialog = new SccmRemoveProducts
                    {
                        ProductSource = productList
                    };
                    RemoveproductsDialog.Closed += (o, args) =>
                    {
                        AddproductsDialog = null;
                    };
                    RemoveproductsDialog.Closing += (o, args) =>
                    {
                        var selectedProducts = RemoveproductsDialog.SelectedItems;

                        selectedProducts.ForEach(p =>
                        {
                            if(GlobalObjects.ViewModel.SccmConfiguration.Products.IndexOf(p) > -1)  
                            GlobalObjects.ViewModel.SccmConfiguration.Products.Remove(p);
                        });
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
                    var languageList = GlobalObjects.ViewModel.SccmConfiguration.Languages.ToList();

                    RemovelanguagesDialog = new SccmRemoveLanguages()
                    {
                        LanguageSource = languageList
                    };
                    RemovelanguagesDialog.Closed += (o, args) =>
                    {
                        AddlanguagesDialog = null;
                    };
                    RemovelanguagesDialog.Closing += (o, args) =>
                    {

                        var selectedLanguages = RemovelanguagesDialog.SelectedItems;

                        selectedLanguages.ForEach(l =>
                        {
                            if (GlobalObjects.ViewModel.SccmConfiguration.Languages.IndexOf(l) > -1)
                            {
                                GlobalObjects.ViewModel.SccmConfiguration.Languages.Remove(l);
                            }
                        });
                        RemovelanguagesDialog = null;
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


                    ExcludeProductsDialog = new SccmExcludeProducts
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

                            if (GlobalObjects.ViewModel.SccmConfiguration.Products.IndexOf(tempProduct) == -1)
                            {
                                GlobalObjects.ViewModel.SccmConfiguration.Products.Add(tempProduct);
                            }
                        });
                        ExcludeProductsDialog = null;
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
