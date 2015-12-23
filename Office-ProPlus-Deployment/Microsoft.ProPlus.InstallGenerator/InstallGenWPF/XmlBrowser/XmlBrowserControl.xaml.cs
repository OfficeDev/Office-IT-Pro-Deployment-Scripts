using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using MetroDemo.Events;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGenerator.Extensions;
using Microsoft.Win32;
using WPFXmlBrowser.Controls;

namespace MahApps.Metro.Controls.XmlBrowser
{
    /// <summary>
    /// Interaction logic for XmlBrowserControl.xaml
    /// </summary>
    public partial class XmlBrowserControl : UserControl
    {
        #region Constructor
        public XmlBrowserControl()
        {
            InitializeComponent();
        }
        #endregion

        #region Dependency Properties
        public static readonly DependencyProperty XmlDocProperty = DependencyProperty.Register("XmlDoc", typeof(string), typeof(XmlBrowserControl), new UIPropertyMetadata(null, OnXmlDocChanged));
        public static readonly DependencyProperty IsEditModeProperty = DependencyProperty.Register("IsEditMode", typeof(bool), typeof(XmlBrowserControl), new UIPropertyMetadata(false, OnIsEditModeChanged));
        public static readonly DependencyProperty XmlSchemaProperty = DependencyProperty.Register("XmlSchema", typeof(string), typeof(XmlBrowserControl), new UIPropertyMetadata(null, OnXmlSchemaChanged));

        #endregion

        #region Properties
        public string XmlDoc
        {
            get
            {
                return (string)GetValue(XmlDocProperty);
            }
            set
            {
                SetValue(XmlDocProperty, value);
            }
        }

        public bool IsEditMode
        {
            get
            {
                return (bool)GetValue(IsEditModeProperty);
            }
            set
            {
                SetValue(IsEditModeProperty, value);
            }
        }

        public string XmlSchema
        {
            get
            {
                return (string)GetValue(XmlSchemaProperty);
            }
            set
            {
                SetValue(XmlSchemaProperty, value);
            }
        }
        #endregion

        #region Property Change Callbacks

        /// <summary>
        /// Executes when IsEditMode DP is changed
        /// </summary>
        /// <param name="d"></param>
        /// <param name="e"></param>
        public static void OnIsEditModeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var browserControl = d as XmlBrowserControl;
            var isEditMode = (bool)e.NewValue;
            if (browserControl == null) return;
            if (isEditMode)
            {
                //browserControl.EditButton.Content = "_save";
                browserControl.WebBrowser.Visibility = Visibility.Collapsed;
                //browserControl.EditText.Visibility = Visibility.Visible;

            }
            else
            {
                //browserControl.EditButton.Content = "_edit";
                browserControl.WebBrowser.Visibility = Visibility.Visible;
                //browserControl.EditText.Visibility = Visibility.Collapsed;
            }
        }

        /// <summary>
        /// Executes when XmlDoc DP is changed, Loads the xml and tranforms it using XSL provided
        /// </summary>
        /// <param name="d"></param>
        /// <param name="e"></param>
        public static void OnXmlDocChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var browserControl = d as XmlBrowserControl;
            if (browserControl == null) return;
            var xmlString = e.NewValue as string;
            
            try
            {

                var xmlDocument = new XmlDocument();

                var xmlDocStyled = new StringBuilder(2500);
                // mark of web - to enable IE to force webpages to run in the security zone of the location the page was saved from
                // http://msdn.microsoft.com/en-us/library/ms537628(v=vs.85).aspx
                xmlDocStyled.Append("<!-- saved from url=(0014)about:internet -->");


                var xslt = new XslCompiledTransform();
                //TODO: Do not forget to change the namespace, if you move the xsl sheet to your application

                var resourceName = typeof (XmlBrowserControl).Assembly.GetManifestResourceNames()
                    .FirstOrDefault(r => r.ToLower().Contains("xml-pretty-print.xsl"));

                var xsltFileStream =
                    typeof(XmlBrowserControl).Assembly.GetManifestResourceStream(
                        resourceName);


                if (xsltFileStream != null)
                {
                    //Load the xsltFile
                    var xmlReader = XmlReader.Create(xsltFileStream);
                    xslt.Load(xmlReader);
                    var settings = new XmlWriterSettings();
                    // writer for transformation
                    var writer = XmlWriter.Create(xmlDocStyled, settings);
                    if (xmlString != null) xmlDocument.LoadXml(xmlString);
                    xslt.Transform(xmlDocument, writer);

                }

                //browserControl.EditText.Text = xmlString;
                browserControl.WebBrowser.NavigateToString(xmlDocStyled.ToString());
                //browserControl.EditButton.Visibility = System.Windows.Visibility.Visible;
                browserControl.CopyClipButton.Visibility = System.Windows.Visibility.Visible;
            }
            catch (Exception ex)
            {
                browserControl.WebBrowser.NavigateToString("Unable to parse xml. Correct the following errors: " + ex.Message);
            }
        }

        /// <summary>
        /// Executes when XmlSchema DP is changed
        /// </summary>
        /// <param name="d"></param>
        /// <param name="e"></param>
        public static void OnXmlSchemaChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var browserControl = d as XmlBrowserControl;
            if (browserControl == null) return;

            if (e.NewValue != null)
                if (!string.IsNullOrEmpty(e.NewValue.ToString()) && !string.IsNullOrEmpty(browserControl.XmlDoc))
                {
                    browserControl.ValidateXmlButton.Visibility = System.Windows.Visibility.Visible;
                    return;
                }
            browserControl.ValidateXmlButton.Visibility = System.Windows.Visibility.Collapsed;
        }
        #endregion

        public InstallOfficeEventHandler InstallOffice { get; set; }

        #region Button Events

        private void SaveToFileButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var fileSave = new SaveFileDialog
                {
                    Filter = "Xml Files (*.xml)|*.xml",
                    FileName = "configuration.xml"
                };

                var saveResult = fileSave.ShowDialog();

                if (saveResult.HasValue && saveResult.Value)
                {
                    File.WriteAllText(fileSave.FileName, XmlDoc.BeautifyXml());
                }
            }
            catch (Exception ex)
            {
                ex.LogException();
            }
        }

        private void InstallOfficeButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (InstallOffice != null)
                {
                    InstallOffice(this, new InstallOfficeEventArgs()
                    {
                        Xml = XmlDoc
                    });
                }
            }
            catch (Exception ex)
            {
                ex.LogException();
            }
        }

        private void EditButtonClick(object sender, RoutedEventArgs e)
        {
            if (!IsEditMode)
            {
                IsEditMode = true;
            }
            else
            {
                // since user is navigating from edit mode, set the edited Text to the xmldoc
                //XmlDoc = EditText.Text;
                IsEditMode = false;
            }
        }

        private void CopyClipButtonClick(object sender, RoutedEventArgs e)
        {
            if (XmlDoc != null)
            {
                Clipboard.SetText(XmlDoc.BeautifyXml());
            }
        }

        private void ValidateXmlButtonClick(object sender, RoutedEventArgs e)
        {
            var xDoc = XDocument.Parse(XmlDoc);
            var xsd = XmlSchema;
            var xmlFileElem = xDoc.Root;
            if (xmlFileElem == null || xsd == null) return;
            var xmlNamespace = xmlFileElem.GetDefaultNamespace().NamespaceName;
            string errors;
            if (new XmlValidator(xmlNamespace, xsd).Validate(xDoc, out errors))
                errors = "No Errors Found!";
            var validationResultControl = new ValidationResultControl
                                              {
                                                  ResultTextBox = { Text = errors }
                                              };

            var validationResultWindow = new Window
                                             {
                                                 WindowStyle = WindowStyle.None,
                                                 Content = validationResultControl,
                                                 WindowStartupLocation = WindowStartupLocation.CenterScreen,
                                                 Height = 200,
                                                 Width = 450
                                             };
            validationResultWindow.ShowDialog();
        }

        #endregion

    }
}
