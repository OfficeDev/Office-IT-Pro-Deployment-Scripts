using System.Windows.Controls;
using System.Windows.Data;
using System.Xml;

namespace XMLViewer
{
    /// <summary>
    /// Interaction logic for Viewer.xaml
    /// </summary>
    public partial class Viewer : UserControl
    {
        private XmlDocument _xmldocument;
        public Viewer()
        {
            InitializeComponent();
        }

        public XmlDocument xmlDocument
        {
            get { return _xmldocument; }
            set
            {
                _xmldocument = value;
                BindXMLDocument();
            }
        }

        private void BindXMLDocument()
        {
            if (_xmldocument == null)
            {
                xmlTree.ItemsSource = null;
                return;
            }

            var provider = new XmlDataProvider {Document = _xmldocument};
            var binding = new Binding {Source = provider, XPath = "child::node()"};
            xmlTree.SetBinding(TreeView.ItemsSourceProperty, binding);
        }
    }
}
