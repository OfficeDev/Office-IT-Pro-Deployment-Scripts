using System.Xml.Schema;
using System.Xml.Linq;
using System.Xml;
using System.IO;

namespace WPFXmlBrowser.Controls
{
    internal class XmlValidator
    {
        readonly string _xmlNamespace;
        readonly string _xsdDoc;

        internal XmlValidator(string xmlNamespace, string xsdDoc)
        {
            _xmlNamespace = xmlNamespace;
            _xsdDoc = xsdDoc;
        }

        /// <summary>
        /// Validates Xml against the configured schema and outs validation erros if any
        /// </summary>
        /// <param name="xmlDoc"></param>
        /// <param name="validationErrors"></param>
        /// <returns></returns>
        public bool Validate(XDocument xmlDoc, out string validationErrors)
        {

            validationErrors = null;

            //This will contain the xml schema
            var schemas = new XmlSchemaSet();

            //Add the schema, targetNamespace not specified
            schemas.Add(_xmlNamespace, XmlReader.Create(new StringReader(_xsdDoc)));

            //Variable to check if there are errors
            var hasValidationErrors = false;
            string xmlValidationErrorMessage = null;
            //Validate the xml against the schema
            xmlDoc.Validate(schemas, (sender, eventArgs) =>
                                         {
                                             hasValidationErrors = true;
                                             xmlValidationErrorMessage = eventArgs.Message;
                                         });

            validationErrors = xmlValidationErrorMessage;

            // validation fails if validation errors
            return !hasValidationErrors;

        }
    }
}
