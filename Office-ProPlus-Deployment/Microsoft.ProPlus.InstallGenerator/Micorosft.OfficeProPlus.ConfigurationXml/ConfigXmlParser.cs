using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Enums;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using OfficeInstallGenerator.Model;

namespace OfficeInstallGenerator
{
    public class ConfigXmlParser
    {
        private XmlDocument _xmlDoc = null;
        public ConfigurationXml ConfigurationXml { get; set; }
        public Guid ObjectId;


        public ConfigXmlParser(string xml)
        {
            ObjectId = Guid.NewGuid();

            LoadXml(xml);
        }

        public void LoadXml(string xml)
        {
            if (File.Exists(xml))
            {
                xml = File.ReadAllText(xml);
            }

            _xmlDoc = new XmlDocument();
            _xmlDoc.LoadXml(xml);

            LoadConfigurationXml();
        }

        public string Xml
        {
            get
            {
                SaveProducts();
                SaveUpdates();
                SaveDisplay();
                SaveProperties();
                SaveLogging();
                
                return _xmlDoc.OuterXml;
            }
        }

        private void LoadConfigurationXml()
        {
            if (_xmlDoc.DocumentElement == null)
            {
                throw (new Exception("Document Element Missing"));
            }

            ConfigurationXml = new ConfigurationXml();

            LoadAdds();

            LoadUpdates();

            LoadDisplay();

            LoadLogging();

            LoadProperties();
        }

        private void LoadAdds()
        {
            var addNodes = _xmlDoc.DocumentElement.SelectNodes("./Add");
            foreach (XmlNode addNode in addNodes)
            {
                var odtAdd = new ODTAdd();

                if (addNode.Attributes["OfficeClientEdition"] != null)
                {
                    var officeEdition = addNode.Attributes["OfficeClientEdition"].Value;
                    if (officeEdition == "32")
                    {
                        odtAdd.OfficeClientEdition = OfficeClientEdition.Office32Bit;
                    }
                    if (officeEdition == "64")
                    {
                        odtAdd.OfficeClientEdition = OfficeClientEdition.Office64Bit;
                    }
                }

                if (addNode.Attributes["Branch"] != null)
                {
                    var branch = addNode.Attributes["Branch"].Value;
                    if (!string.IsNullOrEmpty(branch))
                    {
                        odtAdd.Branch = (Branch)Enum.Parse(typeof(Branch), branch);
                    }
                }

                odtAdd.SourcePath = null;
                if (addNode.Attributes["SourcePath"] != null)
                {
                    var sourcePath = addNode.Attributes["SourcePath"].Value;
                    if (!string.IsNullOrEmpty(sourcePath))
                    {
                        odtAdd.SourcePath = sourcePath;
                    }
                }

                odtAdd.Version = null;
                if (addNode.Attributes["Version"] != null)
                {
                    var version = addNode.Attributes["Version"].Value;
                    if (!string.IsNullOrEmpty(version))
                    {
                        odtAdd.Version = new Version(version);
                    }
                }

                ConfigurationXml.Add = odtAdd;

                LoadProducts(addNode, odtAdd);
            }
        }


        private void LoadProducts(XmlNode xmlNode, ODTAdd addItem)
        {
            var products = xmlNode.SelectNodes("./Product");
            foreach (XmlNode productNode in products)
            {
                var product = new ODTProduct();

                if (productNode.Attributes["ID"] != null)
                {
                    var productId = productNode.Attributes["ID"].Value;
                    if (!string.IsNullOrEmpty(productId))
                    {
                        product.ID = productId;
                    }
                }

                if (addItem.Products == null)
                {
                    addItem.Products = new List<ODTProduct>();
                }

                addItem.Products.Add(product);

                LoadLanguages(productNode, product);

                LoadExcludedApps(productNode, product);
            }


        }

        private void SaveProducts()
        {
            var addNode = _xmlDoc.DocumentElement.SelectSingleNode("./Add");
            if (addNode == null)
            {
                addNode = _xmlDoc.CreateElement("Add");
                _xmlDoc.DocumentElement.AppendChild(addNode);
            }

            foreach (XmlNode childNode in addNode.ChildNodes)
            {
                addNode.RemoveChild(childNode);
            }

            if (this.ConfigurationXml.Add != null)
            {
                if (this.ConfigurationXml.Add.Version != null)
                {
                    SetAttribute(addNode, "Version", this.ConfigurationXml.Add.Version.ToString());
                }
                else
                {
                    RemoveAttribute(addNode, "Version");
                }

                SetAttribute(addNode, "OfficeClientEdition",
                    this.ConfigurationXml.Add.OfficeClientEdition == OfficeClientEdition.Office32Bit ? "32" : "64");

                SetAttribute(addNode, "Branch", this.ConfigurationXml.Add.Branch.ToString());

                if (this.ConfigurationXml.Add.SourcePath != null)
                {
                    SetAttribute(addNode, "SourcePath", this.ConfigurationXml.Add.SourcePath);
                }
                else
                {
                    RemoveAttribute(addNode, "SourcePath");
                }
            }

            if (this.ConfigurationXml.Add != null && this.ConfigurationXml.Add.Products != null)
            {
                foreach (var product in this.ConfigurationXml.Add.Products)
                {
                    var productNode = addNode.SelectSingleNode("./Product[@ID='" + product.ID + "']");
                    if (productNode == null)
                    {
                        productNode = _xmlDoc.CreateElement("Product");
                        SetAttribute(productNode, "ID", product.ID);
                        addNode.AppendChild(productNode);
                    }

                    if (product.Languages != null)
                    {
                        foreach (var language in product.Languages)
                        {
                            var languageNode = productNode.SelectSingleNode("./Language[@ID='" + language.ID + "']");
                            if (languageNode == null)
                            {
                                languageNode = _xmlDoc.CreateElement("Language");
                                SetAttribute(languageNode, "ID", language.ID);
                                productNode.AppendChild(languageNode);
                            }
                        }

                        if (product.ExcludeApps != null)
                        {
                            foreach (var excludedApp in product.ExcludeApps)
                            {
                                var excludeAppNode =
                                    productNode.SelectSingleNode("./ExcludeApp[@ID='" + excludedApp.ID + "']");
                                if (excludeAppNode == null)
                                {
                                    excludeAppNode = _xmlDoc.CreateElement("ExcludeApp");
                                    SetAttribute(excludeAppNode, "ID", excludedApp.ID);
                                    productNode.AppendChild(excludeAppNode);
                                }
                            }
                        }


                    }
                }
            }
        }


        private void LoadExcludedApps(XmlNode xmlNode, ODTProduct addItem)
        {
            var excludeApps = xmlNode.SelectNodes("./ExcludeApp");
            foreach (XmlNode excludeAppNode in excludeApps)
            {
                var excludedApp = new ODTExcludeApp();

                if (excludeAppNode.Attributes["ID"] == null) continue;

                var productId = excludeAppNode.Attributes["ID"].Value;
                if (!string.IsNullOrEmpty(productId))
                {
                    excludedApp.ID = productId;
                }

                if (addItem.ExcludeApps == null)
                {
                    addItem.ExcludeApps = new List<ODTExcludeApp>();
                }

                addItem.ExcludeApps.Add(excludedApp);
            }
        }


        private void LoadLanguages(XmlNode xmlNode, ODTProduct addItem)
        {
            var languages = xmlNode.SelectNodes("./Language");
            foreach (XmlNode languageNode in languages)
            {
                var language = new ODTLanguage();
                if (addItem.Languages == null)
                {
                    addItem.Languages = new List<ODTLanguage>();
                }

                if (languageNode.Attributes["ID"] != null)
                {
                    var languageId = languageNode.Attributes["ID"].Value;
                    if (!string.IsNullOrEmpty(languageId))
                    {
                        language.ID = languageId;
                    }
                }


                addItem.Languages.Add(language);
            }
        }


        private void LoadUpdates()
        {
            var updatesNode = _xmlDoc.DocumentElement.SelectSingleNode("./Updates");

            var updates = new ODTUpdates();
            ConfigurationXml.Updates = updates;

            if (updatesNode == null) return;

            if (updatesNode.Attributes["Enabled"] != null)
            {
                var enabled = updatesNode.Attributes["Enabled"].Value;
                if (!string.IsNullOrEmpty(enabled))
                {
                    if (enabled.ToLower() == "true" || enabled.ToLower() == "false")
                    {
                        updates.Enabled = Convert.ToBoolean(enabled);
                    }
                }
            }

            if (updatesNode.Attributes["UpdatePath"] != null)
            {
                var updatePath = updatesNode.Attributes["UpdatePath"].Value;
                if (!string.IsNullOrEmpty(updatePath))
                {
                    updates.UpdatePath = updatePath;
                }
            }

            if (updatesNode.Attributes["Deadline"] != null)
            {
                var deadline = updatesNode.Attributes["Deadline"].Value;
                if (!string.IsNullOrEmpty(deadline))
                {
                    updates.Deadline = DateTime.Parse(deadline);
                }
            }

            if (updatesNode.Attributes["TargetVersion"] != null)
            {
                var targetVersion = updatesNode.Attributes["TargetVersion"].Value;
                if (!string.IsNullOrEmpty(targetVersion))
                {
                    updates.TargetVersion = new Version(targetVersion);
                }
            }

            if (updatesNode.Attributes["Branch"] != null)
            {
                var branch = updatesNode.Attributes["Branch"].Value;
                if (!string.IsNullOrEmpty(branch))
                {
                    updates.Branch = (Branch)Enum.Parse(typeof(Branch), branch);
                }
            }
        }

        private void SaveUpdates()
        {
            var updatesNode = _xmlDoc.DocumentElement.SelectSingleNode("./Updates");
            if (updatesNode == null)
            {
               updatesNode = _xmlDoc.CreateElement("Updates");
               _xmlDoc.DocumentElement.AppendChild(updatesNode);
            }

            SetAttribute(updatesNode, "Enabled", this.ConfigurationXml.Updates.Enabled.ToString().ToUpper());

            if (!this.ConfigurationXml.Updates.Enabled)
            {
                RemoveAttribute(updatesNode, "Branch");
                RemoveAttribute(updatesNode, "UpdatePath");
                RemoveAttribute(updatesNode, "TargetVersion");
                RemoveAttribute(updatesNode, "Deadline");
            }

            if (this.ConfigurationXml.Updates.Branch.HasValue &&
                !string.IsNullOrEmpty(this.ConfigurationXml.Updates.Branch.Value.ToString()))
            {
                SetAttribute(updatesNode, "Branch", this.ConfigurationXml.Updates.Branch.ToString());
            }
            else
            {
                RemoveAttribute(updatesNode, "Branch");
            }

            if (!string.IsNullOrEmpty(this.ConfigurationXml.Updates.UpdatePath))
            {
                SetAttribute(updatesNode, "UpdatePath", this.ConfigurationXml.Updates.UpdatePath.ToString());
            }
            else
            {
                RemoveAttribute(updatesNode, "UpdatePath");
            }

            if (this.ConfigurationXml.Updates.TargetVersion != null)
            {
                SetAttribute(updatesNode, "TargetVersion", this.ConfigurationXml.Updates.TargetVersion.ToString());
            }
            else
            {
                RemoveAttribute(updatesNode, "TargetVersion");
            }

            if (this.ConfigurationXml.Updates.Deadline.HasValue && !string.IsNullOrEmpty(this.ConfigurationXml.Updates.Deadline.Value.ToString()))
            {
                SetAttribute(updatesNode, "Deadline", this.ConfigurationXml.Updates.Deadline.Value.ToString("MM/dd/yyyy, HH:mm"));
            }
            else
            {
                RemoveAttribute(updatesNode, "Deadline");
            }
        }


        private void LoadDisplay()
        {
            var displayNode = _xmlDoc.DocumentElement.SelectSingleNode("./Display");
        
            var display = new ODTDisplay();
            ConfigurationXml.Display = display;

            if (displayNode == null) return;

            if (displayNode.Attributes["AcceptEULA"] != null)
            {
                var enabled = displayNode.Attributes["AcceptEULA"].Value;
                if (!string.IsNullOrEmpty(enabled))
                {
                    if (enabled.ToLower() == "true" || enabled.ToLower() == "false")
                    {
                        display.AcceptEULA = Convert.ToBoolean(enabled);
                    }
                }
            }

            if (displayNode.Attributes["Level"] != null)
            {
                var level = displayNode.Attributes["Level"].Value;
                if (!string.IsNullOrEmpty(level))
                {
                    display.Level = (DisplayLevel)Enum.Parse(typeof(DisplayLevel), level); ;
                }
            }
        }

        private void SaveDisplay()
        {
            var displayNode = _xmlDoc.DocumentElement.SelectSingleNode("./Display");
            if (displayNode == null)
            {
                displayNode = _xmlDoc.CreateElement("Display");
                _xmlDoc.DocumentElement.AppendChild(displayNode);
            }

            if (this.ConfigurationXml.Display.Level != null && !string.IsNullOrEmpty(this.ConfigurationXml.Display.Level.Value.ToString()))
            {
                SetAttribute(displayNode, "Level", this.ConfigurationXml.Display.Level.ToString());
            }
            else
            {
                RemoveAttribute(displayNode, "Level");
            }

            if (this.ConfigurationXml.Display.AcceptEULA != null && !string.IsNullOrEmpty(this.ConfigurationXml.Display.AcceptEULA.Value.ToString()))
            {
                SetAttribute(displayNode, "AcceptEULA", this.ConfigurationXml.Display.AcceptEULA.ToString());
            }
            else
            {
                RemoveAttribute(displayNode, "AcceptEULA");
            }
        }




        private void LoadProperties()
        {
            var properties = new ODTProperties();
            ConfigurationXml.Properties = properties;

            var autoActivateNode = _xmlDoc.DocumentElement.SelectSingleNode("./Property[@Name='AUTOACTIVATE']");
            if (autoActivateNode != null)
            {
                if (autoActivateNode.Attributes["Value"] != null)
                {
                    properties.AutoActivate = YesNo.No;
                    var value = autoActivateNode.Attributes["Value"].Value.ToString();
                    if (value == "1") properties.AutoActivate = YesNo.Yes; 
                }
            }

            var forceAppShutdownNode = _xmlDoc.DocumentElement.SelectSingleNode("./Property[@Name='FORCEAPPSHUTDOWN']");
            if (forceAppShutdownNode != null)
            {
                properties.ForceAppShutdown = false;
                var value = forceAppShutdownNode.Attributes["Value"].Value.ToString();
                if (value.ToUpper() == "TRUE") properties.ForceAppShutdown = true; 
            }


            var sharedComputerLicensing = _xmlDoc.DocumentElement.SelectSingleNode("./Property[@Name='SharedComputerLicensing']");
            if (sharedComputerLicensing != null)
            {
                properties.SharedComputerLicensing = false;
                var value = sharedComputerLicensing.Attributes["Value"].Value.ToString();
                if (value.ToUpper() == "1") properties.SharedComputerLicensing = true; 
            }
        }


        private void SaveProperties()
        {
            if (this.ConfigurationXml.Properties != null)
            {
                if (this.ConfigurationXml.Properties.AutoActivate != null)
                {
                    var autoActivateNode = _xmlDoc.DocumentElement.SelectSingleNode("./Property[@Name='AUTOACTIVATE']");
                    if (autoActivateNode == null)
                    {
                        autoActivateNode = _xmlDoc.CreateElement("Property");
                        SetAttribute(autoActivateNode, "Name", "AUTOACTIVATE");
                        _xmlDoc.DocumentElement.AppendChild(autoActivateNode);
                    }

                    SetAttribute(autoActivateNode, "Value",
                        this.ConfigurationXml.Properties.AutoActivate == YesNo.Yes ? "1" : "0");
                }

                if (this.ConfigurationXml.Properties.ForceAppShutdown.HasValue)
                {
                    var forceAppShutdownNode =
                        _xmlDoc.DocumentElement.SelectSingleNode("./Property[@Name='FORCEAPPSHUTDOWN']");
                    if (forceAppShutdownNode == null)
                    {
                        forceAppShutdownNode = _xmlDoc.CreateElement("Property");
                        SetAttribute(forceAppShutdownNode, "Name", "FORCEAPPSHUTDOWN");
                        _xmlDoc.DocumentElement.AppendChild(forceAppShutdownNode);
                    }

                    SetAttribute(forceAppShutdownNode, "Value",
                        this.ConfigurationXml.Properties.ForceAppShutdown.Value.ToString().ToUpper());
                }


                if (this.ConfigurationXml.Properties.SharedComputerLicensing.HasValue)
                {
                    var sharedComputerLicensing =
                        _xmlDoc.DocumentElement.SelectSingleNode("./Property[@Name='SharedComputerLicensing']");
                    if (sharedComputerLicensing == null)
                    {
                        sharedComputerLicensing = _xmlDoc.CreateElement("Property");
                        SetAttribute(sharedComputerLicensing, "Name", "SharedComputerLicensing");
                        _xmlDoc.DocumentElement.AppendChild(sharedComputerLicensing);
                    }

                    SetAttribute(sharedComputerLicensing, "Value",
                         this.ConfigurationXml.Properties.SharedComputerLicensing == true ? "1" : "0");
                }


            }
        }


        private void LoadLogging()
        {
            var loggingNode = _xmlDoc.DocumentElement.SelectSingleNode("./Logging");

            var logging = new ODTLogging();
            ConfigurationXml.Logging = logging;

            if (loggingNode == null) return;

            if (loggingNode.Attributes["Path"] != null)
            {
                var path = loggingNode.Attributes["Path"].Value;
                if (!string.IsNullOrEmpty(path))
                {
                    logging.Path = path;
                }
            }

            if (loggingNode.Attributes["Level"] != null)
            {
                var level = loggingNode.Attributes["Level"].Value;
                if (!string.IsNullOrEmpty(level))
                {
                    logging.Level = (LoggingLevel)Enum.Parse(typeof(LoggingLevel), level);
                }
            }
        }

        private void SaveLogging()
        {
            var loggingNode = _xmlDoc.DocumentElement.SelectSingleNode("./Logging");
            if (loggingNode == null)
            {
                loggingNode = _xmlDoc.CreateElement("Logging");
                _xmlDoc.DocumentElement.AppendChild(loggingNode);
            }

            if (!string.IsNullOrEmpty(this.ConfigurationXml.Logging.Level.ToString()))
            {
                SetAttribute(loggingNode, "Level", this.ConfigurationXml.Logging.Level.ToString());
            }
            else
            {
                RemoveAttribute(loggingNode, "Level");
            }

            if (!string.IsNullOrEmpty(this.ConfigurationXml.Logging.Path))
            {
                SetAttribute(loggingNode, "Path", this.ConfigurationXml.Logging.Path);
            }
            else
            {
                RemoveAttribute(loggingNode, "Path");
            }
        }



        private void SetAttribute(XmlNode xmlNode, string name, string value)
        {
            var pathAttr = xmlNode.Attributes[name];
            if (pathAttr == null)
            {
                pathAttr = _xmlDoc.CreateAttribute(name);
                xmlNode.Attributes.Append(pathAttr);
            }
            pathAttr.Value = value;
        }

        private void RemoveAttribute(XmlNode xmlNode, string name)
        {
            var pathAttr = xmlNode.Attributes[name];
            if (pathAttr != null)
            {
                xmlNode.Attributes.Remove(pathAttr);
            }
        }


    }
}
