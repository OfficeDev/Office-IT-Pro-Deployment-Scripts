using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace SelfServiceConfigXmlEditor
{
    public class SelfServiceConfig
    {

        private XmlDocument xmlDocument { get; set; }

        public SelfServiceConfig()
        {
            
        }

        public void Load(string xml)
        {
            xmlDocument = new XmlDocument();
            if (File.Exists(xml))
            {
                xmlDocument.Load(xml);
            }
            else
            {
                xmlDocument.LoadXml(xml);
            }
        }

        public void Save(string xmlFilePath)
        {
            xmlDocument?.Save(xmlFilePath);
        }

        public string Company
        {
            get
            {
                var company = xmlDocument.DocumentElement.SelectSingleNode("./Company");
                if (company == null)
                {
                    company = xmlDocument.CreateElement("Company");
                    xmlDocument.DocumentElement.AppendChild(company);
                }
                return GetXmlAttribute(company, "Name");
            }
            set
            {
                var company = xmlDocument.DocumentElement.SelectSingleNode("./Company");
                if (company == null)
                {
                    company = xmlDocument.CreateElement("Company");
                    xmlDocument.DocumentElement.AppendChild(company);
                }
                SetXmlAttribute(company, "Name", value);
            }
        }

        public string Title
        {
            get
            {
                var title = xmlDocument.DocumentElement.SelectSingleNode("./Banner/Item/Title");
                if (title == null)
                {
                    title = xmlDocument.CreateElement("Title");
                    xmlDocument.DocumentElement.AppendChild(title);
                }
                return title.InnerText;
            }
            set
            {
                var title = xmlDocument.DocumentElement.SelectSingleNode("./Banner/Item/Title");
                if (title == null)
                {
                    title = xmlDocument.CreateElement("Title");
                    xmlDocument.DocumentElement.AppendChild(title);
                }
                title.InnerText = value;
            }
        }

        public string MainText
        {
            get
            {
                var banner = xmlDocument.DocumentElement.SelectSingleNode("./Banner");
                if (banner == null)
                {
                    banner = xmlDocument.CreateElement("Banner");
                    xmlDocument.DocumentElement.AppendChild(banner);
                }

                var item = banner.SelectSingleNode("./Banner/Item/Text");
                if (item == null)
                {
                    item = xmlDocument.CreateElement("Item");
                    banner.AppendChild(item);
                }

                var title = item.SelectSingleNode("./Text");
                if (title == null)
                {
                    title = xmlDocument.CreateElement("Text");
                    item.AppendChild(title);
                }
                return title.InnerText;
            }
            set
            {
                var banner = xmlDocument.DocumentElement.SelectSingleNode("./Banner");
                if (banner == null)
                {
                    banner = xmlDocument.CreateElement("Banner");
                    xmlDocument.DocumentElement.AppendChild(banner);
                }

                var item = banner.SelectSingleNode("./Banner/Item/Text");
                if (item == null)
                {
                    item = xmlDocument.CreateElement("Item");
                    banner.AppendChild(item);
                }

                var title = item.SelectSingleNode("./Text");
                if (title == null)
                {
                    title = xmlDocument.CreateElement("Text");
                    item.AppendChild(title);
                }
                title.InnerText = value;
            }
        }

        public void AddBuild(Build build)
        {
            var existingBuild = xmlDocument.DocumentElement?.SelectSingleNode("./Builds/Build[@ID='" + build.ID + "']");
            if (existingBuild != null) throw(new Exception("Build Already Exists"));

            var parentNode = xmlDocument.DocumentElement?.SelectSingleNode("./Builds");

            var newBuild = xmlDocument.CreateElement("Build");
            SetXmlAttribute(newBuild, "ID", build.ID);
            SetXmlAttribute(newBuild, "DisplayName", build.DisplayName);
            SetXmlAttribute(newBuild, "Location", build.Location);

            var filters = "";
            foreach (var filter in build.Filters)
            {
                if (!string.IsNullOrEmpty(filters))
                {
                    filters += ",";
                }

                filters += filter;
            }

            var languages = "";
            foreach (var language in build.Languages)
            {
                if (!string.IsNullOrEmpty(languages))
                {
                    languages += ",";
                }

                languages += language.ID;
            }


            SetXmlAttribute(newBuild, "Filters", filters);
            SetXmlAttribute(newBuild, "Languages", languages);

            parentNode?.AppendChild(newBuild);
        }

        public void RemoveBuild(Build build)
        {
            var existingBuild = xmlDocument.DocumentElement?.SelectSingleNode("./Builds/Build[@ID='" + build.ID + "']");
            existingBuild?.ParentNode?.RemoveChild(existingBuild);
        }

        public List<Build> Builds
        {
            get
            {
                var returnBuilds = new List<Build>();
                var buildNodes = xmlDocument.DocumentElement?.SelectNodes("./Builds/Build");
                if (buildNodes != null)
                {
                    foreach (XmlNode buildNode in buildNodes)
                    {
                        var buildId = GetXmlAttribute(buildNode, "ID");
                        var displayName = GetXmlAttribute(buildNode, "DisplayName");
                        var location = GetXmlAttribute(buildNode, "Location");
                        var xmlLanguages = GetXmlAttribute(buildNode, "Languages");
                        var xmlFilters = GetXmlAttribute(buildNode, "Filters");

                        var languages = new List<Language>();
                        foreach (var xmlLanguage in xmlLanguages.Split(','))
                        {
                            if (!string.IsNullOrEmpty(xmlLanguage))
                            {
                                languages.Add(new Language()
                                {
                                    ID = xmlLanguage
                                });
                            }
                        }

                        var filters = new List<string>();
                        foreach (var filter in xmlFilters.Split(','))
                        {
                            if (!string.IsNullOrEmpty(filter))
                            {
                                filters.Add(filter);
                            }
                        }

                        var newBuild = new Build()
                        {
                            ID = buildId,
                            DisplayName = displayName,
                            Location = location,
                            Languages = languages,
                            Filters = filters
                        };

                        returnBuilds.Add(newBuild);
                    }
                }
                return returnBuilds;
            }
        }

        private string GetXmlAttribute(XmlNode xmlNode, string name)
        {
            var attr = xmlNode.Attributes[name];
            if (attr != null)
            {
                return attr.Value.ToString();
            }
            return "";
        }

        private void SetXmlAttribute(XmlNode xmlNode, string name, string value)
        {
            var attr = xmlNode.Attributes[name];
            if (attr == null)
            {
                attr = xmlNode.OwnerDocument.CreateAttribute(name);
                xmlNode.Attributes.Append(attr);
            }
            attr.Value = value;
        }

    }

    public class Build
    {
        
        public string ID { get; set; }

        public string DisplayName { get; set; }

        public string Location { get; set; }

        public List<string> Filters { get; set; }

        public List<Language> Languages { get; set; }
    }

    public class Language
    {
        public string ID { get; set; }
    }



}
