using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;

namespace OfficeInstallGenerator.Model
{
    public class ODTAdd
    {
        public OfficeClientEdition OfficeClientEdition { get; set; }

        public Branch? Branch { get; set; }

        public ODTChannel? ODTChannel { get; set; }

        public string SourcePath { get; set; }

        public string DownloadPath { get; set; }

        public Version Version { get; set; }

        public List<ODTProduct> Products { get; set; }

        public bool? OfficeMgmtCOM { get; set; }

        public bool IsLanguagesSameForAllProducts()
        {
            return this.Products.All(productMain => !this.Products.Where(p => !String.Equals(p.ID, productMain.ID, StringComparison.CurrentCultureIgnoreCase))
                   .Any(productComp => productMain.Languages.Any(languageMain => productComp.Languages.All(l => l.ID.ToLower() != languageMain.ID.ToLower()))));
        }
    }
}
