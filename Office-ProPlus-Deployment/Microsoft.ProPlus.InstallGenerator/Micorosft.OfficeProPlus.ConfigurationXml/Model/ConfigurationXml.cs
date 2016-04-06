using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;

namespace OfficeInstallGenerator.Model
{
    public class ConfigurationXml
    {

        public ODTAdd Add { get; set; }

        public ODTRemove Remove { get; set; }

        public ODTUpdates Updates { get; set; }

        public ODTDisplay Display { get; set; }

        public ODTLogging Logging { get; set; }

        public ODTLanguage Lanaguage { get; set; }

        public ODTProperties Properties { get; set; }

    }

    
}
