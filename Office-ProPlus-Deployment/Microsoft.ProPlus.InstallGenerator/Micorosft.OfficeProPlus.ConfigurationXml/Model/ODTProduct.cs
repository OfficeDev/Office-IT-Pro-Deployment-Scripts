using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Micorosft.OfficeProPlus.ConfigurationXml.Model
{
    public class ODTProduct
    {

        public string ID { get; set; }

        public List<ODTLanguage> Languages { get; set; }

        public List<ODTExcludedApp> ExcludeApps { get; set; } 

    }
}
