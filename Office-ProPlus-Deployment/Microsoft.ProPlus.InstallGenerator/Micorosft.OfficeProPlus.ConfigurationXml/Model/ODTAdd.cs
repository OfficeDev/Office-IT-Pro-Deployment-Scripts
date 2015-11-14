using System;
using System.Collections.Generic;
using System.Linq;
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

        public string SourcePath { get; set; }

        public Version Version { get; set; }

        public List<ODTProduct> Products { get; set; }

    }
}
