using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MetroDemo.Models;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGenerator.Models;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Models
{
    public class SccmConfiguration
    {
        public List<SelectedChannel> Channels { get; set; }
        
        public List<Bitness>  Bitnesses { get; set; }

        public List<Language> Languages { get; set; }

        public List<Product> Products { get; set; }

        public string DeploymentDirectory { get; set; }

        public SccmScenario Scenario { get; set; }

        public DeploymentSource DeploymentSource { get; set; }
    }
}
