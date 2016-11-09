using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Micorosft.OfficeProPlus.ConfigurationXml.Enums;

namespace Micorosft.OfficeProPlus.ConfigurationXml.Model
{
    public class ODTProperties
    {
        public YesNo? AutoActivate { get; set; }

        public bool? ForceAppShutdown { get; set; }

        public bool? SharedComputerLicensing { get; set; }

        public bool? PinIconsToTaskbar { get; set; }

    }
}
