using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Micorosft.OfficeProPlus.ConfigurationXml.Enums;

namespace Micorosft.OfficeProPlus.ConfigurationXml.Model
{
    public class ODTLogging
    {
        public LoggingLevel Level { get; set; }

        public string Path { get; set; }

    }
}
