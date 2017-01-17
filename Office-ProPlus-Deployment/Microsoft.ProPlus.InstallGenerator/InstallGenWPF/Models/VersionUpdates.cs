using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Models
{
    public class VersionUpdates
    {
        public bool Latest { get; set; }

        public string Version { get; set; }

        public string LegacyVersion { get; set; }

        public string Build { get; set; }

        public string PublishTime { get; set; }
    }
}
