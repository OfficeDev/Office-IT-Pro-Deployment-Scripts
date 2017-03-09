using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.Downloader.Model
{
    public class Update
    {
        public bool Latest { get; set; }

        public string Version { get; set; }

        public string LegacyVersion { get; set; }

        public string Build { get; set; }

        public DateTime PublishTime { get; set; }
    }
}
