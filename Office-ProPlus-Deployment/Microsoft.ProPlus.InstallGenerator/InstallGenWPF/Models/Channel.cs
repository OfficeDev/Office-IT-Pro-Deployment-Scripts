using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.InstallGenerator.Models;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Models
{
    public class Channel
    {
        public string Name { get; set; }

        public string ChannelName { get; set; }

        public string Version { get; set; }

        public string DisplayVersion { get; set; }

        public bool Selected { get; set; }

        public bool Editable { get; set; }

        public double PercentDownload { get; set; }

        public string PercentDownloadText { get; set; }

        public List<Build> Builds { get; set; }

        public string ForeGround { get; set; }

    }
}
