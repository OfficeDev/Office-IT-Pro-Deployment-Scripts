using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.Downloader
{
    public class Events
    {
        public delegate void DownloadFileCompleteEventHandler(object sender, EventArgs e);

        public delegate void DownloadFileProgressEventHandler(object sender, DownloadFileProgress e);

        public delegate void VersionDetectedEventHandler(object sender, BuildVersion e);

        public class DownloadFileProgress : EventArgs
        {
            public double PercentageComplete { get; set; }

            public long BytesRecieved { get; set; }

            public long TotalBytesToRecieve { get; set; }
        }

        public class BuildVersion : EventArgs
        {
            public string Version { get; set; }
        }
    }
}
