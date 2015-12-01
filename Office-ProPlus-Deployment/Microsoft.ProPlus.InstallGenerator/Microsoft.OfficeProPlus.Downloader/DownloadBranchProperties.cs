using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader.Model;

namespace Microsoft.OfficeProPlus.Downloader
{
    public class DownloadBranchProperties
    {
        public string BranchName { get; set; }

        public OfficeEdition OfficeEdition { get; set; }

        public List<string> Languages { get; set; }

        public string Version { get; set; }

        public string TargetDirectory { get; set; }

    }
}
