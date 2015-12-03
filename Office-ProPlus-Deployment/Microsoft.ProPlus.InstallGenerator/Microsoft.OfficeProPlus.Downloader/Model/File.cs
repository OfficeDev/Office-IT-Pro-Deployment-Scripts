using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.Downloader.Model
{
    public class File
    {
        public string Name { get; set; }

        public string Rename { get; set; }

        public string RelativePath { get; set; }

        public int Language { get; set; }

        public string RemoteUrl { get; set; }

        public long FileSize { get; set; }

        public string LocalFilePath { get; set; }

        public bool Exists
        {
            get
            {
                if (string.IsNullOrEmpty(LocalFilePath)) return false;
                if (!System.IO.File.Exists(LocalFilePath)) return false;

                var fInfo = new System.IO.FileInfo(LocalFilePath);
                return this.FileSize == fInfo.Length;
            }
        }

    }
}
