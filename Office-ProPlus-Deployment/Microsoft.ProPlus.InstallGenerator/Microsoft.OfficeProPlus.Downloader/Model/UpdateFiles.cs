using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.Downloader.Model
{
    public class UpdateFiles
    {
        public UpdateFiles()
        {
            BaseURL = new List<baseURL>();
            Files = new List<File>();
        }

        public List<baseURL> BaseURL { get; set; }

        public List<File> Files { get; set; }

        public OfficeEdition OfficeEdition { get; set; }
    }
}
