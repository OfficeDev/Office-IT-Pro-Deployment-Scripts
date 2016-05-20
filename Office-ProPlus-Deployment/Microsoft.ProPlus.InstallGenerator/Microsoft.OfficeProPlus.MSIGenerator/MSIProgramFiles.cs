using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.MSIGen
{

    public class MsiDirectory
    {
        public MsiDirectory()
        {
            MsiDirectories = new List<MsiDirectory>();
            MsiFiles = new List<MsiFile>();
        }

        public string Name { get; set; }

        public string RootPath { get; set; }

        public string RelativePath { get; set; }

        public List<MsiDirectory> MsiDirectories { get; set; }

        public List<MsiFile> MsiFiles { get; set; }
    }

    public class MsiFile
    {
        public string Path { get; set; }
    }


}
