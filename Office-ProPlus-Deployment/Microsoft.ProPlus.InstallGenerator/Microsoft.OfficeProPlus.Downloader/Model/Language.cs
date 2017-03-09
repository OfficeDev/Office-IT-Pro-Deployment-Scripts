using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader.Model.Enums;

namespace Microsoft.OfficeProPlus.Downloader.Model
{
    public class Language
    {
        public string Name { get; set; }

        public string DisplayName { get; set; }

        public int LCID { get; set; }

        public LanguageType LanguageType { get; set; }

        public string IetfLanguageTag { get; set; }
    }
}
