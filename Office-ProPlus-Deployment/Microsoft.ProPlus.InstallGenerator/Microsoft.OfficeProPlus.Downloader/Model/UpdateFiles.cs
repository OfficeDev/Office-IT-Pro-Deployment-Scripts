using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader.Model.Enums;

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

        public List<Language> Languages
        {
            get
            {
                var languages = new List<Language>();
                foreach (var file in Files)
                {
                    var langNum = file.Language;
                    if (langNum == 0) continue;
                    var cInfo = CultureInfo.GetCultureInfo(langNum);
                    if (languages.Any(l => l.LCID == cInfo.LCID)) continue;

                    languages.Add(new Language()
                    {
                        DisplayName = cInfo.DisplayName,
                        LCID = cInfo.LCID,
                        Name = cInfo.Name,
                        LanguageType = file.LanguageType,
                        IetfLanguageTag = cInfo.IetfLanguageTag
                    });
                }
                return languages;
            }
        }

        public OfficeEdition OfficeEdition { get; set; }
    }
}
