using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Microsoft.OfficeProPlus.InstallGenerator.Models;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Models
{
    public class OfficeBranch
    {
        public Branch Branch { get; set; }

        public string Name { get; set; }

        public string Id { get; set; }

        public string CurrentVersion { get; set; }

        public List<Build> Versions { get; set; }

        public bool Updated { get; set; }

        public string NewName { get; set; }

    }
}
