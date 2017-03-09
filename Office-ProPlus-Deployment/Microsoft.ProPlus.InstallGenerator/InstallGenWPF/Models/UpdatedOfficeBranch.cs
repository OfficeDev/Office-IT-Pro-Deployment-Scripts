using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Models
{
    public class UpdatedOfficeBranch
    {
        public string Name { get; set; }

        public List<VersionUpdates> Updates { get; set; }
    }
}
