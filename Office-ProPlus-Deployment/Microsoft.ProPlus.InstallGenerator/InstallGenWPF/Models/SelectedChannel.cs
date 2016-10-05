using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Models
{
    public class SelectedChannel
    {
        public OfficeBranch Branch { get; set; }

        public BranchVersion SelectedVersion { get; set; }
    }
}
