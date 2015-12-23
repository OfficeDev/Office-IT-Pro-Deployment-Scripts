using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGenerator
{
    public interface IOfficeInstallReturn
    {

        string GeneratedFilePath { get; set; }

    }
}
