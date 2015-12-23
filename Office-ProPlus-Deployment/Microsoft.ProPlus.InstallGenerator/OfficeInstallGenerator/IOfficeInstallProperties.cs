using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeInstallGenerator;

namespace Microsoft.OfficeProPlus.InstallGenerator
{
    public interface IOfficeInstallProperties
    {
        OfficeVersion OfficeVersion { get; set; }

        string ConfigurationXmlPath { get; set; }

        string SourceFilePath { get; set; }

        string ExecutablePath { get; set; }
    }
}
