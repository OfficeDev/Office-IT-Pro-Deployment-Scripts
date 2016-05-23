using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Micorosft.OfficeProPlus.ConfigurationXml;
using OfficeInstallGenerator;

namespace Microsoft.OfficeProPlus.InstallGenerator
{
    public interface IOfficeInstallProperties
    {
        string ProductName { get; set; }

        string ProgramFilesPath { get; set; }

        string ProductId { get; set; }

        Version Version { get; set; }

        string UpgradeCode { get; set; }

        OfficeVersion OfficeVersion { get; set; }

        string BuildVersion { get; set; }

        string ConfigurationXmlPath { get; set; }

        string SourceFilePath { get; set; }

        string ExecutablePath { get; set; }

        string Language { get; set; }

        OfficeClientEdition OfficeClientEdition { get; set; }

    }
}
