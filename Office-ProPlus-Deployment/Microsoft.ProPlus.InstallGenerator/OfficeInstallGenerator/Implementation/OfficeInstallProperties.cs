using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Micorosft.OfficeProPlus.ConfigurationXml;
using OfficeInstallGenerator;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeInstallProperties : IOfficeInstallProperties
    {
        public string ProductName { get; set; }

        public string ProgramFilesPath { get; set; }

        public string ProductId { get; set; }

        public Version Version { get; set; }

        public string UpgradeCode { get; set; }

        public string ExecutablePath { get; set; }

        public string Language { get; set; }

        public OfficeClientEdition OfficeClientEdition { get; set; }

        public OfficeVersion OfficeVersion { get; set; }

        public string BuildVersion { get; set; }

        public string ConfigurationXmlPath { get; set; }

        public string SourceFilePath { get; set; }

         

    }
}
