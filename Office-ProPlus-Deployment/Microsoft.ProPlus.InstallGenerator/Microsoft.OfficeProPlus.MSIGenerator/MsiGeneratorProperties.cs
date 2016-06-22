using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.OfficeProPlus.MSIGen;

namespace Microsoft.OfficeProPlus.InstallGenerator
{
    public class MsiGeneratorProperties
    {

        public string Name { get; set; }

        public string Language { get; set; }

        public string ExecutablePath { get; set; }

        public string MsiPath { get; set; }

        public string ProgramFilesPath { get; set; }

        public string Manufacturer { get; set; }

        public List<string> ProgramFiles { get; set; }

        public string SourceFilePath { get; set; }

        public Guid ProductId { get; set; }

        public Guid UpgradeCode { get; set; }

        public Version Version { get; set; }

        public string WixToolsPath { get; set; }

        public MsiDirectory MsiProgramFiles { get; set; }

    }
}


