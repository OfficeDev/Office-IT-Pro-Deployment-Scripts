using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGenerator.Model
{
    public class OfficeRemoteInstall
    {

        public bool Installed { get; set; }

        public string Version { get; set; }

        public string LatestVersion { get; set; }

        public string Channel { get; set; }

        public bool LatestVersionInstalled
        {
            get
            {
                return Version == LatestVersion;
            }
        }
    }
}
