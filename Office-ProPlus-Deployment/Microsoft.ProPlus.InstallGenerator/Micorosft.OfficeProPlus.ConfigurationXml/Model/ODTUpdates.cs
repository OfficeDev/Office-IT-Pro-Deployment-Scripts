using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Micorosft.OfficeProPlus.ConfigurationXml.Model
{
    public class ODTUpdates
    {
        public bool Enabled { get; set; }

        public string UpdatePath { get; set; }

        public Version TargetVersion { get; set; }

        public DateTime? Deadline { get; set; }

        public Branch? Branch { get; set; }

        public ODTChannel? ODTChannel { get; set; }

    }
}
