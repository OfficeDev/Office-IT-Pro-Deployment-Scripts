using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Models
{
    public class RemoteMachine
    {
        public bool include { get; set; }

        public string Machine { get; set; }

        public string UserName { get; set; }

        public string Password { get; set; }

        public string WorkGroup { get; set; }

        public string Status { get; set; }

        public Channel Channel { get; set; }

        public Channel OriginalChannel { get; set; }

        public List<Channel> Channels { get; set; }

        public officeVersion Version { get; set; }

        public officeVersion OriginalVersion { get; set; }

        public List<officeVersion> Versions { get; set; }

    }
}
