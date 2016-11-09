using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGenerator.Events
{
    public class Events
    {
        public delegate void UpdatingOfficeEventHandler(object sender, UpdatingOfficeArgs e);


        public class UpdatingOfficeArgs : EventArgs
        {
            public string Status { get; set; }

        }
    }
}
