using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.Downloader.Model
{
    public class UpdateChannel
    {
        public UpdateChannel()
        {
            if (Updates == null) Updates = new List<Update>();
        }

        public string Name { get; set; }

        public List<Update> Updates { get; set; }
    }
}
