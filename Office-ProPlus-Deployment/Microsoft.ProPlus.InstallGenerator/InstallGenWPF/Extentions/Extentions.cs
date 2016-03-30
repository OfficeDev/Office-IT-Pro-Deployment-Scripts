using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Extentions
{
    public static class Extentions
    {
        public static string ConvertChannelToShortName(this string channel)
        {
            var channelName = channel.ToLower().Replace(" ", "");
            switch (channelName)
            {
                case "current":
                    return "CC";
                case "deferred":
                    return "DC";
                case "firstreleasedeferred":
                    return "FRDC";
                case "firstreleasecurrent":
                    return "FRCC";
                case "firstreleasebusiness":
                    return "FRDC";
                case "business":
                    return "DC";
            }
            return channel;
        }

    }
}
