using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

        public static bool IsValidVersion(this string version)
        {
            if (string.IsNullOrEmpty(version)) return false;
            var match = Regex.Match(version, @"^\d{2}\.\d\.\d{4}\.\d{4}$");
            return match.Success;
        }
    }
}
