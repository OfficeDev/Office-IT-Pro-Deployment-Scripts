using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Microsoft.OfficeProPlus.Downloader
{
    public static class Extensions
    {

        public static string GetLanguageName(this int lcid)
        {
            var machineCulture = CultureInfo.GetCultures(CultureTypes.AllCultures).FirstOrDefault(c => c.LCID == lcid);
            return machineCulture != null ? machineCulture.IetfLanguageTag : null;
        }

        public static int GetLanguageNumber(this string ietfLanguageTag)
        {
            var machineCulture = CultureInfo.GetCultures(CultureTypes.AllCultures)
                                .FirstOrDefault(c =>  c.IetfLanguageTag.ToLower() == ietfLanguageTag.ToLower());
            return machineCulture != null ? machineCulture.LCID : 0;
        }

        public static string GetAttributeValue(this XmlNode node, string attribute)
        {
            if (node.Attributes[attribute] == null) return null;
            var value = node.Attributes[attribute].Value;
            return value;
        }

        public static bool IsActive(this Task task)
        {
            if (task == null) return false;
            if (task.IsCanceled || task.IsCompleted || task.IsFaulted)
            {
                return false;
            }

            if (task.Status == TaskStatus.Canceled || task.Status == TaskStatus.Faulted ||
                task.Status == TaskStatus.RanToCompletion)
            {
                return false;
            }

            return true;
        }

    }
}
