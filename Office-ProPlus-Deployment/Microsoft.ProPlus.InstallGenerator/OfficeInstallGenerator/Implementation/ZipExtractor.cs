using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class ZipExtractor
    {

        public static string Extract(string zipFile, string extractTo)
        {
            ZipFile.ExtractToDirectory(zipFile, extractTo);
            return extractTo;
        }

        public static string AssemblyDirectory
        {
            get
            {
                var codeBase = Assembly.GetExecutingAssembly().Location;
                var uri = new UriBuilder(codeBase);
                var path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

    }
}
