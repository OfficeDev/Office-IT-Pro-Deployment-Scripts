using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGenerator.Extensions
{
    public class DirectoryHelper
    {

        public static string GetCurrentDirectoryFilePath(string fileName)
        {
            var projectfilePath = System.IO.Directory.GetCurrentDirectory() + @"\Project\" + fileName;
            var filePath = System.IO.Directory.GetCurrentDirectory() + @"\" + fileName;

            if (File.Exists(projectfilePath))
            {
                return projectfilePath;
            }

            return filePath;
        }

    }
}
