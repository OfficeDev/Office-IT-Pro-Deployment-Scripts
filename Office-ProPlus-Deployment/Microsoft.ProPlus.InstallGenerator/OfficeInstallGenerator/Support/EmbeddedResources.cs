using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OfficeInstallGenerator
{
    public class EmbeddedResources
    {

        public static List<string> GetEmbeddedItems(string targetDirectory, string nameSearch = null)
        {
            var returnFiles = new List<string>();
            var assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase);
            if (assemblyPath == null) return returnFiles;

            //var appRoot = new Uri(assemblyPath).LocalPath;
            var assembly = Assembly.GetExecutingAssembly();
            var assemblyName = assembly.GetName().Name;

            foreach (var resourceStreamName in assembly.GetManifestResourceNames())
            {
                using (var input = assembly.GetManifestResourceStream(resourceStreamName))
                {
                    var fileName = Regex.Replace(resourceStreamName, "^" + assemblyName + ".", "", RegexOptions.IgnoreCase);
                    fileName = Regex.Replace(fileName, "^Resources.", "", RegexOptions.IgnoreCase);

                    if (nameSearch != null)
                    {
                        var match = Regex.Match(fileName, nameSearch);
                        if (!match.Success) continue;
                    }

                    var nameSplit = fileName.Split('.');
                    fileName = nameSplit[nameSplit.Length - 2] + "." + nameSplit[nameSplit.Length - 1];

                    returnFiles.Add(fileName);

                    var filePath = Path.Combine(targetDirectory, fileName);
                    try
                    {
                        if (File.Exists(filePath)) File.Delete(filePath);

                        using (Stream output = File.Create(filePath))
                        {
                            CopyStream(input, output);
                        }
                    }
                    catch { }
                }
            }
            return returnFiles;
        }

        private static void CopyStream(Stream input, Stream output)
        {
            // Insert null checking here for production
            var buffer = new byte[8192];

            int bytesRead;
            while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, bytesRead);
            }
        }

        public static string ResourcePath
        {
            get
            {
                return Directory.GetCurrentDirectory();
            }
        }

    }
}
