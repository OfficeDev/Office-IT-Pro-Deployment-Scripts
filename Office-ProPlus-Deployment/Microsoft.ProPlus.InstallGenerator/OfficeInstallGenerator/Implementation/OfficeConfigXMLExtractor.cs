using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeConfigXmlExtractor
    {

        public string ExtractXml(string filePath)
        {
            if (filePath.ToLower().EndsWith(".exe"))
            {
                return ExtractXmlFromExecutable(filePath);
            }
            if (filePath.ToLower().EndsWith(".msi"))
            {
                return ExtractXmlFromMsi(filePath);
            }
            if (filePath.ToLower().EndsWith(".xml"))
            {
                return File.ReadAllText(filePath);
            }
            throw (new Exception("Configuration XML Not Found"));
        }

        private string ExtractXmlFromMsi(string filePath)
        {
            var tmpDir = Environment.ExpandEnvironmentVariables("%temp%");
            var tmpPath = tmpDir + @"\" + Guid.NewGuid().ToString();
            Directory.CreateDirectory(tmpPath);

            try
            {
                var p = new Process
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = "msiexec",
                        Arguments = "/a " + filePath + " /q TARGETDIR=" + tmpPath,
                        CreateNoWindow = true,
                        UseShellExecute = false,
                    },
                };
                p.Start();
                p.WaitForExit();

                var dirInfo = new DirectoryInfo(tmpPath);
                var xmlConfig = dirInfo.GetFiles("*.xml", SearchOption.AllDirectories);

                if (!xmlConfig.Any())
                {
                    throw (new Exception("Configuration XML Not Found"));
                }

                var xml = File.ReadAllText(xmlConfig.FirstOrDefault().FullName);

                if (!xml.Contains("<Configuration"))
                {
                    throw (new Exception("Invalid Configuration XML"));
                }

                return xml;
            }
            finally
            {
                try
                {
                    Directory.Delete(tmpPath, true);
                }
                catch { }
            }
        }

        private string ExtractXmlFromExecutable(string fileName)
        {
            var tmpDir = Environment.ExpandEnvironmentVariables("%temp%");

            var p = new Process
            {
                StartInfo = new ProcessStartInfo()
                {
                    FileName = fileName,
                    Arguments = "/extractxml=" + tmpDir + @"\configuration.xml",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                },
            };
            p.Start();
            p.WaitForExit();

            var xml = File.ReadAllText(tmpDir + @"\configuration.xml");
            return xml;
        }

    }
}
