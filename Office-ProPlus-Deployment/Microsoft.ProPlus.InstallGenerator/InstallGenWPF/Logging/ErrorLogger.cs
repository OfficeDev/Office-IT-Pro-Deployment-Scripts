using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Logging
{
    public static class ErrorLogger
    {

        public static void LogException(this Exception ex, bool showError = true)
        {
            var tmpDir = Environment.ExpandEnvironmentVariables("%temp%");
            var logDir = tmpDir + @"\OfficeProPlusInstallGeneratorLogs";
            Directory.CreateDirectory(logDir);

            var now = DateTime.Now;
            var logFile = "OPPInstallGen-" + now.Year +
                          ConvertDate(now.Month.ToString()) +
                          ConvertDate(now.Day.ToString()) +
                          ConvertDate(now.Hour.ToString()) +
                          ConvertDate(now.Minute.ToString()) +
                          ConvertDate(now.Second.ToString()) + ".log";

            using (var sw = new StreamWriter(logDir + @"\" + logFile))
            {
                sw.WriteLine(ex.ToString());
                sw.WriteLine();
                sw.Flush();
                sw.Close();
            }

            if (showError)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private static string ConvertDate(string datePart)
        {
            if (datePart.Length == 1)
            {
                datePart = "0" + datePart;
            }
            return datePart;
        }

    }
}
