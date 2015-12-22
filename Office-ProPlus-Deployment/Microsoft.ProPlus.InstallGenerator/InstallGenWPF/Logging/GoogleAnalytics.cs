using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Garlic;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Logging
{
    public class GoogleAnalytics
    {

        private static AnalyticsSession session = null;

        public static void Log(string path, string pageName)
        {
            if (session == null)
            {
                session = new AnalyticsSession("officedev.github.io/Office-IT-Pro-Deployment-Scripts/InstallGenerator",
                        "UA-70271323-2");
            }


            var page = session.CreatePageViewRequest(
                  path,       // path
                  pageName); // page title

            // or send page views manually
            page.Send();
        }

    }
}
