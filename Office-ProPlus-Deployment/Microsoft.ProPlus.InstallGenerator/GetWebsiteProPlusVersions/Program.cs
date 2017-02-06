using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader;

namespace GetWebsiteProPlusVersions
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var p = new Program();
                p.RunProgram().Wait();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
            }
        }

        private async Task RunProgram()
        {
            var p = new ProPlusDownloader();
            await p.DownloadVersionsFromWebSite();
        }


       
    }
}
