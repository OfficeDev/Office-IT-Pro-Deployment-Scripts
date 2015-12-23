using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.Downloader
{
    public class FileDownloader
    {

        public async Task DownloadAsync(string url, string filePath, CancellationToken token = new CancellationToken())
        {
            var fSplit = filePath.Split('\\');
            var fileName = fSplit[fSplit.Length - 1];

            var directory = Regex.Replace(filePath, @"\\" + fileName + "$", "");
            Directory.CreateDirectory(directory);

            await Task.Run(async () =>
            {
                using (var client = new WebClient())
                {
                    client.DownloadProgressChanged +=
                        new DownloadProgressChangedEventHandler(client_DownloadProgressChanged);
                    client.DownloadFileCompleted += new AsyncCompletedEventHandler(client_DownloadFileCompleted);
                   

                    if (!token.IsCancellationRequested)
                    {
                        // Register the callback to a method that can unblock.
                        using (var ctr = token.Register(() => client.CancelAsync()))
                        {
                            await client.DownloadFileTaskAsync(new Uri(url), filePath);
                        }
                    }
                }
            }, token);
        }

        public async Task<long> GetFileSizeAsync(string url)
        {
            var request = WebRequest.Create(url);
            using (var response = await request.GetResponseAsync())
            {
                try
                {
                    var webResponse = (HttpWebResponse) response;
                    if (webResponse.StatusCode != HttpStatusCode.OK)
                        throw (new Exception(webResponse.StatusDescription));
                    return webResponse.ContentLength;
                }
                finally
                {
                    response.Close();
                }
            }
        }

        private void client_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            var bytesIn = double.Parse(e.BytesReceived.ToString());
            var totalBytes = double.Parse(e.TotalBytesToReceive.ToString());
            var percentage = bytesIn / totalBytes * 100;

            if (DownloadFileProgress != null)
            {
                DownloadFileProgress(this, new Events.DownloadFileProgress()
                {
                    PercentageComplete =  Math.Truncate(percentage),
                    BytesRecieved = e.BytesReceived,
                    TotalBytesToRecieve = e.TotalBytesToReceive
                });
            }
        }

        private void client_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            if (DownloadFileComplete != null)
            {
                DownloadFileComplete(this, new EventArgs());
            }
        }

        public Events.DownloadFileCompleteEventHandler DownloadFileComplete { get; set; }

        public Events.DownloadFileProgressEventHandler DownloadFileProgress { get; set; }


    }
}
