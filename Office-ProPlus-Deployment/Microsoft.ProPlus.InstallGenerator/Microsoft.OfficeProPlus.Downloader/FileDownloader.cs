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
using System.Configuration;

namespace Microsoft.OfficeProPlus.Downloader
{
    public class FileDownloader
    {

        public async Task DownloadAsync(string url, string filePath, CancellationToken token = new CancellationToken())
        {
            var fSplit = filePath.Split('\\');
            var fileName = fSplit[fSplit.Length - 1];

            var numAttempts = 0;
            var downloadSuccessful = false; //variables for redownload attempts to retry, or kick out of loop if necessary

            var numAllowedRetries = Convert.ToInt32(ConfigurationSettings.AppSettings["NumDownloadRetries"]);
            while (numAttempts <= numAllowedRetries && !downloadSuccessful)//loop for checking number of attempts and if attempt was a success
            {
                try
                {
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

                                using (var ctr = token.Register(() => client.CancelAsync()))
                                {
                                    //actual download, will retry if fails                            
                                    await client.DownloadFileTaskAsync(new Uri(url), filePath);
                                    downloadSuccessful = true;                                      //flag as downloaded to kick out of loop
                                    //end of file download                        
                                }

                                //// Register the callback to a method that can unblock.                        
                                //using (var ctr = token.Register(() => client.CancelAsync()))
                                //using (var file = File.Create(filePath))
                                //using(Stream stream = await client.OpenReadTaskAsync(new Uri(url)))
                                //{
                                //    var buffer = new byte[8192];
                                //    int bytesReceived; 
                                //    //actual download, will retry if fails                   
                                //    while ((bytesReceived = await stream.ReadAsync(buffer,0,buffer.Length,token)) != 0)
                                //    {
                                //        file.Write(buffer,0,bytesReceived);
                                //    }

                                //    stream.Close();
                                //    downloadSuccessful = true;  //flag as downloaded to kick out of loop

                                //    //end of file download                        
                                //}


                            }
                        }
                    }, token);
                    return;
                }
                catch (Exception ex)
                {
                    numAttempts++;
                    if (ex.Message.Contains("The request was aborted"))//If user aborts, will kick out without attempting re-download, also prevents app for displaying "download complete" if user clicks stop
                    {
                        throw ex;
                    }
                    else if (numAttempts >= numAllowedRetries)
                    {
                        
                        throw ex;// on final attempt, throw an error.
                    }
                }
                await Task.Delay(new TimeSpan(0, 0, 3), token);
            }
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
