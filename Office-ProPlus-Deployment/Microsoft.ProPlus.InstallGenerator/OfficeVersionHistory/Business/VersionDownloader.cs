using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Web;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;

namespace OfficeVersionHistory.Business
{
    public class VersionDownloader
    {
        private readonly ProPlusDownloader _proPlusDownloader = null;

        public VersionDownloader()
        {
            _proPlusDownloader = new ProPlusDownloader();
        }

        public async Task<List<UpdateChannel>> GetUpdateChannelsAsync()
        {
            List<UpdateChannel> updateChannels = null;

            var now = DateTime.Now;
            var dateKey = "ChannelsInfo-" + now.Year + now.Month + now.Day + now.Hour;

            if (WebApiConfig.ChannelCache.ContainsKey(dateKey))
            {
                updateChannels = WebApiConfig.ChannelCache[dateKey];
            }
            else
            {
                updateChannels = await _proPlusDownloader.DownloadVersionsFromWebSite();
                if (!(updateChannels != null && updateChannels.Count > 0))
                {
                    updateChannels = await _proPlusDownloader.DownloadReleaseHistoryCabAsync();
                }
                WebApiConfig.ChannelCache[dateKey] = updateChannels;
            }

            return updateChannels;
        }




    }
}