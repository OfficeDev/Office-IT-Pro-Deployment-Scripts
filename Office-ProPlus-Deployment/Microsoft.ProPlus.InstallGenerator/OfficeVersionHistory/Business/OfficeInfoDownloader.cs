using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;

namespace OfficeVersionHistory.Business
{
    public class OfficeInfoDownloader
    {
        private readonly ProPlusDownloader _proPlusDownloader = null;

        public OfficeInfoDownloader()
        {
            _proPlusDownloader = new ProPlusDownloader();
        }

        public async Task<List<UpdateChannel>> GetUpdateChannelsAsync()
        {
            List<UpdateChannel> updateChannels = null;

            var now = DateTime.Now;
            var dateKey = "Channels-" + now.Year + now.Month + now.Day + now.Hour;

            if (WebApiConfig.ChannelCache.ContainsKey(dateKey))
            {
                updateChannels = WebApiConfig.ChannelCache[dateKey];
            }
            else
            {
                updateChannels = await _proPlusDownloader.DownloadReleaseHistoryCabAsync();
                WebApiConfig.ChannelCache[dateKey] = updateChannels;
            }

            return updateChannels;
        }

        public async Task<List<UpdateFiles>> GetUpdateFilesAsync()
        {
            List<UpdateFiles> updateFiles = null;

            var now = DateTime.Now;
            var dateKey = "Files-" + now.Year + now.Month + now.Day + now.Hour;

            if (WebApiConfig.FileCache.ContainsKey(dateKey))
            {
                updateFiles = WebApiConfig.FileCache[dateKey];
            }
            else
            {
                updateFiles = await _proPlusDownloader.DownloadCabAsync();
                WebApiConfig.FileCache[dateKey] = updateFiles;
            }

            return updateFiles;
        }


    }
}