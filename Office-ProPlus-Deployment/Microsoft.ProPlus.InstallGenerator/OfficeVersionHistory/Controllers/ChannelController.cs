using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Cors;
using Microsoft.OfficeProPlus.Downloader.Model;
using OfficeVersionHistory.Business;
using OfficeVersionHistory.CustomAttributes;

namespace OfficeVersionHistory.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class ChannelController : ApiController
    {
        private readonly VersionDownloader _versionDownloader = null;

        public ChannelController()
        {
            _versionDownloader = new VersionDownloader();
        }

        // GET api/Channel
        [HttpHeaderAttribute("Access-Control-Allow-Origin", "*")]
        public async Task<List<UpdateChannel>> Get()
        {
            var updateChannels = await _versionDownloader.GetUpdateChannelsAsync();
            return updateChannels;
        }

        // GET api/Channel/GetChannel
        [HttpHeaderAttribute("Access-Control-Allow-Origin", "*")]
        [Route("api/Channel/GetChannel")]
        public async Task<UpdateChannel> GetChannel(string name)
        {
            var updateChannels = await _versionDownloader.GetUpdateChannelsAsync();
            var selectChannel = updateChannels.FirstOrDefault(c => c.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
            return selectChannel;
        }

    }
}
