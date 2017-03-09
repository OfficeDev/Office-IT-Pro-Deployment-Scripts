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
    public class LanguageController : ApiController
    {
        private readonly OfficeInfoDownloader _officeInfoDownloader = null;

        public LanguageController()
        {
            _officeInfoDownloader = new OfficeInfoDownloader();
        }

        // GET api/Languages
        [HttpHeader("Access-Control-Allow-Origin", "*")]
        public async Task<List<Language>> Get()
        {
            var updateFiles = await _officeInfoDownloader.GetUpdateFilesAsync();
            return updateFiles.FirstOrDefault()?.Languages;
        }
    }
}
