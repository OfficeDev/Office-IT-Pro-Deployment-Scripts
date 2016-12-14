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
    public class FileController : ApiController
    {
        private readonly OfficeInfoDownloader _officeInfoDownloader = null;

        public FileController()
        {
            _officeInfoDownloader = new OfficeInfoDownloader();
        }

        // GET api/File
        [HttpHeader("Access-Control-Allow-Origin", "*")]
        public async Task<List<File>> Get(string officeEdition = "32")
        {
            var updateFiles = await _officeInfoDownloader.GetUpdateFilesAsync();
            if (officeEdition.Contains("64"))
            {
                var updateFile = updateFiles.FirstOrDefault(f => f.OfficeEdition == OfficeEdition.Office64Bit);
                return updateFile?.Files;
            }
            else if (officeEdition.Contains("32"))
            {
                var updateFile = updateFiles.FirstOrDefault(f => f.OfficeEdition == OfficeEdition.Office32Bit);
                return updateFile?.Files;
            }
            else
            {
                return new List<File>();
            }
        }
    }
}
