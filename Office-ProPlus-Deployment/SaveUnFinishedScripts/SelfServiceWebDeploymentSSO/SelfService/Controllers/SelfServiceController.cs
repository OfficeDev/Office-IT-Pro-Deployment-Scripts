using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SelfService.Models;
using SelfService.Utils;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace SelfService.Controllers
{
    //[Authorize]
    public class SelfServiceController : Controller
    {

        private const string TenantIdClaimType = "http://schemas.microsoft.com/identity/claims/tenantid";
        private static readonly string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static readonly string appKey = ConfigurationManager.AppSettings["ida:AppKey"];
        private readonly string graphResourceId = ConfigurationManager.AppSettings["ida:GraphUrl"];

        private readonly string graphUserUrl = "https://graph.windows.net/{0}/me?api-version=" +
                                               ConfigurationManager.AppSettings["ida:GraphApiVersion"];

        //
        // GET: /SelfService/

        public async Task<ActionResult> Index()
        {
            if (!Request.IsAuthenticated) {
                HttpContext.GetOwinContext()
                    .Authentication.Challenge(new AuthenticationProperties { RedirectUri = "/SelfService" },
                        OpenIdConnectAuthenticationDefaults.AuthenticationType);
            }

            return View();
        }


        //
        // GET: /SelfService/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /SelfService/UserInfo

        public async Task<Object> UserInfo()
        {
            string tenantId = ClaimsPrincipal.Current.FindFirst(TenantIdClaimType).Value;
            AuthenticationResult result = null;

            // Get the access token from the cache
            string userObjectID =
                ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier")
                    .Value;
            AuthenticationContext authContext = new AuthenticationContext(Startup.Authority,
                new NaiveSessionCache(userObjectID));
            ClientCredential credential = new ClientCredential(clientId, appKey);
            result = authContext.AcquireTokenSilent(graphResourceId, credential,
                new UserIdentifier(userObjectID, UserIdentifierType.UniqueId));

            // Call the Graph API manually and retrieve the user's profile.
            string requestUrl = String.Format(
                CultureInfo.InvariantCulture,
                graphUserUrl,
                HttpUtility.UrlEncode(tenantId));
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
            HttpResponseMessage response = await client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                string responseString = await response.Content.ReadAsStringAsync();
                var profile = JsonConvert.DeserializeObject(responseString);
                return profile;
            }

            return 0;
        }

        //
        // GET: /SelfService/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /SelfService/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /SelfService/Edit/5

        public ActionResult Edit(int id)
        {
            return View();
        }

        //
        // POST: /SelfService/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /SelfService/Delete/5

        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /SelfService/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
