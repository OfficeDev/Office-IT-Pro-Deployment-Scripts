using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using Owin;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;
using System.Configuration;
using System.Globalization;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Threading.Tasks;
//using WebAppGraphAPI.Controllers;
using SelfService.Utils;
using System.Security.Cryptography.X509Certificates;

namespace SelfService
{
    public partial class Startup
    {
        //
        // The Client ID is used by the application to uniquely identify itself to Azure AD.
        // The App Key is a credential used to authenticate the application to Azure AD.  Azure AD supports password and certificate credentials.
        // The Metadata Address is used by the application to retrieve the signing keys used by Azure AD.
        // The AAD Instance is the instance of Azure, for example public Azure or Azure China.
        // The Authority is the sign-in URL of the tenant.
        // The Post Logout Redirect Uri is the URL where the user will be redirected after they sign out.
        //
        private static string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static string appKey = ConfigurationManager.AppSettings["ida:AppKey"];
        private static string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private static string tenant = ConfigurationManager.AppSettings["ida:Tenant"];
        private static string postLogoutRedirectUri = ConfigurationManager.AppSettings["ida:PostLogoutRedirectUri"];
        private static string certName = ConfigurationManager.AppSettings["ida:CertName"];

        public static readonly string Authority = String.Format(CultureInfo.InvariantCulture, aadInstance, tenant);

        // This is the resource ID of the AAD Graph API.  We'll need this to request a token to call the Graph API.
        string graphResourceId = ConfigurationManager.AppSettings["ida:GraphUrl"];

        public void ConfigureAuth(IAppBuilder app)
        {
            app.SetDefaultSignInAsAuthenticationType(CookieAuthenticationDefaults.AuthenticationType);

            app.UseCookieAuthentication(new CookieAuthenticationOptions());

            app.UseOpenIdConnectAuthentication(
                new OpenIdConnectAuthenticationOptions
                {
                    
                    ClientId = clientId,
                    Authority = Authority,
                    PostLogoutRedirectUri = postLogoutRedirectUri,

                    Notifications = new OpenIdConnectAuthenticationNotifications()
                    {
                        //
                        // If there is a code in the OpenID Connect response, redeem it for an access token and refresh token, and store those away.
                        //
                        
                        AuthorizationCodeReceived = (context) =>
                        {
                            var code = context.Code;

                            if ( certName.Length != 0)
                            {
                                // Create a Client Credential Using a Certificate
                                //
                                // Initialize the Certificate Credential to be used by ADAL.
                                // First find the matching certificate in the cert store.
                                //

                                X509Certificate2 cert = null;
                                X509Store store = new X509Store(StoreLocation.CurrentUser);
                                try
                                {
                                    store.Open(OpenFlags.ReadOnly);
                                    // Place all certificates in an X509Certificate2Collection object.
                                    X509Certificate2Collection certCollection = store.Certificates;
                                    // Find unexpired certificates.
                                    X509Certificate2Collection currentCerts = certCollection.Find(X509FindType.FindByTimeValid, DateTime.Now, false);
                                    // From the collection of unexpired certificates, find the ones with the correct name.
                                    X509Certificate2Collection signingCert = currentCerts.Find(X509FindType.FindBySubjectDistinguishedName, certName, false);
                                    if (signingCert.Count == 0)
                                    {
                                        // No matching certificate found.
                                        return Task.FromResult(0);
                                    }
                                    // Return the first certificate in the collection, has the right name and is current.
                                    cert = signingCert[0];
                                }
                                finally
                                {
                                    store.Close();
                                }

                                // Then create the certificate credential.
                                ClientAssertionCertificate credential = new ClientAssertionCertificate(clientId, cert);

                                string userObjectID = context.AuthenticationTicket.Identity.FindFirst(
                                    "http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
                                AuthenticationContext authContext = new AuthenticationContext(Authority, new NaiveSessionCache(userObjectID));
                                AuthenticationResult result = authContext.AcquireTokenByAuthorizationCode(
                                    code, new Uri(HttpContext.Current.Request.Url.GetLeftPart(UriPartial.Path)), credential, graphResourceId);
                                AuthenticationHelper.token = result.AccessToken;
                            }
                            else
                            {
                                // Create a Client Credential Using an Application Key
                                ClientCredential credential = new ClientCredential(clientId, appKey);
                                string userObjectID = context.AuthenticationTicket.Identity.FindFirst(
                                    "http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
                                AuthenticationContext authContext = new AuthenticationContext(Authority, new NaiveSessionCache(userObjectID));
                                AuthenticationResult result = authContext.AcquireTokenByAuthorizationCode(
                                    code, new Uri(HttpContext.Current.Request.Url.GetLeftPart(UriPartial.Path)), credential, graphResourceId);
                                AuthenticationHelper.token = result.AccessToken;
                            }

                            return Task.FromResult(0);
                        }

                    }

                });
        }
    }
}
