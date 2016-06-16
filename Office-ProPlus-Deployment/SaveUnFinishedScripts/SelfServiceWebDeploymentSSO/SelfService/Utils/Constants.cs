#region

using System.Configuration;

#endregion

namespace SelfService.Utils
{
    internal class Constants
    {
        public static string ResourceUrl = ConfigurationManager.AppSettings["ida:GraphUrl"];
        public static string ClientId = ConfigurationManager.AppSettings["ida:ClientId"];
        public static string AppKey = ConfigurationManager.AppSettings["ida:AppKey"];
        public static string TenantId = ConfigurationManager.AppSettings["ida:TenantId"];
        public static string AuthString = ConfigurationManager.AppSettings["ida:Auth"] +
                                          ConfigurationManager.AppSettings["ida:Tenant"];
        public static string ClientSecret = ConfigurationManager.AppSettings["ida:ClientSecret"];
    }
}