#region

using System.Configuration;

#endregion

namespace SelfService.Utils
{
    internal static class GraphConfiguration
    {
        private const string tenantIdClaimType = "http://schemas.microsoft.com/identity/claims/tenantid";
        private static string graphResourceId = ConfigurationManager.AppSettings["ida:GraphUrl"];
        private static string graphApiVersion = ConfigurationManager.AppSettings["ida:GraphApiVersion"];

        internal static string GraphApiVersion
        {
            get { return graphApiVersion; }
            set { graphApiVersion = value; }
        }

        internal static string GraphResourceId
        {
            get { return graphResourceId; }
            set { graphResourceId = value; }
        }

        internal static string TenantIdClaimType
        {
            get { return tenantIdClaimType; }
        }
    }
}