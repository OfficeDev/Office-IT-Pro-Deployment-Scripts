using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation
{
    public class AppSettings
    {
        private const string MGenerateScriptUrl = "https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/raw/master/Office-ProPlus-Deployment/Generate-ODTConfigurationXML/Generate-ODTConfigurationXML.ps1";

        public static string GenerateScriptUrl
        {
            get
            {
                var objConfig = GetAppSetting<string>("AzureStorageAccountType").ToString();
                return String.IsNullOrEmpty(objConfig) ? MGenerateScriptUrl : objConfig;
            }
        }

        private static dynamic GetAppSetting<T>(string name)
        {
            var objConfig = ConfigurationManager.AppSettings[name] ?? "";

            if (typeof(T) == typeof(string))
            {
                return objConfig;
            }
            if (typeof(T) == typeof(int))
            {
                return Convert.ToInt32(objConfig);
            }
            return objConfig;
        } 

    }
}
