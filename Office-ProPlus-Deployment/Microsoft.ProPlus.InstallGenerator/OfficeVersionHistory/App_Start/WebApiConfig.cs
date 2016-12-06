using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using Microsoft.OfficeProPlus.Downloader.Model;

namespace OfficeVersionHistory
{
    public static class WebApiConfig
    {
        public static ConcurrentDictionary<string, List<UpdateChannel>> ChannelCache = null;

        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services
            ChannelCache = new ConcurrentDictionary<string, List<UpdateChannel>>();

            config.EnableCors();

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
