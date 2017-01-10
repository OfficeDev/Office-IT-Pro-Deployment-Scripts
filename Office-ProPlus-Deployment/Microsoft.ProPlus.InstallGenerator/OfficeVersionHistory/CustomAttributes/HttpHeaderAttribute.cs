using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http.Filters;

namespace OfficeVersionHistory.CustomAttributes
{
    public class HttpHeaderAttribute : ActionFilterAttribute
    {
        /// <summary>
        /// The name of the Http Header
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// The value of the Http Header
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// Set Http Headers in the response
        /// </summary>
        /// <param name="name">Name of the Http header</param>
        /// <param name="value">Value of the Http header</param>
        public HttpHeaderAttribute(string name, string value)
        {
            Name = name;
            Value = value;
        }

        public override void OnActionExecuted(HttpActionExecutedContext actionExecutedContext)
        {
            actionExecutedContext.Response.Headers.Add(Name, Value);
            base.OnActionExecuted(actionExecutedContext);
        }
    }
}