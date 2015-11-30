﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace Microsoft.OfficeProPlus.InstallGenerator.Extensions
{
    public static class Extensions
    {

        static public string Beautify(this XmlDocument doc)
        {
            var sb = new StringBuilder();
            var settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  ",
                NewLineChars = "\r\n",
                NewLineHandling = NewLineHandling.Replace,
               OmitXmlDeclaration = true
            };
            using (var writer = XmlWriter.Create(sb, settings))
            {
                doc.Save(writer);
            }

            var xml = sb.ToString();
            return xml;
        }

        static public string BeautifyXml(this string xml)
        {
            var doc = new XmlDocument();
            doc.LoadXml(xml);
            return doc.Beautify();
        }

    }
}
