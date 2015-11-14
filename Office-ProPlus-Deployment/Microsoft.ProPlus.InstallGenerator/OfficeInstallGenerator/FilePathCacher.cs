using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeInstallGenerator
{
    public class FilePathCacher
    {
        private XmlDocument _xmlDoc = null;
        private string _xmlFilePath = null;

        public FilePathCacher(string xmlFilePath)
        {
            _xmlDoc = new XmlDocument();
            var mainNode = _xmlDoc.CreateElement("Files");
            _xmlDoc.AppendChild(mainNode);
            _xmlDoc.Save(xmlFilePath);
            _xmlFilePath = xmlFilePath;
        }

        public void AddFile(string rootPath, string filePath)
        {
            var fileSplit = filePath.Split('\\');
            var fileName = fileSplit[fileSplit.Length - 1];

            var fileNode = _xmlDoc.SelectSingleNode("/Files/File[@Path='" + filePath + "']");
            if (fileNode == null)
            {
                var saveFilePath = Regex.Replace(filePath, "^" + rootPath.Replace(@"\", @"\\") + @"\\", "");
                var folderPath = Regex.Replace(saveFilePath, @"\\" + fileName + "$", "");

                fileNode = _xmlDoc.CreateElement("File");
                SetAttribute(fileNode, "Path", saveFilePath);
                SetAttribute(fileNode, "FileName", fileName);
                SetAttribute(fileNode, "FolderPath", folderPath);
                _xmlDoc.DocumentElement.AppendChild(fileNode);
            }

            var md5Hash = GenerateMD5Hash(filePath);
            SetAttribute(fileNode, "Hash", md5Hash);

            _xmlDoc.Save(_xmlFilePath);
        }

        private string GenerateMD5Hash(string filePath)
        {
            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(filePath))
                {
                    return BitConverter.ToString(md5.ComputeHash(stream)).Replace("-", "").ToLower();
                }
            }
        }

        private void SetAttribute(XmlNode xmlNode, string name, string value)
        {
            var pathAttr = xmlNode.Attributes[name];
            if (pathAttr == null)
            {
                pathAttr = _xmlDoc.CreateAttribute(name);
                xmlNode.Attributes.Append(pathAttr);
            }
            pathAttr.Value = value;
        }


    }
}
