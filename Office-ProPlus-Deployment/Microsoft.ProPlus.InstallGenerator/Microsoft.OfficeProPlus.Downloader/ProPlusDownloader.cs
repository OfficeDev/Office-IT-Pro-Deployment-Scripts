using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.OfficeProPlus.Downloader.Model;
using File = System.IO.File;

namespace Microsoft.OfficeProPlus.Downloader
{
    public class ProPlusDownloader
    {
        private const string OfficeVersionUrl = "http://officecdn.microsoft.com/pr/wsus/ofl.cab";

        private List<UpdateFiles> _updateFiles { get; set; }

        public async Task DownloadBranch(DownloadBranchProperties properties, CancellationToken token = new CancellationToken())
        {
            var fd = new FileDownloader();

            if (properties.Languages == null) properties.Languages = new List<string>() { "en-us" };

            if (_updateFiles == null)
            {
                _updateFiles = await DownloadCab();
            }

            var selectUpdateFile = properties.OfficeEdition == OfficeEdition.Office32Bit ? _updateFiles.FirstOrDefault(u => u.OfficeEdition == OfficeEdition.Office32Bit) :
                                                                _updateFiles.FirstOrDefault(u => u.OfficeEdition == OfficeEdition.Office64Bit);
            if (selectUpdateFile == null) throw (new Exception("Cannot Find Office Files"));

            var branch = selectUpdateFile.BaseURL.FirstOrDefault(b => b.Branch.ToLower() == properties.BranchName.ToLower());

            var version = properties.Version;
            if (properties.Version == null)
            {
                version = await GetLatestVersionAsync(branch, properties.OfficeEdition);
                if (VersionDetected != null)
                {
                    VersionDetected(this, new Events.BuildVersion()
                    {
                        Version = version
                    });
                }
            }

            var allFiles = new List<Model.File>();
            foreach (var language in properties.Languages)
            {
                var langCode = language.GetLanguageNumber();
                var langfiles = selectUpdateFile.Files.Where(f => f.Language == 0 || f.Language == langCode);

                foreach (var file in langfiles)
                {
                    file.Name = Regex.Replace(file.Name, "%version%", version, RegexOptions.IgnoreCase);
                    file.RelativePath = Regex.Replace(file.RelativePath, "%version%", version, RegexOptions.IgnoreCase);
                    file.RemoteUrl = branch.URL + @"/" + file.RelativePath + file.Name;
                    file.FileSize = await fd.GetFileSizeAsync(file.RemoteUrl);

                    allFiles.Add(file);

                    if (token.IsCancellationRequested)
                    {
                        return;
                    }
                }
            }

            allFiles = allFiles.Distinct().ToList();

            fd = new FileDownloader();

            foreach (var file in allFiles)
            {
                file.LocalFilePath = properties.TargetDirectory + file.RelativePath.Replace("/", "\\") + file.Name;
            }

            double downloadedSize = 0;
            double totalSize = allFiles.Where(f => !f.Exists).Sum(f => f.FileSize);

            foreach (var file in allFiles)
            {
                var localFilePath = properties.TargetDirectory + file.RelativePath.Replace("/", "\\") + file.Name;

                fd.DownloadFileProgress += (sender, progress) =>
                {
                    if (DownloadFileProgress == null) return;
                    if (progress.PercentageComplete == 100.0) return;

                    double bytesIn = downloadedSize + progress.BytesRecieved;
                    double percentage = bytesIn / totalSize * 100;

                    if (percentage > 100)
                    {
                        percentage = 100;
                    }

                    DownloadFileProgress(this, new Events.DownloadFileProgress()
                    {
                        BytesRecieved = (long)(downloadedSize + progress.BytesRecieved),
                        PercentageComplete = Math.Truncate(percentage),
                        TotalBytesToRecieve = (long)totalSize
                    });
                };

                if (file.Exists)
                {
                    continue;
                }

                await fd.DownloadAsync(file.RemoteUrl, localFilePath, token);
                downloadedSize += file.FileSize;

                if (token.IsCancellationRequested)
                {
                    return;
                }
            }

            foreach (var file in allFiles)
            {
                var localFilePath = properties.TargetDirectory + file.RelativePath.Replace("/", "\\") + file.Name;

                if (string.IsNullOrEmpty(file.Rename)) continue;
                var fInfo = new FileInfo(localFilePath);
                File.Copy(localFilePath, fInfo.Directory.FullName + @"\" + file.Rename, true);
            }

            double percentageEnd = downloadedSize / totalSize * 100;
            if (percentageEnd == 99.0) percentageEnd = 100;
            DownloadFileProgress(this, new Events.DownloadFileProgress()
            {
                BytesRecieved = (long)(downloadedSize),
                PercentageComplete = Math.Truncate(percentageEnd),
                TotalBytesToRecieve = (long)totalSize
            });

        }

        public async Task<List<UpdateFiles>> DownloadCab()
        {
            var guid = Guid.NewGuid().ToString();

            var localCabPath = Environment.ExpandEnvironmentVariables(@"%temp%\" + guid + ".cab");
            if (File.Exists(localCabPath)) File.Delete(localCabPath);

            var fd = new FileDownloader();
            await fd.DownloadAsync(OfficeVersionUrl, localCabPath);

            var cabExtractor = new CabExtractor(localCabPath);
            cabExtractor.ExtractCabFiles();

            var xml32Path = Environment.ExpandEnvironmentVariables(@"%temp%\ExtractedFiles\o365client_32bit.xml");
            var xml64Path = Environment.ExpandEnvironmentVariables(@"%temp%\ExtractedFiles\o365client_64bit.xml");

            var updateFiles32 = GenerateUpdateFiles(xml32Path);
            var updateFiles64 = GenerateUpdateFiles(xml64Path);

            return new List<UpdateFiles>()
            {
                updateFiles32,
                updateFiles64
            };
        }

        public async Task<string> GetLatestVersionAsync(string branch, OfficeEdition officeEdition)
        {
            if (_updateFiles == null)
            {
                _updateFiles = await DownloadCab();
            }

            var selectUpdateFiles = _updateFiles.FirstOrDefault(f => f.OfficeEdition == officeEdition);
            if (selectUpdateFiles == null) return null;

            var branchBaseUrl = selectUpdateFiles.BaseURL.FirstOrDefault(b => b.Branch.ToLower() == branch.ToLower());
            return await GetLatestVersionAsync(branchBaseUrl, officeEdition);
        }

        public async Task<string> GetLatestVersionAsync(baseURL branchUrl, OfficeEdition officeEdition)
        {
            var fileName = "v32.cab";
            if (officeEdition == OfficeEdition.Office64Bit)
            {
                fileName = "v64.cab";
            }

            var guid = Guid.NewGuid().ToString();

            var vcabFileDir = Environment.ExpandEnvironmentVariables(@"%temp%\OfficeProPlus\" + branchUrl.Branch + @"\" + guid);

            var vcabFilePath = vcabFileDir + @"\" + fileName;
            var vcabExtFilePath = vcabFileDir + @"\ExtractedFiles\VersionDescriptor.xml";

            Directory.CreateDirectory(vcabFileDir);

            var fd = new FileDownloader();
            await fd.DownloadAsync(branchUrl.URL + @"/Office/Data/" + fileName, vcabFilePath);

            var cabExtractor = new CabExtractor(vcabFilePath);
            cabExtractor.ExtractCabFiles();

            var version = GetCabVersion(vcabExtFilePath);
            return version;
        }

        private UpdateFiles GenerateUpdateFiles(string xmlFilePath)
        {
            var updateFiles = new UpdateFiles
            {
                OfficeEdition = OfficeEdition.Office64Bit
            };

            if (xmlFilePath.Contains("32"))
            {
                updateFiles.OfficeEdition = OfficeEdition.Office32Bit;
            }

            var xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFilePath);

            var baseUrlNodes = xmlDoc.SelectNodes("/UpdateFiles/baseURL");
            foreach (XmlNode baseUrlNode in baseUrlNodes)
            {
                var branch = baseUrlNode.GetAttributeValue("branch");
                var url = baseUrlNode.GetAttributeValue("URL");
                updateFiles.BaseURL.Add(new baseURL()
                {
                    Branch = branch,
                    URL = url
                });
            }

            var fileNodes = xmlDoc.SelectNodes("/UpdateFiles/File");
            foreach (XmlNode fileNode in fileNodes)
            {
                var name = fileNode.GetAttributeValue("name");
                var rename = fileNode.GetAttributeValue("rename");
                var relativePath = fileNode.GetAttributeValue("relativePath");
                var language = Convert.ToInt32(fileNode.GetAttributeValue("language"));
                updateFiles.Files.Add(new Model.File()
                {
                    Name = name,
                    Rename = rename,
                    RelativePath = relativePath,
                    Language = language
                });
            }
            return updateFiles;
        }

        private string GetCabVersion(string xmlFilePath)
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFilePath);

            var availableNode = xmlDoc.DocumentElement.SelectSingleNode("./Available");
            if (availableNode == null) return null;

            var buildVersion = availableNode.GetAttributeValue("Build");
            return buildVersion;
        }

        public Events.DownloadFileProgressEventHandler DownloadFileProgress { get; set; }

        public Events.VersionDetectedEventHandler VersionDetected { get; set; }

    }
}
