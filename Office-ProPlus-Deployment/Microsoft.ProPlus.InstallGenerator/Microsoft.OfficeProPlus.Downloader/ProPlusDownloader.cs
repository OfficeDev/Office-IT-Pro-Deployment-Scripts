using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
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
        private const string OfficeVersionHistoryUrl = "http://officecdn.microsoft.com/pr/wsus/releasehistory.cab";
        private const string OfficeVersionWebSite = "https://technet.microsoft.com/en-us/library/mt592918.aspx";

        private List<UpdateFiles> _updateFiles { get; set; }


        private static AsyncLock myLock = new AsyncLock();
        private static AsyncLock myLock2 = new AsyncLock();

        public async Task DownloadBranch(DownloadBranchProperties properties, CancellationToken token = new CancellationToken())
        {
            var fd = new FileDownloader();

            if (properties.Languages == null) properties.Languages = new List<string>() { "en-us" };

            if (_updateFiles == null)
            {
                for (var t = 1; t <= 20; t++)
                {
                    try
                    {
                        _updateFiles = await DownloadCabAsync();
                        break;
                    }
                    catch (Exception ex)
                    {
                       
                    }
                    await Task.Delay(1000, token);
                }
            }

            var selectUpdateFile = new UpdateFiles();


            if (properties.OfficeEdition == OfficeEdition.Office32Bit)
            {
                selectUpdateFile = _updateFiles.FirstOrDefault(u => u.OfficeEdition == OfficeEdition.Office32Bit);
            }
            else if (properties.OfficeEdition == OfficeEdition.Office64Bit)
            {
                selectUpdateFile = _updateFiles.FirstOrDefault(u => u.OfficeEdition == OfficeEdition.Office64Bit);
            }
            else if (properties.OfficeEdition == OfficeEdition.Both)
            {
                var selectUpdateFile32 = _updateFiles.FirstOrDefault(u => u.OfficeEdition == OfficeEdition.Office32Bit);
                var selectUpdateFile64 = _updateFiles.FirstOrDefault(u => u.OfficeEdition == OfficeEdition.Office64Bit);

                selectUpdateFile32.Files.AddRange(selectUpdateFile64.Files);
                selectUpdateFile = selectUpdateFile32;
            }


            if (selectUpdateFile == null) throw (new Exception("Cannot Find Office Files"));

            var branch = selectUpdateFile.BaseURL.FirstOrDefault(b => b.Branch.ToLower() == properties.BranchName.ToLower());

            var version = properties.Version;
            if (string.IsNullOrEmpty(properties.Version))
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

            if (DownloadFileProgress != null)
            {
                DownloadFileProgress(this, new Events.DownloadFileProgress()
                {
                    BytesRecieved = (long) (downloadedSize),
                    PercentageComplete = Math.Truncate(percentageEnd),
                    TotalBytesToRecieve = (long) totalSize
                });
            }

            if (DownloadFileComplete != null)
            {
                DownloadFileComplete(this, new Events.DownloadFileProgress()
                {
                    BytesRecieved = (long) (downloadedSize),
                    PercentageComplete = Math.Truncate(percentageEnd),
                    TotalBytesToRecieve = (long) totalSize
                });
            }

        }

        public async Task<bool> ValidateSourceFiles(DownloadBranchProperties properties, CancellationToken token = new CancellationToken())
        {
            var fd = new FileDownloader();

            if (properties.Languages == null) properties.Languages = new List<string>() { "en-us" };

            if (_updateFiles == null)
            {
                for (var t = 1; t <= 20; t++)
                {
                    try
                    {
                        _updateFiles = await DownloadCabAsync();
                        break;
                    }
                    catch (Exception ex)
                    {

                    }
                    await Task.Delay(1000, token);
                }
            }

            var selectUpdateFile = new UpdateFiles();


            if (properties.OfficeEdition == OfficeEdition.Office32Bit)
            {
                selectUpdateFile = _updateFiles.FirstOrDefault(u => u.OfficeEdition == OfficeEdition.Office32Bit);
            }
            else if (properties.OfficeEdition == OfficeEdition.Office64Bit)
            {
                selectUpdateFile = _updateFiles.FirstOrDefault(u => u.OfficeEdition == OfficeEdition.Office64Bit);
            }
            else if (properties.OfficeEdition == OfficeEdition.Both)
            {
                var selectUpdateFile32 = _updateFiles.FirstOrDefault(u => u.OfficeEdition == OfficeEdition.Office32Bit);
                var selectUpdateFile64 = _updateFiles.FirstOrDefault(u => u.OfficeEdition == OfficeEdition.Office64Bit);

                selectUpdateFile32.Files.AddRange(selectUpdateFile64.Files);
                selectUpdateFile = selectUpdateFile32;
            }


            if (selectUpdateFile == null) throw (new Exception("Cannot Find Office Files"));

            var branch = selectUpdateFile.BaseURL.FirstOrDefault(b => b.Branch.ToLower() == properties.BranchName.ToLower());

            var version = properties.Version;
            if (string.IsNullOrEmpty(properties.Version))
            {
                version = await GetLatestVersionAsync(branch, properties.OfficeEdition);
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
                }
            }

            allFiles = allFiles.Distinct().ToList();

            foreach (var file in allFiles)
            {
                file.LocalFilePath = properties.TargetDirectory + file.RelativePath.Replace("/", "\\") + file.Name;
            }

            foreach (var file in allFiles)
            {
                var localFilePath = properties.TargetDirectory + file.RelativePath.Replace("/", "\\") + file.Name;
                if (!File.Exists(localFilePath))
                {
                    return false;
                }
            }

            return true;
        }


        public async Task<List<UpdateFiles>> DownloadCabAsync()
        {
            var guid = Guid.NewGuid().ToString();

            var cabPath = Environment.ExpandEnvironmentVariables(@"%temp%\" + guid);
            Directory.CreateDirectory(cabPath);

            var localCabPath = Environment.ExpandEnvironmentVariables(@"%temp%\" + guid + @"\" + guid + ".cab");
            if (File.Exists(localCabPath)) File.Delete(localCabPath);

            var now = DateTime.Now;

            var tmpFile32 = Environment.ExpandEnvironmentVariables(@"%temp%\" + now.Year + now.Month + now.Day + now.Hour + "_32.xml");
            var tmpFile64 = Environment.ExpandEnvironmentVariables(@"%temp%\" + now.Year + now.Month + now.Day + now.Hour + "_64.xml");

            if (!File.Exists(tmpFile32) || !File.Exists(tmpFile64))
            {
                using (var releaser = await myLock.LockAsync())
                {
                    var tmpFile =
                        Environment.ExpandEnvironmentVariables(@"%temp%\" + now.Year + now.Month + now.Day + now.Hour +
                                                               ".cab");

                    if (File.Exists(tmpFile))
                    {
                        Retry.Block(10, 1, () => File.Copy(tmpFile, localCabPath, true));
                    }

                    if (!File.Exists(localCabPath))
                    {
                        var fd = new FileDownloader();
                        await fd.DownloadAsync(OfficeVersionUrl, localCabPath);
                        try
                        {
                            File.Copy(localCabPath, tmpFile);
                        }
                        catch
                        {
                        }
                    }

                    var cabExtractor = new CabExtractor(localCabPath);
                    cabExtractor.ExtractCabFiles();
                }
            }

            var xml32Path = Environment.ExpandEnvironmentVariables(@"%temp%\" + guid + @"\ExtractedFiles\o365client_32bit.xml");
            var xml64Path = Environment.ExpandEnvironmentVariables(@"%temp%\" + guid + @"\ExtractedFiles\o365client_64bit.xml");

            if (File.Exists(tmpFile32))
            {
                xml32Path = tmpFile32;
            }
            else
            {
                Retry.Block(10, 1, () => File.Copy(xml32Path, tmpFile32, true));
            }

            if (File.Exists(tmpFile64))
            {
                xml64Path = tmpFile64;
            }
            else
            {
                Retry.Block(10, 1, () => File.Copy(xml64Path, tmpFile64, true));
            }

            var updateFiles32 = GenerateUpdateFiles(xml32Path);
            var updateFiles64 = GenerateUpdateFiles(xml64Path);

            try
            {
                if (File.Exists(localCabPath)) File.Delete(localCabPath);
                if (Directory.Exists(cabPath)) Directory.Delete(cabPath, true);
            }
            catch { }

            return new List<UpdateFiles>()
            {
                updateFiles32,
                updateFiles64
            };
        }

        public async Task<List<UpdateChannel>> DownloadReleaseHistoryCabAsync()
        {
            var guid = Guid.NewGuid().ToString();

            var cabPath = Environment.ExpandEnvironmentVariables(@"%temp%\" + guid);
            Directory.CreateDirectory(cabPath);

            var savePath = Environment.ExpandEnvironmentVariables(@"%temp%\ReleaseHistory.xml");
            var releaseHistoryPath = "";

            var localCabPath = Environment.ExpandEnvironmentVariables(@"%temp%\" + guid + @"\" + guid + ".cab");
            try
            {
                var now = DateTime.Now;

                var tmpReleaseHistory =
                    Environment.ExpandEnvironmentVariables(@"%temp%\" + now.Year + now.Month + now.Day + now.Hour +
                                                           "_ReleaseHistory.xml");
                using (var releaser = await myLock.LockAsync())
                {
                    if (File.Exists(localCabPath)) File.Delete(localCabPath);

                    if (!File.Exists(tmpReleaseHistory))
                    {
                        if (!File.Exists(localCabPath))
                        {
                            var fd = new FileDownloader();
                            await fd.DownloadAsync(OfficeVersionHistoryUrl, localCabPath);
                        }
                        var cabExtractor = new CabExtractor(localCabPath);
                        cabExtractor.ExtractCabFiles();
                    }

                    var dirInfo = new DirectoryInfo(Environment.ExpandEnvironmentVariables(@"%temp%"));
                    foreach (var file in dirInfo.GetFiles("*_ReleaseHistory.xml"))
                    {
                        try
                        {
                            if (!file.FullName.Equals(tmpReleaseHistory, StringComparison.InvariantCultureIgnoreCase))
                            {
                                File.Delete(file.FullName);
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }

                    releaseHistoryPath =
                        Environment.ExpandEnvironmentVariables(@"%temp%\" + guid + @"\ExtractedFiles\ReleaseHistory.xml");
                    if (File.Exists(tmpReleaseHistory))
                    {
                        releaseHistoryPath = tmpReleaseHistory;
                    }
                    else
                    {
                        Retry.Block(10, 1, () => File.Copy(releaseHistoryPath, tmpReleaseHistory, true));
                    }

                }
            }
            catch (Exception ex)
            {
                releaseHistoryPath = savePath;
            }

            var releaseHistory = GenerateReleaseHistory(releaseHistoryPath);
            try
            {
                using (var releaser = await myLock.LockAsync())
                {
                    if (!string.IsNullOrEmpty(releaseHistoryPath))
                    {
                        if (File.Exists(releaseHistoryPath))
                        {
                            if (releaseHistoryPath != savePath)
                            {
                                File.Copy(releaseHistoryPath, savePath, true);
                            }
                        }
                    }
                    if (File.Exists(localCabPath)) File.Delete(localCabPath);
                    if (Directory.Exists(cabPath)) Directory.Delete(cabPath, true);
                }
            }
            catch { }

            return releaseHistory;
        }

        public async Task<List<UpdateChannel>> DownloadVersionsFromWebSite()
        {
            var lstReturn = new List<UpdateChannel>();
            var ccUpdateChannel = new UpdateChannel()
            {
                Name = "Current",
                Updates = new List<Update>()
            };
            var dcUpdateChannel = new UpdateChannel()
            {
                Name = "Deferred",
                Updates = new List<Update>()
            };
            var frdcUpdateChannel = new UpdateChannel()
            {
                Name = "FirstReleaseDeferred",
                Updates = new List<Update>()
            };
            lstReturn.Add(ccUpdateChannel);
            lstReturn.Add(dcUpdateChannel);
            lstReturn.Add(frdcUpdateChannel);

            var webClient = new WebClient();
            var page = await webClient.DownloadStringTaskAsync(OfficeVersionWebSite);

            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(page);

            var tables = doc.DocumentNode.SelectNodes("//table").ToList();

            foreach (var table in tables)
            {
                var headerRow = table
                    .Descendants("thead")
                    .Where(thhead => thhead.Elements("tr").Count() == 1)
                    .Select(tr => tr.Descendants("td").Select(td => td.InnerText.Trim()).ToList())
                    .FirstOrDefault();
                if (headerRow == null) continue;
                if (headerRow[0].ToLower() != "version") continue;
             
                var rows = table
                    .Descendants("tr")
                    .Where(tr => tr.Elements("td").Count() > 1)
                    .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
                    .ToList();
                if (rows.Count == 0) continue;

                var ccIndex = 1;
                var frdcIndex = 2;
                var dcIndex = 3;
                for (var i = 1; i <= 3; i++)
                {
                    var rowHeader = headerRow[i];
                    if (rowHeader.ToLower().Contains("current"))
                    {
                        ccIndex = i;
                    }
                    if (rowHeader.ToLower().Contains("deferred") &&
                        rowHeader.ToLower().Contains("first release"))
                    {
                        frdcIndex = i;
                    }
                    if (rowHeader.ToLower().Contains("deferred") &&
                        !rowHeader.ToLower().Contains("first release"))
                    {
                        dcIndex = i;
                    }
                }

                foreach (var row in rows)
                {
                    var version = row[0];
                    if (!Regex.Match(version, @"\d{4}").Success) continue;

                    var currentChannel = row[ccIndex];
                    var ccVersions = Regex.Matches(currentChannel, @"\d{4}\.\d{4}\s");
                    

                    var frdc = row[frdcIndex];
                    var frdcVersions = Regex.Matches(frdc, @"\d{4}\.\d{4}");
                    

                    var dc = row[dcIndex];
                    var dcVersions = Regex.Matches(dc, @"\d{4}\.\d{4}");
                    
                    
                    foreach (Match build in ccVersions)
                    {
                        ccUpdateChannel?.Updates.Add(new Update()
                        {
                            Build = build.Value,
                            LegacyVersion = "16.0." + build.Value,
                            Version = version
                        });
                    }

                    foreach (Match build in dcVersions)
                    {
                        dcUpdateChannel?.Updates.Add(new Update()
                        {
                            Build = build.Value,
                            LegacyVersion = "16.0." + build.Value,
                            Version = version
                        });
                    }

                    foreach (Match build in frdcVersions)
                    {
                        frdcUpdateChannel?.Updates.Add(new Update()
                        {
                            Build = build.Value,
                            LegacyVersion = "16.0." + build.Value,
                            Version = version
                        });
                    }

                }
            }

            return lstReturn;
        }



        public async Task<string> GetChannelNameFromUrlAsync(string channelUrl, OfficeEdition officeEdition)
        {
            if (_updateFiles == null)
            {
                using (var releaser = await myLock2.LockAsync())
                {
                    if (_updateFiles == null)
                    {
                        await Retry.BlockAsync(10, 1, async () => {
                            _updateFiles = await DownloadCabAsync();
                        });
                    }
                }
            }

            var selectUpdateFiles = _updateFiles.FirstOrDefault(f => f.OfficeEdition == officeEdition);
            if (selectUpdateFiles == null) return null;

            var branchBase = selectUpdateFiles.BaseURL.FirstOrDefault(b => b.URL.ToLower() == channelUrl.ToLower());
            if (branchBase == null) return null;
            return branchBase.Branch;
        }

        public async Task<string> GetChannelBaseUrlAsync(string channel, OfficeEdition officeEdition)
        {
            if (_updateFiles == null)
            {
                using (var releaser = await myLock2.LockAsync())
                {
                    if (_updateFiles == null)
                    {
                        await Retry.BlockAsync(10, 1, async () => {
                            _updateFiles = await DownloadCabAsync();
                        });
                    }
                }
            }

            var selectUpdateFiles = _updateFiles.FirstOrDefault(f => f.OfficeEdition == officeEdition);
            if (selectUpdateFiles == null) return null;

            var branchBase = selectUpdateFiles.BaseURL.FirstOrDefault(b => b.Branch.ToLower() == channel.ToLower());
            if (branchBase == null) return null;
            return branchBase.URL;
        }

        public async Task<string> GetLatestVersionAsync(string branch, OfficeEdition officeEdition)
        {
            if (_updateFiles == null)
            {
                using (var releaser = await myLock2.LockAsync())
                {
                    if (_updateFiles == null)
                    {
                        _updateFiles = await DownloadCabAsync();
                    }
                }
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

        public async Task<string> GetChannelVersionJson()
        {
            var channelVersionJson = "";
            using (var webClient = new WebClient())
            {
                channelVersionJson = await webClient.DownloadStringTaskAsync("https://officeproplusinfo2.azurewebsites.net/api/Channel");
                //channelVersionJson = await webClient.DownloadStringTaskAsync(AppSettings.BranchVersionUrl);
            }
            return channelVersionJson;
        }

        private List<UpdateChannel> GenerateReleaseHistory(string xmlFilePath)
        {
            var updateChannels = new List<UpdateChannel>();

            var xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFilePath);

            var xmlUpdateChannels = xmlDoc.SelectNodes("/ReleaseHistory/UpdateChannel");
            foreach (XmlNode xmlUpdateChannel in xmlUpdateChannels)
            {
                var name = xmlUpdateChannel.GetAttributeValue("Name");

                var updateChannel = new UpdateChannel
                {
                    Name = name,
                    Updates = new List<Update>()
                };

                var updates = xmlUpdateChannel.SelectNodes("./Update");                
                foreach (XmlNode xmlUpdate in updates)
                {
                    var latest = Convert.ToBoolean(xmlUpdate.GetAttributeValue("Latest"));
                    var version = xmlUpdate.GetAttributeValue("Version");
                    var legacyVersion = xmlUpdate.GetAttributeValue("LegacyVersion");
                    var build = xmlUpdate.GetAttributeValue("Build");
                    var pubTime = xmlUpdate.GetAttributeValue("PubTime");

                    var publishTime = XmlConvert.ToDateTime(pubTime, XmlDateTimeSerializationMode.Utc);

                    var update = new Update()
                    {
                        Build = build,
                        Latest = latest,
                        LegacyVersion = legacyVersion,
                        Version = version,
                        PublishTime = publishTime
                    };

                    updateChannel.Updates.Add(update);
                }

                updateChannels.Add(updateChannel);
            }

            return updateChannels;
        }

        private UpdateFiles GenerateUpdateFiles(string xmlFilePath)
        {
            var updateFiles = new UpdateFiles
            {
                OfficeEdition = OfficeEdition.Office64Bit
            };

            if (xmlFilePath.Contains("_32"))
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

        public string GetCabVersion(string xmlFilePath)
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFilePath);

            if (xmlDoc.DocumentElement == null) return null;
            var availableNode = xmlDoc.DocumentElement.SelectSingleNode("./Available");

            var buildVersion = availableNode?.GetAttributeValue("Build");
            return buildVersion;
        }

        public Events.DownloadFileProgressEventHandler DownloadFileComplete { get; set; }

        public Events.DownloadFileProgressEventHandler DownloadFileProgress { get; set; }

        public Events.VersionDetectedEventHandler VersionDetected { get; set; }

    }
}
