using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeInstallGenerator;

namespace OfficeInstallGeneratorTests
{
    [TestClass]
    public class UnitTests
    {

        [TestMethod]
        public async Task TestProPlusGenerator()
        {

            var proPlusDownloader = new ProPlusDownloader();
            proPlusDownloader.DownloadFileProgress += (sender, progress) =>
            {
                var percent = progress.PercentageComplete;
                if (percent != null)
                {
                    
                }
            };

            await proPlusDownloader.DownloadBranch(new DownloadBranchProperties()
            {
                BranchName = "Current",
                OfficeEdition = OfficeEdition.Office64Bit,
                TargetDirectory = @"e:\Office",
                Languages = new List<string>() { "en-us", "fr-fr"}
            });

            
        }

        [TestMethod]
        public async Task TestProPlusDownloadVersionHistory()
        {

            var proPlusDownloader = new ProPlusDownloader();

            await proPlusDownloader.DownloadReleaseHistoryCabAsync();

        }

        [TestMethod]
        public void TestGetOfficeVersions()
        {

            var officeInstall = new InstallOffice2();

            officeInstall.GetOfficeVersion();
        }

        [TestMethod]
        public void TestConfigXmlParser()
        {


            var configXmlParser = new ConfigXmlParser(@"E:\Users\rsmith.vcg\Desktop\configurationTest.xml");

            var clientEdition = configXmlParser.ConfigurationXml.Add.OfficeClientEdition;


            var test = "";


        }

        [TestMethod]
        public void TestOfficeInstallExec()
        {
            var installGen = new OfficeInstallExecutableGenerator();


            //installGen.WaitForOfficeCTRUpadate();

        }

    }
}
