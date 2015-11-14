using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeInstallGenerator;

namespace OfficeInstallGeneratorTests
{
    [TestClass]
    public class UnitTests
    {
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
