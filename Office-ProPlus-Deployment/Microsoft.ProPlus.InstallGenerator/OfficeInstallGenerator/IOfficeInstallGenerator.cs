using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGenerator
{
    public interface IOfficeInstallGenerator
    {

        IOfficeInstallReturn Generate(IOfficeInstallProperties installProperties, string remoteLogPath = "", string uninstallstring = "", string changeChannelString = "");

        void InstallOffice(string configurationXml);

 

    }
}
