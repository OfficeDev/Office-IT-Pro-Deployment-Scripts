using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Enums
{
    public enum ApplicationMode
    {
        InstallGenerator = 0,
        ManageLocal = 1,
        LanguagePack = 2,
        ManageRemote = 3,
        ManageSccm = 4
    }

    public enum SccmScenario
    {
        Deploy = 0, 
        ChangeChannel = 1, 
        Rollback = 2, 
        UpdateConfigMgr = 3, 
        UpdateScheduledTask = 4 
    }

    public enum DeploymentSource
    {
        CDN = 0,
        DistributionPoint =1,
        Local = 2
    }
}
