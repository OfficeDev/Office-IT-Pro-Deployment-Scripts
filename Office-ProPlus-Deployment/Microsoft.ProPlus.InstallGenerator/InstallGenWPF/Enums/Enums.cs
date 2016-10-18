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
        ManageCM = 4
    }

    public enum CMScenario
    {
        Deploy = 0, 
        ChangeChannel = 1, 
        Rollback = 2, 
        UpdateConfigMgr = 3, 
        UpdateScheduledTask = 4,
        DeployLanguagePack = 5
    }

    public enum DeploymentSource
    {
        CDN = 0,
        DistributionPoint =1,
        Local = 2
    }

    public enum BranchVersion
    {
        Latest = 0,
        Previous = 1
    }

    public enum ProductAction
    {
        Install = 0,
        Exclude = 1
    }

    public enum ProgramType
    {
        DeployWithScript = 0,
        DeployWithConfigurationFile = 1,
        ChangeChannel = 2,
        RollBack = 3,
        UpdateWithConfigMgr = 4,
        UpdateWithTask = 5
    }

    public enum DeploymentPurpose
    {
        Default = 0,
        Required = 1,
        Available = 2
    }

    public enum DeploymentType
    {
        DeployWithScript = 0,
        DeployWithConfigurationFile =1
    }
}
