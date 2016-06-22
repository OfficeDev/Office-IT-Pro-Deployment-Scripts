using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Micorosft.OfficeProPlus.ConfigurationXml
{
    public enum Branch
    {
        Current=0,
        FirstReleaseCurrent=4,
        Business=1,
        Validation=2,
        FirstReleaseBusiness=3
    }
    public enum ODTChannel
    {
        Current = 0,
        FirstReleaseCurrent = 4,
        Deferred = 1,
        Validation = 2,
        FirstReleaseDeferred = 3
    }
}
