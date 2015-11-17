using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeInstallGenerator;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeInstallMsiGenerator : IOfficeInstallGenerator
    {
        public IOfficeInstallReturn Generate(IOfficeInstallProperties installProperties)
        {
            var exeGenerator = new OfficeInstallExecutableGenerator();
            var exeReturn = exeGenerator.Generate(installProperties);

            var exeFilePath = exeReturn.GeneratedFilePath;

            return null;
        }

    }
}
