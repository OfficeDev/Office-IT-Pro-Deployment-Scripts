using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCode
{
    class Program
    {
        static void Main(string[] args)
        {

            var officeInstall = new InstallOffice();

            officeInstall.UpdateLanguagePackInstall(@"E:\Users\rsmith.VCG\Desktop\config1.xml", true);

        }
    }
}
