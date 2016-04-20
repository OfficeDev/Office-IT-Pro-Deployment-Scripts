using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.OfficeProPlus.InstallGenerator.Extensions;

namespace TestCode
{
    class Program
    {
        static void Main(string[] args)
        {

            var strEnus = "en-us";
            var guid = strEnus.GenerateGuid();

            var officeInstall = new InstallOffice2();

            var officeProducts = officeInstall.GetOfficeVersion();
            if (officeProducts != null)
            {
                


            }

        }
    }
}
