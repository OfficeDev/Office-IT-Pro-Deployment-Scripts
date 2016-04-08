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

            var officeInstall = new InstallOffice2();

            var officeProducts = officeInstall.GetOfficeVersion();
            if (officeProducts != null)
            {
                


            }

        }
    }
}
