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
            try
            {
                var productId = ""; //args[0];

                var officeInstall = new InstallOffice {ProductId = productId };
                officeInstall.IsLanguageInstalled(@"E:\Users\rsmith.VCG\Desktop\config1.xml", "en-us");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
            }
            finally
            {
                Console.WriteLine("Done");
                Console.ReadLine();
            }
        }
    }
}
