using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace LaunchPowershellScript
{
    class Program
    {
        static void Main(string[] args)
        {
            //couple of things to note about this program
            //1. The powershell script you wish to run must not require an arguments, and must not require a function call
            //2. To launch, the call will look like the following: C:\temp>LaunchPowershellScript.exe psscript1.ps1
            //
            Process p = new Process();
            p.StartInfo.FileName = "Powershell.exe";
            p.StartInfo.Arguments = @"-ExecutionPolicy Bypass -NoExit -File .\"+args[0];
            p.Start();
            p.Close();
        }
    }
}
