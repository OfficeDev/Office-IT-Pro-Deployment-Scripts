using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeInstallGenerator
{
    public class CmdArguments
    {
        public static List<CmdArgument> GetArguments()
        {
            var returnArgs = new List<CmdArgument>();
            var args = Environment.GetCommandLineArgs();

            foreach (var arg in args)
            {
                var argSplit = new string[] {};
                if (arg.Contains(":"))
                {
                    argSplit = arg.Split(':');
                }
                if (arg.Contains("="))
                {
                    argSplit = arg.Split('=');
                }

                var newArg = new CmdArgument()
                {
                    Name = argSplit[0],
                    Value = argSplit[1]
                };
                returnArgs.Add(newArg);
            }

            return returnArgs;
        } 
    }

    public class CmdArgument
    {
        public string Name { get; set; }

        public string Value { get; set; }
    }
}
