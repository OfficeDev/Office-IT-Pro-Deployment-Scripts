using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


    public class BackOff
    {

        public BackOff()
        {
            a = 0;
            b = 1;
        }

        public async Task RunAsync()
        {
            var temp = a;
            a = b;
            b = temp + b;
            await Task.Delay(b * 1000);
        }

        public void Run()
        {
            var temp = a;
            a = b;
            b = temp + b;
            Thread.Sleep(b * 1000);
        }

        private int a { get; set; }
        private int b { get; set; }

    }

