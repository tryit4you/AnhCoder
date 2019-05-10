using System;
using System.Collections.Generic;
using System.IO;
using System.ServiceProcess;
using System.Text;

namespace WindowService
{
    internal class TestSevice : ServiceBase
    {
        public TestSevice()
        {
            ServiceName = "TestService";
        }

        protected override void OnStart(string[] args)
        {
            string filename = CheckFileExists();
            File.AppendAllText(filename, $"{DateTime.Now} started.{Environment.NewLine}");
        }

        protected override void OnStop()
        {
            string filename = CheckFileExists();
            File.AppendAllText(filename, $"{DateTime.Now} stopped.{Environment.NewLine}");
        }

        private static string CheckFileExists()
        {
            string filename = @"c:\tmp\MyService.txt";
            if (!File.Exists(filename))
            {
                File.Create(filename);
            }

            return filename;
        }

    }
}
