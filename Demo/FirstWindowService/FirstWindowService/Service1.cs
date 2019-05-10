using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace FirstWindowService
{
    public partial class Service1 : ServiceBase
    {
        private Timer timer = null;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            timer = new Timer();
            timer.Interval = 60000;
            timer.Elapsed += timer_Tick;
            timer.Enabled = true;
            Utilities.WriteLogError("Test for 1st run WindowsService");
        }
        private void timer_Tick(object sender, ElapsedEventArgs args)
        {
            // Xử lý một vài logic ở đây
            Utilities.WriteLogError("Timer has ticked for doing something!!!");
        }
        protected override void OnStop()
        {
            // Ghi log lại khi Services đã được stop
            timer.Enabled = false;
            Utilities.WriteLogError("1st WindowsService has been stop");
        }
    }
}
