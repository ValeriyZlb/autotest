using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace autotest
{
    class SPDevices
    {
        public string StartProperties;
        public string DeviceName;
        public SPDevices(string devname, string devid)
        {
            DeviceName = devname;
            StartProperties = devid;
        }
        public int RequestData(int RequestTime, TextBox consoleOut)
        {
            string RqTime = String.Format("/S={0:d2}000900 /F={1:d2}000800", RequestTime, RequestTime + 1);
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = @"C:\Program Files (x86)\Logika\Spnet95\Sphone95.exe";
            consoleOut.AppendText(DateTime.Now + " Опрос: " + DeviceName + "...\r\n");
            startInfo.Arguments = "/P=" + StartProperties + RqTime;
            startInfo.UseShellExecute = true;
            var procStat = Process.Start(startInfo);
            procStat.WaitForExit();
            return 0;
        }
    }
}
