using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WF_SendMailAutomation
{
    public static class WriteLog
    {
        private static string fileName = Path.Combine(Environment.CurrentDirectory.Replace("\\bin\\Debug", ""), "Files", "log.txt");
        private static void Log(string logMessage, TextWriter w)
        {
            w.WriteLine($" {logMessage}");
            w.WriteLine("-------------------------------------");
        }

        public static void Write(string logMess)
        {
            using (StreamWriter w = File.AppendText(fileName))
            {
                Log(logMess, w);
            }
        }
    }
}
