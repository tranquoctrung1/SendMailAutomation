using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WF_SendMailAutomation.Helper;

namespace WF_SendMailAutomation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Run()
        {
            try
            {
                if (DateTime.Now.Minute % 1 == 0 && DateTime.Now.Second == 0)
                {
                    if (DateTime.Now.Minute == 0)
                    {
                        var writeExcelLog = (new SendMail()).WriteFile();
                        rtxtLogs.Text += writeExcelLog + "\r\n";
                        WriteLog.Write(writeExcelLog);
                    }

                    if (DateTime.Now.Day == 1 && DateTime.Now.Hour == 8 && DateTime.Now.Minute == 0 && DateTime.Now.Second == 0)
                    {
                        rtxtLogs.Text += (new SendMail()).Send();
                    }
                    rtxtLogs.Text += (new SendMailAlarm()).Send();
                    //rtxtLogs.Text += (new DeleteOldData()).log;
                }
            }
            catch (Exception ex)
            {
                rtxtLogs.Text += ex.ToString() + "/r/n";
                WriteLog.Write(ex.ToString());
            }
        }

        private void Form_Load(object sender, EventArgs e)
        {
            rtxtLogs.Text += "App start successful\r\n";
            rtxtLogs.Text += "--------------------------\r\n";
            var writeExcelLog = (new SendMail()).WriteFile();
            rtxtLogs.Text += writeExcelLog + "\r\n";
            WriteLog.Write(writeExcelLog);
            timer1.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            Run();
            //rtxtLogs.Text += "Hello\r\n";
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                this.ShowInTaskbar = false;
                notifyIcon1.Visible = true;
            }   
        }

        private void ntfIcon_DbClick(object sender, MouseEventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            notifyIcon1.Visible = false;
        }

        private void btnSendNow_Click(object sender, EventArgs e)
        {
            try
            {
                //var now = DateTime.Now;
                //var startThisMonth = new DateTime(now.Year, now.Month, 1);
                var log = (new SendMail()).Send();
                rtxtLogs.Text += log;
                WriteLog.Write(log);
            }
            catch(Exception ex)
            {
                rtxtLogs.Text += ex.ToString() + "/r/n";
                WriteLog.Write(ex.ToString());
            }
            
        }
    }
}
