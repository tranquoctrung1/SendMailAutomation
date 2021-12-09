using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace WF_SendMailAutomation.Helper
{
    public class SendMailAlarm
    {
        string connectionString = ConfigurationManager.ConnectionStrings["dtb"].ConnectionString;
        private string senderEmail = ConfigurationManager.AppSettings["senderEmail"].ToString();
        private string senderPass = ConfigurationManager.AppSettings["senderPass"].ToString();
        private string receiverEmail = ConfigurationManager.AppSettings["receiverEmail"].ToString();
        private string subject = "Alarm";
        private string now = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

        public string Send()
        {
            string log = "";
            var thisMonth = DateTime.Now.ToString("MM_yyyy");
            string filePath = Path.Combine(Environment.CurrentDirectory.Replace("\\bin\\Debug", ""), "Files", "Report_" + thisMonth + ".xls");
            log = now + ": File path is: " + filePath + "\r\n";

            var tagNames = GetWarningTagName();
            if (tagNames.Any())
            {
                string body = "";
                foreach (var name in tagNames)
                {
                    body += name + " Error<br/>";
                }
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(senderEmail);
                mail.To.Add(receiverEmail);
                mail.Subject = subject + "_" + now;
                mail.Body = body;
                mail.IsBodyHtml = true;
                using (var file = new Attachment(filePath))
                {
                    mail.Attachments.Add(file);
                    using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                    {
                        //smtp.Host = "smtp.gmail.com";
                        smtp.Credentials = new System.Net.NetworkCredential(senderEmail, senderPass);
                        smtp.EnableSsl = true;

                        try
                        {
                            smtp.Send(mail);
                            WriteLog.Write(now + ": Gửi mail cảnh báo thành công");
                            log = now + ": Send alarm mail Successful\r\n";
                        }
                        catch (Exception ex)
                        {
                            WriteLog.Write(now + ": Gửi mail cảnh báo thất bại: " + ex.ToString());
                            log = now + ": Send alarm mail ERROR\r\n";
                        }
                    }
                }
                
                mail.Dispose();
            }
            else
            {
                log = now + ": No alarm\r\n";
            }
            return log;
        }

        private List<string> GetWarningTagName()
        {
            var tagNames = new List<string>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand("p_Get_Alarm_Tags", connection);
                cmd.CommandType = CommandType.StoredProcedure;
                connection.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    tagNames.Add((string)rd["TagName"]);
                }
            }
            return tagNames;
        }
    }
}
