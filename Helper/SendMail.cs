using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using System.Net.Mail;
using System.IO;
using System.Diagnostics;
using NPOI.POIFS.FileSystem;

namespace WF_SendMailAutomation.Helper
{
    public class SendMail
    {
        private string senderEmail = ConfigurationManager.AppSettings["senderEmail"].ToString();
        private string senderPass = ConfigurationManager.AppSettings["senderPass"].ToString();
        private string receiverEmail = ConfigurationManager.AppSettings["receiverEmail"].ToString();
        private string ccMail = ConfigurationManager.AppSettings["ccMail"].ToString();

        private string body = "See attached file below";
        private string now = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        public string log = "";

        public string Send()
        {
            var thisMonth = "";
            if(DateTime.Now.Day == 1)
            {
                thisMonth = DateTime.Now.AddMonths(-1).ToString("MM_yyyy");
            }
            else
            {
                thisMonth = DateTime.Now.ToString("MM_yyyy");
            }
            
            string filePath = Path.Combine(Environment.CurrentDirectory.Replace("\\bin\\Debug", ""), "Files", "Report_" + thisMonth + ".xls");
            log = now + ": File path is: " + filePath + "\r\n";
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(senderEmail);
                    mail.To.Add(receiverEmail);
                    mail.CC.Add(ccMail);
                    mail.Subject = "Report_" + thisMonth;
                    mail.Body = body;
                    mail.IsBodyHtml = true;
                    //filePath = Path.Combine(Environment.CurrentDirectory.Replace("\\bin\\Debug", ""), "Files", "test2.xls");
                    using (var file = new Attachment(filePath))
                    {
                        mail.Attachments.Add(file);
                        using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                        {
                            //smtp.Host = "smtp.gmail.com";
                            smtp.Credentials = new System.Net.NetworkCredential(senderEmail, senderPass);
                            smtp.EnableSsl = true;
                            smtp.Timeout = 300000;

                            smtp.Send(mail);
                            WriteLog.Write(now + ": Gửi mail thành công");
                            log += now + ": Send Monthly Report " + thisMonth + " Successful\r\n";
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                WriteLog.Write(now + ": Gửi mail thất bại: " + ex.ToString());
                log += now + ": Send Monthly Report " + thisMonth + " ERROR: No Data To Send\r\n";
            }
            return log;
        }

        private void CreateFile(DateTime start, DateTime end)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            var data = (new GetData(start, end)).lst;
            //var tags = data.Select(d => d.TagName).Distinct().ToList();
            //var tags = new List<string>() {
            //    "[FPC_1]VFD1_CURR_FB", "[FPC_1]VFD2_CURR_FB", "[FPC_1]VFD3C_CURR_FB", "[FPC_1]VFD3D_CURR_FB",
            //    "[FPC_1]VFD1_SPD_FB", "[FPC_1]VFD2_SPD_FB", "[FPC_1]VFD3C_SPD_FB", "[FPC_1]VFD3D_SPD_FB",
            //};

            var tags = new List<string>() {
                "[FPC_1]VFD1_CURR_FB", "[FPC_1]VFD2_CURR_FB", "[FPC_1]VFD3C_CURR_FB", "[FPC_1]VFD3D_CURR_FB",
                "[FPC_1]VFD1_SPD_FB", "[FPC_1]VFD2_SPD_FB", "[FPC_1]VFD3C_SPD_FB", "[FPC_1]VFD3D_SPD_FB",
                "[FPC_1]VFD1_KW_FB", "[FPC_1]VFD2_KW_FB", "[FPC_1]VFD3C_KW_FB",
                "[FPC_1]VFD3D_KW_FB", "Power1.Power.Power", "Power2.Power.Power", "FLOWRATE(m3/h)"
            };
            //var tagHeaders = new List<string>() {
            //    "1P1A_VFD1_CURRENT_FB", "1P1B_VFD2_CURRENT_FB", "1P1C_VFD3_CURRENT_FB", "1P1D_VFD3_CURRENT_FB",
            //    "1P1A_VFD1_SPEED_FB", "1P1B_VFD2_SPEED_FB", "1P1C_VFD3_SPEED_FB", "1P1D_VFD3_SPEED_FB"
            //};

            var tagHeaders = new List<string>()  {
                "[FPC_1]VFD1_CURR_FB", "[FPC_1]VFD2_CURR_FB", "[FPC_1]VFD3C_CURR_FB", "[FPC_1]VFD3D_CURR_FB",
                "[FPC_1]VFD1_SPD_FB", "[FPC_1]VFD2_SPD_FB", "[FPC_1]VFD3C_SPD_FB", "[FPC_1]VFD3D_SPD_FB",
                "[FPC_1]VFD1_KW_FB", "[FPC_1]VFD2_KW_FB", "[FPC_1]VFD3C_KW_FB",
                "[FPC_1]VFD3D_KW_FB",  "LV-Multimeter 1", "LV-Multimeter 2", "FLOWRATE(m3/h)"
            };

            var timeStamps = data.Select(d => d.DateAndTime).Distinct().ToList();

            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            var sheet = hssfworkbook.CreateSheet("Sheet1");
            //CREATE HEADERS
            var row0 = sheet.CreateRow(0);
            for (int i = 0; i < tagHeaders.Count; i++)
            {
                var cell_i0 = row0.CreateCell(i);
                if (i > 0)
                {
                    cell_i0.SetCellValue(tagHeaders[i - 1]);
                }
                else cell_i0.SetCellValue("TimeStamp/TagName");
            }
            //CREATE DATA
            int row_c = 1;
            foreach (var timeStamp in timeStamps)
            {
                var row = sheet.CreateRow(row_c);
                var cell_1 = row.CreateCell(0);
                cell_1.SetCellValue(timeStamp.ToString("dd/MM/yyyy HH:mm:ss"));
                var row_data = data.Where(d => d.DateAndTime == timeStamp);
                for (int i = 1; i <= tags.Count; i++)
                {
                    var cell = row.CreateCell(i);
                    var cell_data = row_data.FirstOrDefault(d => d.TagName == tags[i]);
                    if (cell_data != null)
                    {
                        if (cell_data.Val != null)
                        {
                            //Làm tròn 2 số lẻ
                            cell.SetCellValue(Math.Round((double)cell_data.Val, 2));
                        }
                        else cell.SetCellValue("");
                    }
                    else cell.SetCellValue("");
                }
                row_c++;
            }


            for (int i = 0; i < tags.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            ((HSSFSheet)hssfworkbook.GetSheetAt(0)).AlternativeFormula = false;
            ((HSSFSheet)hssfworkbook.GetSheetAt(0)).AlternativeExpression = false;

            FileStream file = new FileStream(Path.Combine(Environment.CurrentDirectory.Replace("\\bin\\Debug", ""), "Files", "Report_" + start.ToString("MM_yyyy") + ".xls"), FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
            hssfworkbook.Close();

            sw.Stop();
            log += "Create file successful. Time esplase: " + sw.ElapsedMilliseconds + "ms";
        }

        public string WriteFile()
        {
            try
            {
                //Check file có chưa, chưa thì tạo, có thì đọc
                var filename = "";
                if (DateTime.Now.Day == 1)
                {
                    filename = filename = "Report_" + DateTime.Now.AddMonths(-1).ToString("MM_yyyy") + ".xls";
                }
                else
                {
                    filename = "Report_" + DateTime.Now.ToString("MM_yyyy") + ".xls";
                }
                
                var path = Path.Combine(Environment.CurrentDirectory.Replace("\\bin\\Debug", ""), "Files", filename);
                FileStream file = null;
                HSSFWorkbook hssfworkbook = null;
                if (File.Exists(path))
                {
                    file = new FileStream(path, FileMode.Open, FileAccess.Read);
                    hssfworkbook = new HSSFWorkbook(file);
                }
                else
                {
                    hssfworkbook = new HSSFWorkbook();
                }

                //Mở file ra
                var sheet = File.Exists(path) ? hssfworkbook.GetSheet("Sheet1") : hssfworkbook.CreateSheet("Sheet1");

                //Lấy time cuối cùng từng add vô file
                var start = (new Model.LastTimeUpdateRepository()).GetLastTimeUpdate();
                if (start == null)
                {
                    start = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                }

                //Lấy ra data theo giờ
                var data = (new GetData((DateTime)start, DateTime.Now)).lst;
                WriteLog.Write("LastUpdate: " + ((DateTime)start).ToString("dd/MM/yyyy HH:mm:ss") + ";Now: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + ";DataCount: " + data.Count);
                if (data.Any())
                {
                    //var tags = new List<string>() {
                    //    "[FPC_1]VFD1_CURR_FB", "[FPC_1]VFD2_CURR_FB", "[FPC_1]VFD3C_CURR_FB", "[FPC_1]VFD3D_CURR_FB",
                    //    "[FPC_1]VFD1_SPD_FB", "[FPC_1]VFD2_SPD_FB", "[FPC_1]VFD3C_SPD_FB", "[FPC_1]VFD3D_SPD_FB"
                    //};
                    //var tagHeaders = new List<string>() {
                    //    "1P1A_VFD1_CURRENT_FB", "1P1B_VFD2_CURRENT_FB", "1P1C_VFD3_CURRENT_FB", "1P1D_VFD3_CURRENT_FB",
                    //    "1P1A_VFD1_SPEED_FB", "1P1B_VFD2_SPEED_FB", "1P1C_VFD3_SPEED_FB", "1P1D_VFD3_SPEED_FB"
                    //};

                    var tags = new List<string>() {
                        "[FPC_1]VFD1_CURR_FB", "[FPC_1]VFD2_CURR_FB", "[FPC_1]VFD3C_CURR_FB", "[FPC_1]VFD3D_CURR_FB",
                        "[FPC_1]VFD1_SPD_FB", "[FPC_1]VFD2_SPD_FB", "[FPC_1]VFD3C_SPD_FB", "[FPC_1]VFD3D_SPD_FB",
                        "[FPC_1]VFD1_KW_FB", "[FPC_1]VFD2_KW_FB", "[FPC_1]VFD3C_KW_FB",
                        "[FPC_1]VFD3D_KW_FB",  "Power1.Power.Power", "Power2.Power.Power", "FLOWRATE(m3/h)"
                    };
                    var tagHeaders = new List<string>()  {
                        "[FPC_1]VFD1_CURR_FB", "[FPC_1]VFD2_CURR_FB", "[FPC_1]VFD3C_CURR_FB", "[FPC_1]VFD3D_CURR_FB",
                        "[FPC_1]VFD1_SPD_FB", "[FPC_1]VFD2_SPD_FB", "[FPC_1]VFD3C_SPD_FB", "[FPC_1]VFD3D_SPD_FB",
                        "[FPC_1]VFD1_KW_FB", "[FPC_1]VFD2_KW_FB", "[FPC_1]VFD3C_KW_FB",
                        "[FPC_1]VFD3D_KW_FB",   "LV-Multimeter 1", "LV-Multimeter 2", "FLOWRATE(m3/h)"
                    };

                    var timeStamps = data.Select(d => d.DateAndTime).Distinct().ToList();

                    var row_num = sheet.PhysicalNumberOfRows;
                    //Chưa có row thì tạo header
                    if (row_num == 0)
                    {
                        var row0 = sheet.CreateRow(0);
                        var cell_i0 = row0.CreateCell(0);
                        cell_i0.SetCellValue("TimeStamp/TagName");
                        for (int i = 0; i < tagHeaders.Count; i++)
                        {
                            var cell_i = row0.CreateCell(i + 1);
                            cell_i.SetCellValue(tagHeaders[i]);
                        }
                        row_num++;
                    }
                    //có rồi thì add data mới vô

                    foreach (var timeStamp in timeStamps)
                    {
                        var row = sheet.CreateRow(row_num);
                        var cell_1 = row.CreateCell(0);
                        cell_1.SetCellValue(timeStamp.ToString("dd/MM/yyyy HH:mm:ss"));
                        var row_data = data.Where(d => d.DateAndTime == timeStamp);
                        for (int i = 0; i < tags.Count; i++)
                        {
                            var cell = row.CreateCell(i + 1);
                            var cell_data = row_data.FirstOrDefault(d => d.TagName == tags[i]);
                            if (cell_data != null)
                            {
                                if (cell_data.Val != null)
                                {
                                    //Làm tròn 2 số lẻ
                                    cell.SetCellValue(Math.Round((double)cell_data.Val, 2));
                                }
                                else cell.SetCellValue("");
                            }
                            else cell.SetCellValue("");
                        }
                        row_num++;
                    }

                    if (timeStamps.Any())
                    {
                        for (int i = 0; i < tags.Count + 1; i++)
                        {
                            sheet.AutoSizeColumn(i);
                        }
                        ((HSSFSheet)hssfworkbook.GetSheetAt(0)).AlternativeFormula = false;
                        ((HSSFSheet)hssfworkbook.GetSheetAt(0)).AlternativeExpression = false;

                        if (!File.Exists(path))
                        {
                            using (file = new FileStream(path, FileMode.Create))
                            {
                                hssfworkbook.Write(file);
                            }
                        }
                        else
                        {
                            using (file = File.Open(path, FileMode.Append, FileAccess.Write))
                            {
                                hssfworkbook.Write(file);
                            }
                        }

                        //hssfworkbook.Close();

                        //Update Last time
                        var last = timeStamps.Max();
                        (new Model.LastTimeUpdateRepository()).UpdateLastTime(new DateTime(last.Year, last.Month, last.Day, last.Hour, 0, 0).AddHours(1));

                        return "File " + filename + " was updated Data from " + timeStamps.Min().ToString("dd/MM/yyyy HH:mm:ss") +
                            " to " + timeStamps.Max().ToString("dd/MM/yyyy HH:mm:ss");
                    }
                }
                return "No Data was updated";
            }
            catch (Exception ex)
            {
                return "Update excel failed. " + ex.ToString();
            }
        }
    }
}
