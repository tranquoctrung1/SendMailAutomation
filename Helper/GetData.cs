using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WF_SendMailAutomation
{
    public class GetData
    {
        string connectionString = ConfigurationManager.ConnectionStrings["dtb"].ConnectionString;

        public List<DataViewModel> lst = new List<DataViewModel>();

        public GetData(DateTime start, DateTime end)
        {
            try
            {
                //var now = DateTime.Now;
                //var startThisMonth = new DateTime(now.Year, now.Month, 1);
                //var startLastMonth = startThisMonth.AddMonths(-1);
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand("p_Get_Data_Custom_Report", connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(
                        new SqlParameter("@interval", 60));
                    cmd.Parameters.Add(
                        new SqlParameter("@start", start));
                    cmd.Parameters.Add(
                        new SqlParameter("@end", end));

                    SqlDataReader rd = cmd.ExecuteReader();
                    while (rd.Read())
                    {
                        DataViewModel s = new DataViewModel();

                        if (rd["DateAndTime"] != DBNull.Value)
                            s.DateAndTime = (DateTime)rd["DateAndTime"];
                        s.TagName = (string)rd["TagName"].ToString();
                        if (rd["Val"] != DBNull.Value)
                            s.Val = (double?)rd["Val"];
                        lst.Add(s);
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
