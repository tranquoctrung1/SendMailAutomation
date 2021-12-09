using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WF_SendMailAutomation.Model
{
    public class LastTimeUpdateRepository
    {
        string connectionString = ConfigurationManager.ConnectionStrings["dtb"].ConnectionString;
        public DateTime? GetLastTimeUpdate()
        {
            DateTime? LastTimeUpdate = null;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand("select top 1 * from LastTimeUpdate", connection);

                SqlDataReader rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    if (rd["LastTimeUpdate"] != DBNull.Value)
                        LastTimeUpdate = (DateTime)rd["LastTimeUpdate"];
                }
            }
            return LastTimeUpdate;
        }

        public void UpdateLastTime(DateTime lastTime)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand("Update LastTimeUpdate set LastTimeUpdate = '" + lastTime.ToString("yyyy-MM-dd HH:mm:ss") +"'", connection);
                SqlDataReader rd = cmd.ExecuteReader();
            }
        }
    }
}
