using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WF_SendMailAutomation.Helper
{
    public class DeleteOldData
    {
        public string log = "";
        private string now = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

        public DeleteOldData()
        {
            DateTime? dateDelete = null;
            string connectionString = ConfigurationManager.ConnectionStrings["dtb"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand("p_Delete_Old_Data", connection);
                cmd.CommandType = CommandType.StoredProcedure;

                SqlDataReader rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    if (rd["maxDateTime"] != DBNull.Value)
                        dateDelete = (DateTime)rd["maxDateTime"];
                }
                connection.Close();
            }

            if (dateDelete != null)
            {
                log = now + ": Deleted data before date " + ((DateTime)dateDelete).ToString("dd/MM/yyyy HH:mm:ss") + "\r\n";
            }
            else log = now + ": No data was deleted\r\n";
            WriteLog.Write(log);
        }
    }
}
