using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PtacDealerExcelToTableService
{
    class SqlUtility
    {
        public static void exeuteStorProc(string storedProcedureName, List<KeyValuePair<string, string>> paramsData)
        {
            try
            {
                String strConnString = ConfigurationManager.ConnectionStrings["conString"].ConnectionString;

                using (SqlConnection connection = new SqlConnection(strConnString))
                {
                    connection.Open();
                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = storedProcedureName;
                    foreach (var item in paramsData)
                    {
                        cmd.Parameters.Add(item.Key, SqlDbType.Xml).Value = item.Value;
                    }
                    cmd.Connection = connection;
                    cmd.ExecuteNonQuery();
                }
            }

            catch (Exception ex)
            {
                Logger.Info(ex.Message);
            }
        }
    }
}
