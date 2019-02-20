using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace GAMAWEB.ReferenceCode
{
    public class SysLog
    {
        static string strConnectString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["GGMConnectionString"].ToString();
        public void ErrorLog(Exception ex, string strFunction,string strProgram)
        {
            using (SqlConnection conn1 = new SqlConnection(strConnectString))
            {
                SqlCommand command1 = conn1.CreateCommand();
                SqlTransaction transaction1;
                conn1.Open();
                transaction1 = conn1.BeginTransaction("createOrder");

                command1.Connection = conn1;
                command1.Transaction = transaction1;

                command1.CommandText = "INSERT INTO [dbo].[ERROR_LOG]([ERROR_LOG],[ERROR_PROGRAM],[ERROR_FUNCTION],[CREATED_BY]) " +
                    "VALUES (@ERROR_LOG,@ERROR_PROGRAM,@ERROR_FUNCTION,@CREATED_BY) ";
                command1.Parameters.Add("@ERROR_LOG", SqlDbType.Text).Value = "'Error1:" + ex.Message;
                command1.Parameters.Add("@ERROR_PROGRAM", SqlDbType.NVarChar).Value = strProgram;
                command1.Parameters.Add("@ERROR_FUNCTION", SqlDbType.NVarChar).Value = strFunction;
                command1.Parameters.Add("@CREATED_BY", SqlDbType.NVarChar).Value = "Program";

                command1.ExecuteNonQuery();
                command1.Parameters.Clear();

                transaction1.Commit();
                conn1.Close();
                conn1.Dispose();
                command1.Dispose();

            }
        }
        public void ErrorLog(string ex, string strFunction, string strProgram)
        {
            using (SqlConnection conn1 = new SqlConnection(strConnectString))
            {
                SqlCommand command1 = conn1.CreateCommand();
                SqlTransaction transaction1;
                conn1.Open();
                transaction1 = conn1.BeginTransaction("createOrder1");

                command1.Connection = conn1;
                command1.Transaction = transaction1;

                command1.CommandText = "INSERT INTO [dbo].[ERROR_LOG]([ERROR_LOG],[ERROR_PROGRAM],[ERROR_FUNCTION],[CREATED_BY]) " +
                    "VALUES (@ERROR_LOG,@ERROR_PROGRAM,@ERROR_FUNCTION,@CREATED_BY) ";
                command1.Parameters.Add("@ERROR_LOG", SqlDbType.Text).Value = "'Error1:" + ex;
                command1.Parameters.Add("@ERROR_PROGRAM", SqlDbType.NVarChar).Value = strProgram;
                command1.Parameters.Add("@ERROR_FUNCTION", SqlDbType.NVarChar).Value = strFunction;
                command1.Parameters.Add("@CREATED_BY", SqlDbType.NVarChar).Value = "Program";

                command1.ExecuteNonQuery();
                command1.Parameters.Clear();

                transaction1.Commit();
                conn1.Close();
                conn1.Dispose();
                command1.Dispose();

            }
        }
    }
}