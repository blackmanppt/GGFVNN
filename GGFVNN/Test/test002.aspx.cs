using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GGFVNN.Test
{
    public partial class test002 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            // Set the ServerType to 1 for connect to the embedded server
            string connectionString =
            "User=SYSDBA;" +
            "Password=lv@0612;" +
            "Database=D:\\DB\\DB\\GSS.FDB;" +
            "DataSource=192.168.2.200;" +
            "Port=3050;" +
            "Dialect=3;" +
            "Charset=NONE;" +
            "Role=;" +
            "Connection lifetime=15;" +
            "Pooling=true;" +
            "MinPoolSize=0;" +
            "MaxPoolSize=50;" +
            "Packet Size=8192;" +
            "ServerType=0";

            FbConnection myConnection1 = new FbConnection(connectionString);
            try
            {
                // Open two connections.
                Console.WriteLine("Open two connections.");
                myConnection1.Open();
                DataTable dt = new DataTable();
                FbDataAdapter da = new FbDataAdapter("select * from PLAN_DAY where ACTUALHOURS >0", myConnection1);
                //da.SelectCommand.Parameters.Add("@id", 123);
                da.Fill(dt);   //---- 這時候執行SQL指令。取出資料，放進 DataSet。
                if (dt.Rows.Count > 0)
                {
                    ReferenceCode.ConvertToExcel xx = new ReferenceCode.ConvertToExcel();
                    xx.ExcelWithNPOI(dt, @"xlsx");
                }
                myConnection1.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            // Set the ServerType to 1 for connect to the embedded server
            string connectionString =
            "User=SYSDBA;" +
            "Password=lv@0612;" +
            "Database=D:\\DB\\DB\\GSS.FDB;" +
            "DataSource=192.168.2.200;" +
            "Port=3050;" +
            "Dialect=3;" +
            "Charset=NONE;" +
            "Role=;" +
            "Connection lifetime=15;" +
            "Pooling=true;" +
            "MinPoolSize=0;" +
            "MaxPoolSize=50;" +
            "Packet Size=8192;" +
            "ServerType=0";

            FbConnection myConnection1 = new FbConnection(connectionString);
            try
            {
                // Open two connections.
                Console.WriteLine("Open two connections.");
                myConnection1.Open();
                DataTable dt = new DataTable();
                FbDataAdapter da = new FbDataAdapter(@"select ORGANIZE_NAME as 組別,
                                                                PLAN_CODE as 款號A,
                                                                ACTUALHOURS as 有效工時,
                                                                WORKNUM	as 上班工人,
                                                                PLANQTY	as 當日目標產量,
                                                                FINISQTY as 完成量,
                                                                ERORQTY as 瑕疵數,
                                                                SAH as 有效時間,
                                                                EFF as 每小時效率
                                                             from PLAN_DAY where ACTUALHOURS >0", myConnection1);
                //da.SelectCommand.Parameters.Add("@id", 123);
                da.Fill(dt);   //---- 這時候執行SQL指令。取出資料，放進 DataSet。
                if (dt.Rows.Count > 0)
                {
                    GridView1.DataSource = dt;
                    GridView1.DataBind();
                }
                myConnection1.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}