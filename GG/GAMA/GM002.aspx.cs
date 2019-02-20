using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace GG.GAMA
{
    public partial class GM002 : System.Web.UI.Page
    {
        static string strConnectString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["GGMConnectionString"].ToString();
        protected void Page_Load(object sender, EventArgs e)
        {
            StartTB.Attributes["readonly"] = "readonly";
            EndTB.Attributes["readonly"] = "readonly";
        }
        protected void ClearBT_Click(object sender, EventArgs e)
        {
            StartTB.Text = "";
            EndTB.Text = "";
        }

        protected void SearchBT_Click(object sender, EventArgs e)
        {
            MessageLT.Text = "";
            DbInit();
        }
        protected void DbInit()
        {

            DataTable dt = new DataTable();
            using (SqlConnection Conn = new SqlConnection(strConnectString))
            {
                SqlDataAdapter myAdapter = new SqlDataAdapter(selectsql().ToString(), Conn);
                myAdapter.Fill(dt);    //---- 這時候執行SQL指令。取出資料，放進 DataSet。

            }
            if (dt.Rows.Count > 0)
            {
                ReportViewer1.Visible = true;
                ReportViewer1.ProcessingMode = ProcessingMode.Local;
                ReportDataSource source = new ReportDataSource("工段匯入", dt);
                ReportViewer1.LocalReport.DataSources.Clear();
                ReportViewer1.LocalReport.DataSources.Add(source);
                ReportViewer1.DataBind();
                ReportViewer1.LocalReport.Refresh();
            }
            else
            {
                MessageLT.Text = @"                            
                                    <div class='form-group'>
                                        <h3 class='text-info text-center'>沒有資料 </ h3 >
                                    </div>";
            }
        }

        private StringBuilder selectsql()
        {

            StringBuilder strsql = new StringBuilder(@" 
                                        SELECT [uid]
                                              ,[id]
                                              ,[閱卷序號]
                                              ,[款號]
                                              ,[組別]
                                              ,[日期]
                                              ,[工號]
                                              ,[工段]
                                              ,[數量]
                                              ,[CREATEDATE]
                                              ,[MODIFYDATE]
                                              ,[IsDelete]
                                          FROM [dbo].[工段總表明細]
                                        ");
            strsql.AppendFormat(" where  日期 between '{0}' and '{1}'  and  [IsDelete] =0", (!String.IsNullOrEmpty(StartTB.Text)) ? StartTB.Text : "2000/01/01", (!String.IsNullOrEmpty(EndTB.Text)) ? EndTB.Text : "2999/01/01");
            return strsql;
        }
    }
}