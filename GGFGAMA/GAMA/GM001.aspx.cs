using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace GGFGAMA.GAMA
{

    public partial class GM001 : System.Web.UI.Page
    {
        static string strConnectString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["GGMConnectionString"].ToString();
        protected void Page_Load(object sender, EventArgs e)
        {
            //if (Session["username"] == null)
            //    Response.Redirect("~/LoginIndex.aspx");
            //else if (Session["username"].ToString() != "VNN")
            //    Response.Redirect("~/LoginIndex.aspx");
            StartTB.Attributes["readonly"] = "readonly";
            EndTB.Attributes["readonly"] = "readonly";
            //if (YearDDL.Items.Count == 0)
            //{
            //    //int iCountYear = DateTime.Now.Year - 2015;
            //    DateTime dtNow = DateTime.Now;
            //    //dtNow = DateTime.Parse("2020-12-01"); //測試用
            //    int iCountMonth = (DateTime.Now.Year - 2015);


            //    for (int i = 1; i < iCountMonth; i++)
            //    {
            //        if (i == 1)
            //        {
            //            YearDDL.Items.Add("");
            //        }
            //        YearDDL.Items.Add(DateTime.Now.AddMonths(-i).ToString("yyyy"));
            //    }
            //}
            //StartDay.Attributes["readonly"] = "readonly";
            //EndDay.Attributes["readonly"] = "readonly";
        }

        protected void ClearBT_Click(object sender, EventArgs e)
        {
            //YearDDL.SelectedValue = "";
            //MonthDDL.SelectedValue = "";
            //AreaDDL.SelectedValue = "";
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
            //if (!String.IsNullOrEmpty(StartTB.Text) && !String.IsNullOrEmpty(EndTB.Text))
            //{
            //    if (DateTime.Parse( StartTB.Text)> DateTime.Parse(EndTB.Text))
            //    {
            //        MessageLT.Text = @"                            
            //                        <div class='form-group'>
            //                            <h3 class='text-info text-center'>出貨日期錯誤 </ h3 >
            //                        </div>";
            //    }
            //    else
            //    {
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
            //    }
            //}
            //else
            //{
                
            //}

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
  FROM [GGF].[dbo].[工段總表明細]
                                        ");
            strsql.AppendFormat(" where  日期 between '{0}' and '{1}'  and  [IsDelete] =0", (!String.IsNullOrEmpty(StartTB.Text))?StartTB.Text:"2000/01/01", (!String.IsNullOrEmpty(EndTB.Text)) ? EndTB.Text : "2999/01/01");
            return strsql;
        }
        
    }
}