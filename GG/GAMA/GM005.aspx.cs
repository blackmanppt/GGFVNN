using GG.ReferenceCode;
using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GG.GAMA
{
    public partial class GM005 : System.Web.UI.Page
    {
        //字串處理 切字串 = new 字串處理();
        static string strConnectString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["GGFConnectionString"].ToString();
        SysLog Log = new SysLog();
        string StrPageName = "Overtime Download", StrProgram = "GM005.aspx";
        protected void Page_PreInit(object sender, EventArgs e)
        {
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            #region 網頁Layout基本參數
            //網頁標題

            BrandLB.Text = StrPageName;
            Page.Title = StrPageName;
            //StrError名稱 = "";
            //StrProgram = "TempCode2.aspx";
            DateRangeTB.Attributes["readonly"] = "readonly";
            #endregion
        }
        protected void DbInit()
        {
            DataTable dt = new DataTable();
            using (SqlConnection Conn = new SqlConnection(strConnectString))
            {
                SqlDataAdapter myAdapter = new SqlDataAdapter(selectsql().ToString(), Conn);
                myAdapter.Fill(dt);    //---- 這時候執行SQL指令。取出資料，放進 DataSet。

            }
            #region query 使用 In
            //        using (SqlConnection conn1 = new SqlConnection(strConnectString))
            //        {
            //            SqlCommand command1 = conn1.CreateCommand();
            //            SqlTransaction transaction1;
            //            conn1.Open();
            //            transaction1 = conn1.BeginTransaction("createExcelImport");
            //            try
            //            {
            //                command1.Connection = conn1;
            //                command1.Transaction = transaction1;

            //                #region 匯入
            //                string[] parameters = SamArray.Select((s, i) => "@sam_nbr" + i.ToString()).ToArray();
            //                command1.CommandText = string.Format(@"SELECT          a.receipt_date AS 發版日期, 
            //                    b.cus_name AS 客戶名稱, a.cus_style_no AS 款號, 
            //                    a.sam_nbr AS 打樣單號, dbo.F_DateToNull(a.samc_fin_date) AS 打版完成日期, 
            //                    a.samc_remark60 AS 備註, a.plan_fin_date AS 預計完成日, 
            //                    a.online_date AS 上線日期, a.samc_plan_fin_date AS 打版預計完成日, 
            //                    a.plan_fin_date AS 預計完日, dbo.F_DateToNull(a.last_date) AS 需求日
            //                    ,dbo.F_DateToNull(a.samc_fin_date) 打版完成日
            //from samc_reqm a left join bas_cus_master b on a.site=b.site and a.cus_id=b.cus_id
            //                     where sam_nbr in ( {0} ) and a.site='GGF'
            //                     ", string.Join(",", parameters));
            //                command1.Parameters.Add("@samc_fin_date", SqlDbType.DateTime).Value = DateRangeTB.Text;
            //                for (int i = 0; i < SamArray.Length; i++)
            //                    command1.Parameters.AddWithValue(parameters[i], SamArray[i]);
            //                //command1.Parameters.Add("@sam_nbr", SqlDbType.DateTime).Value = DateRangeTB.Text;
            //                command1.ExecuteNonQuery();
            //                SqlDataReader dr = command1.ExecuteReader(CommandBehavior.CloseConnection);
            //                dt.Load(dr);
            //                #endregion
            //                //transaction1.Commit();
            //            }
            //            catch (Exception ex)
            //            {
            //                Log.ErrorLog(ex, "上傳失敗", StrProgram);
            //                transaction1.Rollback();
            //                throw;
            //            }
            //            finally
            //            {
            //                conn1.Close();
            //                transaction1.Dispose();
            //            }
            //        }
            #endregion
            if (dt.Rows.Count > 0)
            {
                ReportViewer1.Visible = true;
                ReportViewer1.ProcessingMode = ProcessingMode.Local;
                ReportDataSource source = new ReportDataSource("加班單匯出", dt);
                ReportViewer1.LocalReport.DataSources.Clear();
                ReportViewer1.LocalReport.DataSources.Add(source);
                ReportViewer1.DataBind();
                ReportViewer1.LocalReport.Refresh();
            }
            else
                F_ErrorShow("搜尋不到資料");
        }

        private StringBuilder selectsql()
        {

            StringBuilder strsql = new StringBuilder(string.Format(@" select * from View加班單 where [日期] between '{0}' and '{1}' ", DateRangeTB.Text,Convert.ToDateTime(DateRangeTB.Text).AddDays(7).ToString("yyyy-MM-dd")));
            //if (!string.IsNullOrEmpty(年度DDL.SelectedValue))
            //    strsql.AppendFormat(" and upper([季節年度])  = '{0}' ", 年度DDL.SelectedValue.ToUpper());
            //if (!string.IsNullOrEmpty(季節DDL.SelectedValue))
            //    strsql.AppendFormat(" and upper([季節])  = '{0}' ", 季節DDL.SelectedValue.ToUpper());
            //if (!string.IsNullOrEmpty(款號TB.Text))
            //    strsql.AppendFormat(" and upper([Style])  like '%{0}%' ", 款號TB.Text.ToUpper());
            //if (!string.IsNullOrEmpty(品牌TB.Text))
            //    strsql.AppendFormat(" and upper([品牌])  = '{0}' ", 品牌TB.Text.ToUpper());
            //if (!string.IsNullOrEmpty(代理商TB.Text))
            //    strsql.AppendFormat(" and upper([代理商])  = '{0}' ", 代理商TB.Text.ToUpper());
            //if (主料CB.Checked)
            //    strsql.Append(" and upper([主副料])  = 'M' ");
            //if (入庫CB.Checked)
            //    strsql.Append(" and upper([採購單狀態])  = 'IN' ");
            return strsql;
        }
        public bool SearchCheck()
        {
            bool bCheck = false;
            //if (!string.IsNullOrEmpty(年度DDL.SelectedValue))
            //    bCheck = true;
            //if (!string.IsNullOrEmpty(季節DDL.SelectedValue))
            //    bCheck = true;
            //if (!string.IsNullOrEmpty(款號TB.Text))
            //    bCheck = true;
            //if (!string.IsNullOrEmpty(品牌TB.Text))
            //    bCheck = true;
            //if (!string.IsNullOrEmpty(代理商TB.Text))
            //    bCheck = true;
            return bCheck;

        }

        protected void SearchBT_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(DateRangeTB.Text.Trim()))
            {
                F_ErrorShow("Please Select day");
            }
            else
            {
                DbInit();
            }
        }

        public void F_ErrorShow(string strError)
        {
            MessageLB.Text = strError;
            AlertPanel_ModalPopupExtender.Show();
        }
    }
}