using System;
using System.Web.UI;
using GG.ReferenceCode;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Data;

namespace GG.GAMA
{
    public partial class GM001 : System.Web.UI.Page
    {
        SysLog Log = new SysLog();
        //ReferenceCode.DataCheck 確認LOCK = new ReferenceCode.DataCheck();

        static string strGGMConnectString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["GGMConnectionString"].ToString();

        protected void Page_Load(object sender, EventArgs e)
        {
            SearchTB.Attributes["readonly"] = "readonly";
        }
        protected void CheckBT_Click(object sender, EventArgs e)
        {
            工時Column GetExcelDefine = new 工時Column();
            //導入匯入table
            GetExcelDefine.工時DT();
            String savePath = Server.MapPath(@"~\ExcelUpLoad\GAMA\");

            DataTable D_table = new DataTable("Excel");
            D_table = GetExcelDefine.ExcelTable.Copy();
            DataTable D_errortable = new DataTable("Error");
            //實際顯示欄位
            int Excel欄位數 = D_table.Columns.Count - 2;
            if (SearchTB.Text.Length > 0 && F_CheckData())
            {
                if (FileUpload1.HasFile)
                {
                    String fileName = FileUpload1.FileName;
                    Session["FileName"] = fileName;
                    savePath = savePath + fileName;
                    FileUpload1.SaveAs(savePath);
                    string str頁簽名稱 = "";
                    Label1.Text = "excel load success---- " + fileName;
                    //--------------------------------------------------
                    //---- （以上是）上傳 FileUpload的部分，成功運作！
                    //--------------------------------------------------




                    #region ErrorTable
                    D_errortable.Columns.Add("閱卷序號");
                    D_errortable.Columns.Add("款號");
                    D_errortable.Columns.Add("組別");
                    D_errortable.Columns.Add("日期");
                    D_errortable.Columns.Add("工號");
                    D_errortable.Columns.Add("Error");
                    #endregion


                    if (fileName.Substring(fileName.Length - 4, 4).ToUpper() == "XLSX")
                    {
                        XSSFWorkbook workbook = new XSSFWorkbook(FileUpload1.FileContent);  //==只能讀取 System.IO.Stream

                        for (int x = 0; x < workbook.NumberOfSheets; x++)
                        {
                            XSSFSheet u_sheet = (XSSFSheet)workbook.GetSheetAt(x);  //-- 0表示：第一個 worksheet工作表
                            XSSFRow headerRow = (XSSFRow)u_sheet.GetRow(3);  //-- Excel 表頭列
                            IRow DateRow = (IRow)u_sheet.GetRow(2);             //-- v.1.2.4版修改
                            Session["Date"] = SearchTB.Text;
                            str頁簽名稱 = u_sheet.SheetName.ToString();

                            //-- for迴圈的「啟始值」要加一，表示不包含 Excel表頭列
                            // for (int i = (u_sheet.FirstRowNum + 1); i <= u_sheet.LastRowNum; i++)   //-- 每一列做迴圈
                            //i=1第二列開始
                            for (int i = 1; i <= u_sheet.LastRowNum; i++)   //-- 每一列做迴圈
                            {
                                //--不包含 Excel表頭列的 "其他資料列"
                                IRow row = (IRow)u_sheet.GetRow(i);
                                F_資料確認(D_table, D_errortable, str頁簽名稱, row);
                            }
                            //-- 釋放 NPOI的資源
                            u_sheet = null;
                        }
                        //-- 釋放 NPOI的資源
                        workbook = null;
                        ////--Excel資料顯示             
                        //DataView D_View2 = new DataView(D_table);
                        //ExcelGV.DataSource = D_View2;
                        //ExcelGV.DataBind();


                        //--錯誤資料顯示
                        if (D_errortable.Rows.Count > 0)
                        {
                            DataView D_View3 = new DataView(D_errortable);
                            ErrorGV.DataSource = D_View3;
                            ErrorGV.DataBind();
                        }

                        //--------------------------------------------------
                        //---- （以下是）上傳 FileUpload的部分！
                        //--------------------------------------------------
                    }
                    else
                    {

                        HSSFWorkbook workbook = new HSSFWorkbook(FileUpload1.FileContent);  //==只能讀取 System.IO.Stream
                        for (int x = 0; x < workbook.NumberOfSheets; x++)
                        {
                            HSSFSheet u_sheet = (HSSFSheet)workbook.GetSheetAt(x);  //-- 0表示：第一個 worksheet工作表
                            HSSFRow headerRow = (HSSFRow)u_sheet.GetRow(3);  //-- Excel 表頭列
                            IRow DateRow = (IRow)u_sheet.GetRow(2);             //-- v.1.2.4版修改
                            Session["Date"] = SearchTB.Text;
                            str頁簽名稱 = u_sheet.SheetName.ToString();

                            for (int i = 1; i <= u_sheet.LastRowNum; i++)   //-- 每一列做迴圈
                            {

                                //--不包含 Excel表頭列的 "其他資料列"
                                IRow row = (IRow)u_sheet.GetRow(i);
                                F_資料確認(D_table, D_errortable, str頁簽名稱, row);
                            }
                            //-- 釋放 NPOI的資源
                            u_sheet = null;
                        }
                        //-- 釋放 NPOI的資源
                        workbook = null;
                        ////--Excel資料顯示             
                        //DataView D_View2 = new DataView(D_table);
                        //ExcelGV.DataSource = D_View2;
                        //ExcelGV.DataBind();
                        //--錯誤資料顯示
                        if (D_errortable.Rows.Count > 0)
                        {
                            DataView D_View3 = new DataView(D_errortable);
                            ErrorGV.DataSource = D_View3;
                            ErrorGV.DataBind();
                        }

                        //--------------------------------------------------
                        //---- （以下是）上傳 FileUpload的部分！
                        //--------------------------------------------------

                    }
                }
                else
                {
                    Label1.Text = "????  ...... Please check excel";
                }   // FileUpload使用的第一個 if判別式

                if (D_table.Rows.Count > 0)
                    Session["Excel"] = D_table;
                else
                {
                    Session["Excel"] = null;
                    //ExcelGV.DataSource = null;
                    //ExcelGV.DataBind();

                }

                if (D_errortable.Rows.Count > 0)
                    Session["ExcelError"] = D_errortable;
                else
                {
                    Session["ExcelError"] = null;
                    ErrorGV.DataSource = null;
                    ErrorGV.DataBind();
                }
            }
            else
            {
                if (F_CheckData() == false)
                    Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('There is already import on the day');</script>");
                else
                    Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('Please Check upday');</script>");
            }

        }

        private void F_資料確認(DataTable D_table, DataTable D_errortable, string str頁簽名稱, IRow row)
        {
            bool berror = false;
            StringBuilder sbError = new StringBuilder();
            string str閱卷序號 = "", str款號 = "", str組別 = "";
            string strRegex工號 = "(G|N)[0-9]{4}", strRegex工段 = "[0-9]{3}", strRegex數量 = "[0-9]{4}";
            //string strRegex日期 = "\\b(?<year>\\d{4})(?<month>\\d{2})(?<day>\\d{2})\\b";

            if (!string.IsNullOrEmpty(row.GetCell(0).ToString()))
            {
                if (row.GetCell(0).ToString().ToUpper() == "NULL")
                {
                    berror = 錯誤訊息(sbError, "閱卷序號資料為NULL");
                }
                else
                {
                    str閱卷序號 = row.GetCell(0).ToString().ToUpper();
                    sbError.AppendFormat("閱卷序號：{0}", str閱卷序號);
                }
            }
            if (!string.IsNullOrEmpty(row.GetCell(1).ToString()))
            {
                if (row.GetCell(1).ToString().ToUpper() == "NULL")
                {
                    berror = 錯誤訊息(sbError, "Style is NULL");
                }
                else
                {
                    str款號 = row.GetCell(1).ToString().ToUpper();
                    using (SqlConnection conn = new SqlConnection(strGGMConnectString))
                    {
                        //try
                        //{
                        //    SqlCommand command = new SqlCommand();
                        //    command.Connection = conn;
                        //    command.CommandText = @"SELECT top 1
                        //                        *
                        //                    FROM [dbo].[ordc_bah1]
                        //                    where [cus_item_no] = @cus_item_no";
                        //    command.CommandType = CommandType.Text;
                        //    command.Parameters.Add("@cus_item_no", SqlDbType.NVarChar).Value = str款號;
                        //    conn.Open();
                        //    SqlDataReader reader = command.ExecuteReader();

                        //    if (!reader.HasRows)
                        //    {
                        //        berror = 錯誤訊息(sbError, "無訂單款號資料");
                        //    }
                        //    reader.Close();
                        //}
                        //catch (Exception ex)
                        //{
                        //    berror = 錯誤訊息(sbError, "搜尋訂單款號資料異常"+ex.ToString());
                        //}
                    }
                }

            }
            else
                berror = 錯誤訊息(sbError, "no style、");

            if (!string.IsNullOrEmpty(row.GetCell(2).ToString()))
            {
                if (row.GetCell(2).ToString().ToUpper() == "NULL")
                {
                    berror = 錯誤訊息(sbError, "組別資料為NULL");
                }
                else
                    str組別 = row.GetCell(2).ToString().ToUpper();
            }
            else
                berror = 錯誤訊息(sbError, "沒有組別、");

            for (int z = 3; z < 24; z = z + 7)
            {
                string str工號 = "";
                //工號V9999 工段99 數量9999

                bool b工號Error = false;
                if (!string.IsNullOrEmpty(row.GetCell(z).ToString()))
                {
                    str工號 = row.GetCell(z).ToString().Trim().ToUpper();
                    Regex reg = new Regex(strRegex工號);
                    b工號Error = (!reg.IsMatch(str工號) && str工號.Length != 5) ? true : false;
                    //工號轉換
                    if (!b工號Error)
                        str工號 = (str工號.Substring(0, 1) == "N") ? "NGM" + str工號.Substring(1) : "GM" + str工號.Substring(1);
                }
                else
                    b工號Error = true;

                //檢查1~3工段數量
                for (int zz = 0; zz < 5; zz = zz + 2)
                {
                    string str工段 = "", str數量 = "", str工段轉換 = "";
                    bool b工段Error = false, b數量Error = false, b工段轉換Error = false;
                    DataRow D_dataRow = D_table.NewRow();
                    DataRow D_erroraRow = D_errortable.NewRow();
                    D_dataRow[0] = str頁簽名稱;
                    D_dataRow[1] = SearchTB.Text;
                    if (!string.IsNullOrEmpty(row.GetCell(z + zz + 1).ToString()))
                    {
                        //永琦工段要三碼，補0
                        str工段 = "0" + row.GetCell(z + zz + 1).ToString();
                        Regex reg = new Regex(strRegex工段);
                        b工段Error = (!reg.IsMatch(str工段) || str工段.Length != 3) ? true : false;
                    }
                    else
                        b工段Error = true;
                    //工段轉換
                    if (!b工段Error && !berror)
                    {
                        using (SqlConnection conn = new SqlConnection(strGGMConnectString))
                        {
                            try
                            {
                                SqlCommand command = new SqlCommand();
                                command.Connection = conn;
                                command.CommandText = @"SELECT top 1
                                                [STP_ID]
                                            FROM [dbo].[FILH01A]
                                            where rtrim(ltrim([SEQ_NO])) = @SEQ_NO and rtrim(ltrim([ORD_NO])) = @ORD_NO";
                                command.CommandType = CommandType.Text;
                                command.Parameters.Add("@SEQ_NO", SqlDbType.NVarChar).Value = str工段.Trim();
                                command.Parameters.Add("@ORD_NO", SqlDbType.NVarChar).Value = str款號.Trim();
                                conn.Open();
                                SqlDataReader reader = command.ExecuteReader();

                                if (reader.HasRows)
                                {
                                    reader.Read();
                                    str工段 = reader["STP_ID"].ToString();
                                }
                                else
                                {
                                    b工段轉換Error = true;
                                    str工段轉換 = "工段資料轉換失敗,沒有資料";
                                }
                                reader.Close();
                            }
                            catch (Exception ex)
                            {
                                b工段轉換Error = true;
                                str工段轉換 = " 工段資料轉換失敗,沒有資料," + ex.ToString();
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(row.GetCell(z + zz + 2).ToString()))
                    {
                        str數量 = row.GetCell(z + zz + 2).ToString();
                        Regex reg = new Regex(strRegex數量);
                        b數量Error = (!reg.IsMatch(str數量) || str數量.Length != 4) ? true : false;
                    }
                    else
                        b數量Error = true;

                    //工段2，3沒資料跳過
                    if (str數量 == "" && str工段 == "" && zz > 0)
                        continue;
                    if (str數量 == "" && str工段 == "" && str工號 == "")
                        continue;
                    //if (b工段Error || b數量Error || berror || b工號Error || b工段轉換Error)
                    if (b數量Error || berror || b工號Error )
                    {
                        if (b工號Error)
                            錯誤訊息(sbError, string.Format(" 工號錯誤：{0}", (str工號 == "") ? "沒有工號" : str工號));
                        //if (b工段Error)
                        //    錯誤訊息(sbError, string.Format(" 工段錯誤：{0}", (str工段 == "") ? "沒有工段" : str工段));
                        if (b數量Error)
                            錯誤訊息(sbError, string.Format(" 數量錯誤：{0}", (str數量 == "") ? "數量為0" : str數量));
                        //if (b工段轉換Error)
                        //    錯誤訊息(sbError, str工段轉換);
                        //str閱卷序號 = "", str款號 = "", str組別 = "", str日期 = "";
                        //D_erroraRow[0] = str頁簽名稱;
                        D_erroraRow[0] = str閱卷序號;
                        D_erroraRow[1] = str款號;
                        D_erroraRow[2] = str組別;
                        D_erroraRow[3] = SearchTB.Text;
                        D_erroraRow[4] = str工號;
                        D_erroraRow[5] = "Error Data：" + sbError;
                        D_errortable.Rows.Add(D_erroraRow);
                    }
                    else
                    {
                        //D_dataRow[0] = str頁簽名稱;
                        D_dataRow[0] = str閱卷序號;
                        D_dataRow[1] = str款號;
                        D_dataRow[2] = str組別;
                        D_dataRow[3] = SearchTB.Text;
                        D_dataRow[4] = str工號;
                        D_dataRow[5] = str工段;
                        D_dataRow[6] = str數量;
                        D_table.Rows.Add(D_dataRow);
                    }
                }
            }
        }

        private static bool 錯誤訊息(StringBuilder sbError, string strerror)
        {
            bool berror = true;
            sbError.Append(strerror);
            return berror;
        }

        //抓HeadID
        private int GetExcelIdex()
        {
            Int32 ProductivityHeadID = 0;
            string sql =
                @"INSERT INTO [dbo].[工段總表單頭]
                    (filename,日期)
                    VALUES(@FileName,@Date); 
                    SELECT CAST(scope_identity() AS int)";
            using (SqlConnection conn = new SqlConnection(strGGMConnectString))
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.Add("@Date", SqlDbType.NVarChar);
                cmd.Parameters.Add("@FileName", SqlDbType.NVarChar);
                cmd.Parameters["@Date"].Value = SearchTB.Text;
                cmd.Parameters["@FileName"].Value = Session["FileName"].ToString();
                //cmd.Parameters["@Area"].Value = strArea;
                //cmd.Parameters["@Team"].Value = strImportType;
                try
                {
                    conn.Open();
                    ProductivityHeadID = (Int32)cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    Log.ErrorLog(ex, "Get 工段總表明細 uid Error" + Session["FileName"].ToString() + ":", "GM001.aspx");
                }
            }
            return (int)ProductivityHeadID;
        }

        protected void UpLoadBT_Click(object sender, EventArgs e)
        {
            if (SearchTB.Text.Trim() != "" && F_CheckData() && Session["ExcelError"] == null)
            {
                if (Session["Excel"] != null)
                {
                    DataTable dt = (DataTable)Session["Excel"];
                    //TypeLB.Text = dt.Rows[2][2].ToString();
                    int iIndex = 0;
                    iIndex = GetExcelIdex();
                    if (iIndex > 0)
                        using (SqlConnection conn1 = new SqlConnection(strGGMConnectString))
                        {
                            SqlCommand command1 = conn1.CreateCommand();
                            SqlTransaction transaction1;
                            conn1.Open();
                            transaction1 = conn1.BeginTransaction("createExcelImport");

                            command1.Connection = conn1;
                            command1.Transaction = transaction1;
                            try
                            {
                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    //string strInsertColumn="", strInsertData="";
                                    command1.CommandText = string.Format(@"INSERT INTO [dbo].[工段總表明細]
                                                      ([uid]
                                                      ,[閱卷序號]
                                                      ,[款號]
                                                      ,[組別]
                                                      ,[日期]
                                                      ,[工號]
                                                      ,[工段]
                                                      ,[數量])
                                                 VALUES
                                                       ({0}
                                                       ,@閱卷序號
                                                       ,@款號
                                                       ,@組別
                                                       ,@日期
                                                       ,@工號
                                                       ,@工段
                                                       ,@數量
                                                        )
                                                       ", iIndex);
                                    command1.Parameters.Add("@閱卷序號", SqlDbType.NVarChar).Value = (dt.Rows[i]["閱卷序號"] == null) ? "" : dt.Rows[i]["閱卷序號"].ToString();
                                    command1.Parameters.Add("@款號", SqlDbType.NVarChar).Value = (dt.Rows[i]["款號"] == null) ? "" : dt.Rows[i]["款號"].ToString();
                                    command1.Parameters.Add("@組別", SqlDbType.NVarChar).Value = (dt.Rows[i]["組別"] == null) ? "" : dt.Rows[i]["組別"].ToString();
                                    command1.Parameters.Add("@日期", SqlDbType.NVarChar).Value = (dt.Rows[i]["日期"] == null) ? "" : dt.Rows[i]["日期"].ToString();
                                    command1.Parameters.Add("@工號", SqlDbType.NVarChar).Value = (dt.Rows[i]["工號"] == null) ? "" : dt.Rows[i]["工號"].ToString();
                                    command1.Parameters.Add("@工段", SqlDbType.NVarChar).Value = (dt.Rows[i]["工段"] == null) ? "" : dt.Rows[i]["工段"].ToString();
                                    command1.Parameters.Add("@數量", SqlDbType.NVarChar).Value = (dt.Rows[i]["數量"] == null) ? "" : dt.Rows[i]["數量"].ToString();

                                    command1.ExecuteNonQuery();
                                    command1.Parameters.Clear();
                                }
                                ////上傳成功更新Head狀態
                                //command1.CommandText = string.Format(@"UPDATE [dbo].[Productivity_Head] SET [Flag] = 1 ,[Date] = @Date WHERE uid = {0} ", iIndex);
                                //command1.Parameters.Add("@Date", SqlDbType.NVarChar).Value = Session["Date"].ToString();
                                //command1.ExecuteNonQuery();
                                transaction1.Commit();
                            }
                            catch (Exception ex1)
                            {
                                try
                                {
                                    Log.ErrorLog(ex1, "Import Excel Error :" + Session["FileName"].ToString(), "GM001.aspx");
                                }
                                catch (Exception ex2)
                                {
                                    Log.ErrorLog(ex2, "Insert Error Error:" + Session["FileName"].ToString(), "GM001.aspx");
                                }
                                finally
                                {
                                    transaction1.Rollback();
                                    Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('Upload fail');</script>");
                                }
                            }
                            finally
                            {
                                conn1.Close();
                                conn1.Dispose();
                                command1.Dispose();
                                Session.RemoveAll();
                                Label1.Text = "Upload success";
                            }
                        }
                    else
                        Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('Create fail');</script>");
                }
            }
            else
            {
                if (Session["ExcelError"] != null)
                    Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('Data error');</script>");
                else if (SearchTB.Text.Trim() != "")
                    Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('There is already import on the day');</script>");
                else
                    Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('Please Check upday');</script>");
            }
        }


        protected void DeleteBT_Click(object sender, EventArgs e)
        {
            using (SqlConnection conn1 = new SqlConnection(strGGMConnectString))
            {
                SqlCommand command1 = conn1.CreateCommand();
                SqlTransaction transaction1;
                conn1.Open();
                transaction1 = conn1.BeginTransaction("createExcelImport");

                command1.Connection = conn1;
                command1.Transaction = transaction1;
                try
                {
                    command1.CommandText = string.Format(@"UPDATE [dbo].[工段總表單頭] SET [IsDelete] = 1,[MODIFYDATE]=GETDATE()  WHERE [日期] = @Date and [IsDelete] = 0 ");
                    command1.Parameters.Add("@Date", SqlDbType.NVarChar).Value = SearchTB.Text;
                    command1.ExecuteNonQuery();
                    command1.Parameters.Clear();
                    command1.CommandText = string.Format(@"UPDATE [dbo].[工段總表明細] SET [IsDelete] = 1,[MODIFYDATE]=GETDATE()  WHERE [日期] = @Date and [IsDelete] = 0 ");
                    command1.Parameters.Add("@Date", SqlDbType.NVarChar).Value = SearchTB.Text;
                    command1.ExecuteNonQuery();

                    transaction1.Commit();
                    Label1.Text = "Data was deleted.Please Upload again.";
                }
                catch (Exception ex1)
                {
                    try
                    {
                        Log.ErrorLog(ex1, "Delete Error :", "GM001.aspx");
                    }
                    catch (Exception ex2)
                    {
                        Log.ErrorLog(ex2, "Delete Error Error:", "GM001.aspx");
                    }
                    finally
                    {
                        transaction1.Rollback();
                        Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('Deleted fail');</script>");
                    }
                }
                finally
                {
                    conn1.Close();
                    conn1.Dispose();
                    command1.Dispose();
                    Session.RemoveAll();
                }
            }

        }

        private void DataLock()
        {
            Label1.Text = "資料已鎖定，請洽管理者";
        }

        public Boolean F_CheckData()
        {
            bool bcheck = true;

            using (SqlConnection conn = new SqlConnection(strGGMConnectString))
            {
                SqlCommand command = new SqlCommand();
                command.Connection = conn;
                command.CommandText = @"SELECT top 1 [uid]
                                              ,[filename]
                                              ,[日期]
                                              ,[CREATEDATE]
                                              ,[MODIFYDATE]
                                              ,[IsDelete]
                                            FROM [dbo].[工段總表單頭]
                                            where [日期] = @Date and [IsDelete] = 0";
                command.CommandType = CommandType.Text;
                command.Parameters.Add("@Date", SqlDbType.NVarChar).Value = SearchTB.Text;
                conn.Open();
                SqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    bcheck = false;
                }
                reader.Close();
            }
            return bcheck;
        }
    }
}