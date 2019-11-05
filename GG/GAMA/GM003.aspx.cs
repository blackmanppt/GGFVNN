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
using System.Linq;

namespace GG.GAMA
{
    public partial class GM003 : System.Web.UI.Page
    {
        SysLog Log = new SysLog();
        //ReferenceCode.DataCheck 確認LOCK = new ReferenceCode.DataCheck();


        static string strGGMConnectString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["GGFConnectionString"].ToString();

        protected void Page_Load(object sender, EventArgs e)
        {
            SearchTB.Attributes["readonly"] = "readonly";
        }
        protected void CheckBT_Click(object sender, EventArgs e)
        {
            工時Column GetExcelDefine = new 工時Column();
            //導入匯入table
            GetExcelDefine.加班DT();
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
                    D_errortable.Columns.Add("Line");
                    D_errortable.Columns.Add("Name");
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
                DataTable DtCount = new DataTable();
                if (D_table.Rows.Count > 0)
                {
                    Session["Excel"] = D_table;
                    
                    DtCount.Columns.Add("Error");
                    foreach (var row in D_table.AsEnumerable().GroupBy(p => new {
                        Group=p.Field<string>("閱卷序號"),
                        Line = p.Field<string>("組別"),
                        ID = p.Field<string>("工號")
                    }).Select(x => new { ID = x.Key, Count = x.Count() }))
                    {
                        {
                            if (row.Count > 1)
                            {
                                DataRow dr = DtCount.NewRow();
                                dr[0] = string.Format("{0} count: {1} ", row.ID, row.Count);
                                DtCount.Rows.Add(dr);
                            }
                        }
                    }
                    if(DtCount.Rows.Count>0)
                    {
                        Session["Error2"] = DtCount;
                        ErrorGV2.DataSource = DtCount;
                        ErrorGV2.DataBind();
                    }
                    
                }
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
            bool berror = false, b工時Error = false;
            StringBuilder sbError = new StringBuilder();

            string str閱卷序號 = "", str組別 = "", str工時 = "", strOther = "", str夜班="";
            string strRegex工號 = "(G|N|T)[0-9]{4}", strRegex工時 = "^[0-9]*$";
            float f工時 = 0;
            //string strRegex日期 = "\\b(?<year>\\d{4})(?<month>\\d{2})(?<day>\\d{2})\\b";
            Session["ExcelError"] = null;
            Session["Error2"] = null;
            try
            {
                if (!string.IsNullOrEmpty(row.GetCell(0).ToString().Trim()))
                {
                    if (row.GetCell(0).ToString().ToUpper() == "NULL")
                    {
                        berror = 錯誤訊息(sbError, "閱卷序號資料為NULL");
                    }
                    else
                    {
                        str閱卷序號 = row.GetCell(0).ToString().ToUpper();
                        //sbError.AppendFormat("閱卷序號：{0}", str閱卷序號);
                    }
                }

                if (!string.IsNullOrEmpty(row.GetCell(1).ToString().Trim()))
                {
                    if (row.GetCell(1).ToString().ToUpper() == "NULL")
                    {
                        berror = 錯誤訊息(sbError, "Line is NULL");
                    }
                    else
                    {
                        str組別 = row.GetCell(1).ToString().Trim().ToUpper();
                        switch (str組別)
                        {
                            case "ST":
                            case "SAM":
                            case "SF":
                            case "CL":
                            case "CU":
                            case "PA":
                            case "PR":
                            case "MA":
                            case "ME":
                            case "WA":
                            case "OT":
                            case "QC":
                            case "QA":
                            case "IE":
                            case "PL":
                            case "CK":
                                break;
                            default:
                                if (str組別.Substring(0, 1) == "G" && str組別.Length == 3)
                                {
                                    string str_i = str組別.Substring(1);
                                    int i = 0;
                                    try
                                    {
                                        i = int.Parse(str_i);
                                        if (i <= 0 || i > 30)
                                            berror = 錯誤訊息(sbError, "Line Error");
                                    }
                                    catch (Exception)
                                    {
                                        berror = 錯誤訊息(sbError, "Line Error");
                                    }
                                }
                                else if(str組別.Substring(0, 1) == "N" && str組別.Length == 3)
                                {
                                    string str_i = str組別.Substring(1);
                                    int i = 0;
                                    try
                                    {
                                        i = int.Parse(str_i);
                                        if (i < 29 || i > 40)
                                            berror = 錯誤訊息(sbError, "Line Error");
                                    }
                                    catch (Exception)
                                    {
                                        berror = 錯誤訊息(sbError, "Line Error");
                                    }
                                }
                                else
                                    berror = 錯誤訊息(sbError, "Line Error");
                                break;
                        }
                    }
                }
                else
                    berror = 錯誤訊息(sbError, "no Line、");

                if (!string.IsNullOrEmpty(row.GetCell(3).ToString().Trim()))
                {
                    if (row.GetCell(3).ToString().ToUpper() == "NULL")
                    {
                        berror = 錯誤訊息(sbError, "Hour is NULL");
                    }
                    else
                    {
                        str工時 = row.GetCell(3).ToString().Trim();
                        Regex reg = new Regex(strRegex工時);
                        if(!reg.IsMatch(str工時))
                            berror = 錯誤訊息(sbError, "Hour is Error");
                        else
                            try
                            {
                                f工時 = float.Parse(str工時);
                            }
                            catch (Exception)
                            {
                                berror = 錯誤訊息(sbError, "Convert hour Error");
                            }
                    }
                }


                if (!string.IsNullOrEmpty(row.GetCell(4).ToString().Trim()))
                {
                    if (row.GetCell(4).ToString().Trim() == "0.5" || row.GetCell(4).ToString().Trim() == ".5")
                        f工時 += 0.5f;
                    else
                        berror = 錯誤訊息(sbError, "half-Hour is Error , Please check data equal 0.5 ");
                }
                else
                {
                    //不會有0.5小時的加班
                    if (string.IsNullOrEmpty(row.GetCell(3).ToString().Trim()))
                    {
                        b工時Error = true;
                    }
                }
                try
                {
                    if (!string.IsNullOrEmpty(row.GetCell(18).ToString().Trim()))
                    {
                        str夜班 = row.GetCell(18).ToString().Trim();
                    }
                }
                catch (Exception)
                {

                    
                }
                
                if(row.Cells.Count>5)
                    for (int z = 6; z < row.Cells.Count-1; z++)
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
                            {
                                switch (str工號.Substring(0, 1))
                                {
                                    case "N":
                                        str工號 = "NGM" + str工號.Substring(1);
                                        break;
                                    case "M":
                                        str工號 = "GM" + str工號.Substring(1);
                                        break;
                                    default:
                                        str工號 = "TGM" + str工號.Substring(1);
                                        break;
                                }
                            }
                        }
                        else
                            continue;


                        var D_dataRow = D_table.NewRow();
                        DataRow D_erroraRow = D_errortable.NewRow();
                        D_dataRow[0] = str頁簽名稱;
                        D_dataRow[1] = SearchTB.Text;

                        if (b工時Error || berror || b工號Error)
                        {
                            string strerror = "";
                            if (b工號Error)
                                strerror = string.Format(@" Name{1} Error： {0} ", (str工號 == "") ? "No name" : str工號, z - 5);
                            if (b工時Error)
                                strerror = string.Format(" Hour Error：{0}", (str工時 == "") ? "hour is empty" : str工時);
                            D_erroraRow[0] = str閱卷序號;
                            D_erroraRow[1] = str組別;
                            D_erroraRow[2] = str工號;
                            D_erroraRow[3] = "Error Data：" + sbError + strerror;
                            D_errortable.Rows.Add(D_erroraRow);
                        }
                        else
                        {
                            D_dataRow[0] = str閱卷序號;
                            D_dataRow[1] = str組別;
                            D_dataRow[2] = SearchTB.Text;
                            D_dataRow[3] = f工時;
                            D_dataRow[4] = strOther;
                            D_dataRow[5] = str工號;
                            D_dataRow[6] = str夜班;
                            D_table.Rows.Add(D_dataRow);
                        }
                    }
                    else
                    {
                        DataRow D_erroraRow = D_errortable.NewRow();
                        D_erroraRow[0] = str閱卷序號;
                        D_erroraRow[1] = str組別;
                    
                        D_erroraRow[3] = "Error Data：No Name" ;
                        D_errortable.Rows.Add(D_erroraRow);
                    }
            }
            catch (Exception ex)
            {
                string xxx = ex.ToString();
                
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
            Int32 馬島加班單HeadID = 0;
            string sql =
                @"INSERT INTO [dbo].[馬島加班單Head]
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
                    馬島加班單HeadID = (Int32)cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    Log.ErrorLog(ex, "Get 馬島加班單Head uid Error" + Session["FileName"].ToString() + ":", "GM001.aspx");
                }
            }
            return (int)馬島加班單HeadID;
        }

        protected void UpLoadBT_Click(object sender, EventArgs e)
        {
            if (SearchTB.Text.Trim() != "" && F_CheckData() && Session["ExcelError"] == null && Session["Error2"]==null)
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
                                    command1.CommandText = string.Format(@"INSERT INTO [dbo].[馬島加班單Line]
                                                      ([uid]
                                                      ,[閱卷序號]
                                                      ,[組別]
                                                      ,[日期]
                                                      ,[工號]
                                                      ,[時數]
                                                      ,[other]
                                                      ,[夜班])
                                                 VALUES
                                                       ({0}
                                                       ,@閱卷序號
                                                       ,@組別
                                                       ,@日期
                                                       ,@工號
                                                       ,@時數
                                                       ,@other
                                                       ,@夜班
                                                        )
                                                       ", iIndex);
                                    command1.Parameters.Add("@閱卷序號", SqlDbType.NVarChar).Value = (dt.Rows[i]["閱卷序號"] == null) ? "" : dt.Rows[i]["閱卷序號"].ToString();
                                    command1.Parameters.Add("@組別", SqlDbType.NVarChar).Value = (dt.Rows[i]["組別"] == null) ? "" : dt.Rows[i]["組別"].ToString();
                                    command1.Parameters.Add("@日期", SqlDbType.SmallDateTime).Value = (dt.Rows[i]["日期"] == null) ? "" : dt.Rows[i]["日期"].ToString();
                                    command1.Parameters.Add("@工號", SqlDbType.NVarChar).Value = (dt.Rows[i]["工號"] == null) ? "" : dt.Rows[i]["工號"].ToString();
                                    command1.Parameters.Add("@時數", SqlDbType.Float).Value = (dt.Rows[i]["工時"] == null) ? "" : dt.Rows[i]["工時"].ToString();
                                    command1.Parameters.Add("@other", SqlDbType.NVarChar).Value = (string.IsNullOrEmpty(dt.Rows[i]["例外"].ToString())) ? "" : dt.Rows[i]["例外"].ToString();
                                    command1.Parameters.Add("@夜班", SqlDbType.Bit).Value = (string.IsNullOrEmpty(dt.Rows[i]["夜班"].ToString())) ? "false": "true" ;

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
                                    Log.ErrorLog(ex1, "Import Excel Error :" + Session["FileName"].ToString(), "GM003.aspx");
                                }
                                catch (Exception ex2)
                                {
                                    Log.ErrorLog(ex2, "Insert Error Error:" + Session["FileName"].ToString(), "GM003.aspx");
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
                if (Session["ExcelError"] != null || Session["Error2"] != null)
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
                transaction1 = conn1.BeginTransaction("createExcelImportG3");

                command1.Connection = conn1;
                command1.Transaction = transaction1;
                try
                {
                    command1.CommandText = string.Format(@"UPDATE [dbo].[馬島加班單Head] SET [是否刪除] = 1,[修改日期]=GETDATE()  WHERE [日期] = @Date and [是否刪除] = 0 ");
                    command1.Parameters.Add("@Date", SqlDbType.NVarChar).Value = SearchTB.Text;
                    command1.ExecuteNonQuery();
                    command1.Parameters.Clear();
                    command1.CommandText = string.Format(@"UPDATE [dbo].[馬島加班單Line] SET [是否刪除] = 1,[MODIFYDATE]=GETDATE()  WHERE [日期] = @Date and [是否刪除] = 0 ");
                    command1.Parameters.Add("@Date", SqlDbType.NVarChar).Value = SearchTB.Text;
                    command1.ExecuteNonQuery();

                    transaction1.Commit();
                    Label1.Text = "Data was deleted.Please Upload again.";
                }
                catch (Exception ex1)
                {
                    try
                    {
                        Log.ErrorLog(ex1, "Delete Error :", "GM003.aspx");
                    }
                    catch (Exception ex2)
                    {
                        Log.ErrorLog(ex2, "Delete Error Error:", "GM003.aspx");
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
                command.CommandText = @"SELECT top 1 *
                                            FROM [dbo].[馬島加班單Head]
                                            where [日期] = @Date and [是否刪除] = 0";
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