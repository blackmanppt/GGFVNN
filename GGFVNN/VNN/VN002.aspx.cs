using System;
using System.Data;
using System.Web.UI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Data.SqlClient;
using System.Web.Configuration;
using System.Net;
using GGFVNN.ReferenceCode;
using System.Text;
using System.Text.RegularExpressions;

namespace GGFVNN.VNN
{
    public partial class VN002 : System.Web.UI.Page
    {
        SysLog Log = new SysLog();
        //ReferenceCode.DataCheck 確認LOCK = new ReferenceCode.DataCheck();
        
        static string strConnectString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["DBConnectionString"].ToString();
        static string strConnectString1 = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["GGFConnectionString"].ToString();
        //static string strImportType = "Package";
        string strArea = "";
        //string strImportType = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            SearchTB.Attributes["readonly"] = "readonly";
            strArea = (Request.Params["AREA"] != null)? Request.Params["AREA"] :"";
            //strImportType = (Request.Params["TYPE"] != null) ? Request.Params["TYPE"] : "";
            if (strArea == "" )
                Response.Redirect("VNindex.aspx");
            AreaLB.Text = strArea;
            
        }

        protected void TeamCodeBT_Click(object sender, EventArgs e)
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["GGFConnectionString1"].ConnectionString.ToString());
            SqlDataAdapter myAdapter = new SqlDataAdapter(@"SELECT [dept_no] ,[dept_name] ,remark as Old_dept_no FROM [dbo].[GGB_dept] 
                                                    where dept_no like 'A%' or  dept_no like 'C%' or  dept_no like 'D%' 
                                                    or dept_no like 'R%' or  dept_no like 'V%'", Conn);
            DataSet ds = new DataSet();
            myAdapter.Fill(ds, "GGB_dept");    //---- 這時候執行SQL指令。取出資料，放進 DataSet。
            ReferenceCode.ConvertToExcel xx = new ReferenceCode.ConvertToExcel();
            xx.ExcelWithNPOI(ds.Tables["GGB_dept"], @"xlsx");
        }
        protected void CheckBT_Click(object sender, EventArgs e)
        {

                
            ReferenceCode.工時Column GetExcelDefine = new ReferenceCode.工時Column();

            String savePath = Server.MapPath(@"~\ExcelUpLoad\VN\");

            DataTable D_table = new DataTable("Excel");
            D_table = GetExcelDefine.ExcelTable.Copy();
            DataTable D_errortable = new DataTable("Error");
            //實際顯示欄位
            int Excel欄位數 = D_table.Columns.Count-2;
            if (SearchTB.Text.Length > 0 && F_CheckData())
            {
                if (FileUpload1.HasFile)
                {
                    String fileName = FileUpload1.FileName;
                    Session["FileName"] = fileName;
                    savePath = savePath + fileName;
                    FileUpload1.SaveAs(savePath);

                    Label1.Text = "Kiểm tra tệp tin dữ liệu thành công, tên tệp tin---- " + fileName;
                    //--------------------------------------------------
                    //---- （以上是）上傳 FileUpload的部分，成功運作！
                    //--------------------------------------------------

                       
                    #region ErrorTable
                    D_errortable.Columns.Add("SheetName");
                    D_errortable.Columns.Add("Dept");
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

                            //-- for迴圈的「啟始值」要加一，表示不包含 Excel表頭列
                            // for (int i = (u_sheet.FirstRowNum + 1); i <= u_sheet.LastRowNum; i++)   //-- 每一列做迴圈
                            //i=1第二列開始
                            for (int i = 1; i <= u_sheet.LastRowNum; i++)   //-- 每一列做迴圈
                            {
                                //--不包含 Excel表頭列的 "其他資料列"
                                IRow row = (IRow)u_sheet.GetRow(i);
                                DataRow D_dataRow = D_table.NewRow();
                                DataRow D_erroraRow = D_errortable.NewRow();
                                D_dataRow[0] = u_sheet.SheetName.ToString();
                                D_dataRow[1] = SearchTB.Text;
                                bool bcheck = true, berror = false;
                                string sError = "";
                                StringBuilder sbError = new StringBuilder();
                                string str閱卷序號 = "", str款號 = "", str組別 = "", str日期 = "";
                                #region 舊的驗證
                                ////for (int j = row.FirstCellNum; j < row.LastCellNum; j++)   //-- 每一個欄位做迴圈
                                //for (int j = row.FirstCellNum; j < Excel欄位數; j++)   //有些EXCEL會沒填資料
                                //{

                                ////沒有Style就不抓
                                //if (row.GetCell(1) == null)
                                //{
                                //    bcheck = false;
                                //    break;
                                //}

                                //switch (GetExcelDefine.VNExcel[j + 2].ColumnType)
                                //{
                                //    // Type 1：數字 , Type 2：String , Type 3：日期  , Type 6：不需要資料 String, Type 7：不需要資料 int
                                //    case 1:
                                //        DataCheck.IntData(row, D_dataRow, ref berror, ref sError, i, j, 1, GetExcelDefine.VNExcel[j + 2].ChineseName);
                                //        break;
                                //    case 2:
                                //        DataCheck.StringData(row, D_dataRow, ref berror, ref sError, i, j, 1, GetExcelDefine.VNExcel[j + 2].ChineseName);
                                //        break;
                                //    case 3:
                                //        try
                                //        {
                                //            D_dataRow[j + 2] = (string.IsNullOrEmpty(row.GetCell(j).ToString())) ? "" : row.GetCell(j).DateCellValue.ToString("yyyyMMdd");
                                //            //轉換日期格式
                                //        }
                                //        catch 
                                //        {
                                //            //D_dataRow[j] = row.GetCell(j).CellFormula.ToString();
                                //            berror = true;
                                //            sError += "閱卷序號：" + (i - 4).ToString() + "行、" + GetExcelDefine.VNExcel[j + 2].ChineseName + "日期格式錯誤。";
                                //            D_dataRow[j + 2] = (row.GetCell(j) == null) ? "" : row.GetCell(j).ToString();  //--每一個欄位，都加入同一列 DataRow
                                //                                                                                            //throw;
                                //        }
                                //        break;
                                //    case 4:
                                //        DataCheck.FloatData(row, D_dataRow, ref berror, ref sError, i, j, 1, GetExcelDefine.VNExcel[j + 2].ChineseName);
                                //        break;
                                //    case 6:
                                //        DataCheck.StringData(row, D_dataRow, ref berror, ref sError, i, j, 0, GetExcelDefine.VNExcel[j + 2].ChineseName);

                                //        break;
                                //    case 7:
                                //        DataCheck.IntData(row, D_dataRow, ref berror, ref sError, i, j, 0, GetExcelDefine.VNExcel[j + 2].ChineseName);
                                //        break;
                                //    case 8:
                                //        DataCheck.FloatData(row, D_dataRow, ref berror, ref sError, i, j, 0, GetExcelDefine.VNExcel[j + 2].ChineseName);

                                //        break;

                                //    default:
                                //        break;

                                //        //-- CellType需要搭配「NPOI.SS.UserModel命名空間」
                                //}

                                //}
                                #endregion
                                if (row.GetCell(0) != null)
                                {
                                    str閱卷序號 = row.GetCell(0).ToString().ToUpper();
                                    sbError.AppendFormat("閱卷序號：{1}", str閱卷序號);
                                }
                                if (row.GetCell(1) != null)
                                {
                                    str款號 = row.GetCell(1).ToString().ToUpper();
                                }
                                else
                                    berror = 錯誤訊息(sbError, "沒有款號");
                                
                                if (row.GetCell(2) != null)
                                {
                                    str組別 = row.GetCell(2).ToString().ToUpper();
                                }
                                else
                                    berror = 錯誤訊息(sbError, "沒有組別");
                                if (row.GetCell(3) != null)
                                {
                                    str日期 = row.GetCell(3).ToString().ToUpper();
                                }
                                else
                                    berror = 錯誤訊息(sbError, "沒有日期");
                                for (int z = 4; z < 25; z=z+7)
                                {
                                    string str工號 = "";
                                    string strRegex工號 = "V[0-9]{4}", strRegex工段= "[0-9]{2}", strRegex數量= "[0-9]{4}";
                                    bool b確認工號 = true;
                                    if (row.GetCell(z) != null)
                                    {
                                        str工號 = row.GetCell(z).ToString().Trim().ToUpper();
                                        Regex reg = new Regex(strRegex工號);
                                        b確認工號 = (!reg.IsMatch(str工號) && str工號.Length != 5)? false : true;
                                    }
                                    else
                                        b確認工號 = 錯誤訊息(sbError, "沒有工號");
                                    //檢查1~3工段數量
                                    for (int zz = 0; zz < 3; zz++)
                                    {
                                        string str工段 = "", str數量 = "";
                                        bool b確認工段 = true, b確認數量 = true;
                                        if (row.GetCell(z+zz + 1) != null)
                                        {
                                            //永琦工段要三碼，補0
                                            str工段 = "0" + row.GetCell(z + zz + 1).ToString();
                                            Regex reg = new Regex(strRegex工段);
                                            b確認工段 = (!reg.IsMatch(str工段) && str工段.Length != 2) ? false : true;
                                        }
                                        else
                                            b確認工段 = 錯誤訊息(sbError, "沒有工段"+(z+1).ToString());

                                        if (row.GetCell(z + zz + 2) != null)
                                        {
                                            str數量 = row.GetCell(z + zz + 2).ToString();
                                            Regex reg = new Regex(strRegex數量);
                                            b確認數量 = (!reg.IsMatch(str數量) && str數量.Length != 4) ? false : true;
                                        }
                                        else
                                            b確認數量 = 錯誤訊息(sbError, "沒有數量" + (z + 1).ToString());

                                        //工段23沒資料跳過
                                        if (str數量.Length>0 && str數量.Length>0&&zz>0)
                                            continue;

                                        if (!b確認工段|| !b確認數量||sError.Length>0)
                                        {
                                            //str閱卷序號 = "", str款號 = "", str組別 = "", str日期 = "";
                                            D_erroraRow[0] = u_sheet.SheetName.ToString();
                                            D_erroraRow[1] = str閱卷序號;
                                            D_erroraRow[2] = str款號;
                                            D_erroraRow[3] = str組別;
                                            D_erroraRow[4] = str日期;
                                            D_erroraRow[5] = str工號;
                                            D_erroraRow[6] = "錯誤資料：" + i.ToString() + sError;
                                            D_errortable.Rows.Add(D_erroraRow);
                                        }
                                        else
                                        {
                                            D_dataRow[0] = u_sheet.SheetName.ToString();
                                            D_dataRow[1] = str閱卷序號;
                                            D_dataRow[2] = str款號;
                                            D_dataRow[3] = str組別;
                                            D_dataRow[4] = str日期;
                                            D_dataRow[5] = str工號;
                                            D_dataRow[6] = str工段;
                                            D_dataRow[6] = str數量;
                                            D_table.Rows.Add(D_dataRow);
                                        }

                                    }
                                    //if (bcheck)
                                    //{
                                    //    D_table.Rows.Add(D_dataRow);

                                    //}

                                    //if (berror)
                                    //{
                                    //    D_erroraRow[0] = u_sheet.SheetName.ToString();
                                    //    D_erroraRow[1] = row.GetCell(0).ToString();
                                    //    D_erroraRow[2] = "Hàng thứ " + i.ToString() + sError;
                                    //    D_errortable.Rows.Add(D_erroraRow);
                                    //}
                                }




                            }
                            //-- 釋放 NPOI的資源
                            u_sheet = null;
                        }
                        //-- 釋放 NPOI的資源
                        workbook = null;
                        //--Excel資料顯示             
                        DataView D_View2 = new DataView(D_table);
                        ExcelGV.DataSource = D_View2;
                        ExcelGV.DataBind();
                        for (int i = 0; i < Excel欄位數+2; i++)
                        {
                            ExcelGV.HeaderRow.Cells[i].Text = GetExcelDefine.VNExcel[i].VNName;
                        }
                            
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

                            for (int i = 3; i <= u_sheet.LastRowNum; i++)   //-- 每一列做迴圈
                            {
                                //--不包含 Excel表頭列的 "其他資料列"
                                IRow row = (IRow)u_sheet.GetRow(i);
                                DataRow D_dataRow = D_table.NewRow();
                                DataRow D_erroraRow = D_errortable.NewRow();
                                D_dataRow[0] = u_sheet.SheetName.ToString();
                                D_dataRow[1] = SearchTB.Text;
                                bool bcheck = true, berror = false;
                                string sError = "";
                                //for (int j = row.FirstCellNum; j < row.LastCellNum; j++)   //-- 每一個欄位做迴圈
                                for (int j = row.FirstCellNum; j < Excel欄位數; j++)   //有些EXCEL會沒填資料
                                {
                                    //沒有Style就不抓
                                    if (row.GetCell(2) == null)
                                    {
                                        bcheck = false;
                                        break;
                                    }
                                    switch (GetExcelDefine.VNExcel[j + 2].ColumnType)
                                    {
                                        // Type 1：數字 , Type 2：String , Type 3：日期  , Type 6：不需要資料 String, Type 7：不需要資料 int
                                        case 1:
                                            DataCheck.IntData(row, D_dataRow, ref berror, ref sError, i, j, 1, GetExcelDefine.VNExcel[j + 2].ChineseName);
                                            break;
                                        case 2:
                                            DataCheck.StringData(row, D_dataRow, ref berror, ref sError, i, j, 1, GetExcelDefine.VNExcel[j + 2].ChineseName);
                                            break;
                                        case 3:
                                            try
                                            {
                                                D_dataRow[j + 2] = (string.IsNullOrEmpty(row.GetCell(j).ToString())) ? "" : row.GetCell(j).DateCellValue.ToString("yyyyMMdd");
                                                //轉換日期格式
                                            }
                                            catch 
                                            {
                                                //D_dataRow[j] = row.GetCell(j).CellFormula.ToString();
                                                berror = true;
                                                sError += "Hàng thứ " + i.ToString() + "、" + GetExcelDefine.VNExcel[j + 2].ChineseName + "error5。";
                                                D_dataRow[j + 2] = (row.GetCell(j) == null) ? "" : row.GetCell(j).ToString();  //--每一個欄位，都加入同一列 DataRow
                                                                                                                                //throw;
                                            }
                                            break;
                                        case 4:
                                        DataCheck.FloatData(row, D_dataRow, ref berror, ref sError, i, j, 1, GetExcelDefine.VNExcel[j + 2].ChineseName);
                                            break;
                                        case 6:
                                        DataCheck.StringData(row, D_dataRow, ref berror, ref sError, i, j, 0, GetExcelDefine.VNExcel[j + 2].ChineseName);
                                            break;
                                        case 7:
                                        DataCheck.IntData(row, D_dataRow, ref berror, ref sError, i, j, 0, GetExcelDefine.VNExcel[j + 2].ChineseName);
                                            break;
                                        case 8:
                                        DataCheck.FloatData(row, D_dataRow, ref berror, ref sError, i, j, 0, GetExcelDefine.VNExcel[j + 2].ChineseName);
                                            break;
                                        default:
                                            break;
                                    }
                                }
                                if (bcheck)
                                    D_table.Rows.Add(D_dataRow);
                                if (berror)
                                {
                                    D_erroraRow[0] = u_sheet.SheetName.ToString();
                                    D_erroraRow[1] = row.GetCell(0).ToString();
                                    D_erroraRow[2] = "Hàng thứ " + i.ToString()+ sError;
                                    D_errortable.Rows.Add(D_erroraRow);
                                }
                            }
                            //-- 釋放 NPOI的資源
                            u_sheet = null;
                        }
                        //-- 釋放 NPOI的資源
                        workbook = null;
                        //--Excel資料顯示             
                        DataView D_View2 = new DataView(D_table);
                        //GridView1.DataSource = D_View2;
                        //GridView1.DataBind();
                        ExcelGV.DataSource = D_View2;
                        ExcelGV.DataBind();
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
                    Label1.Text = "????  ...... 請先挑選檔案之後，再來上傳";
                }   // FileUpload使用的第一個 if判別式

                if (D_table.Rows.Count > 0)
                    Session["Excel"] = D_table;
                else
                {
                    Session["Excel"] = null;
                    ExcelGV.DataSource = null;
                    ExcelGV.DataBind();

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
                    Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('當日已有匯入資料');</script>");
                else
                    Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('請選擇匯入日期');</script>");
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
                @"INSERT INTO [dbo].[Productivity_Head]
                    (FileName,Area,Creator,Team,Date)
                    VALUES(@FileName,@Area,'Program',@Team,@Date); 
                    SELECT CAST(scope_identity() AS int)";
            using (SqlConnection conn = new SqlConnection(strConnectString))
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.Add("@Date", SqlDbType.NVarChar);
                cmd.Parameters.Add("@FileName", SqlDbType.NVarChar);
                cmd.Parameters.Add("@Area", SqlDbType.NVarChar);
                cmd.Parameters.Add("@Team", SqlDbType.NVarChar);
                cmd.Parameters["@Date"].Value = SearchTB.Text;
                cmd.Parameters["@FileName"].Value = Session["FileName"].ToString();
                cmd.Parameters["@Area"].Value = strArea;
                //cmd.Parameters["@Team"].Value = strImportType;
                try
                {
                    conn.Open();
                    ProductivityHeadID = (Int32)cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    Log.ErrorLog(ex, "Get Productivity_Head uid Error" + Session["FileName"].ToString() + ":", "VN002.aspx");
                }
            }
            return (int)ProductivityHeadID;
        }

        protected void UpLoadBT_Click(object sender, EventArgs e)
        {
            if(SearchTB.Text.Trim()!=""&&F_CheckData())
            { 
                if (Session["ExcelError"] == null)
                    if (Session["Excel"] != null)
                    {
                        DataTable dt = (DataTable)Session["Excel"];
                        //TypeLB.Text = dt.Rows[2][2].ToString();
                        int iIndex = 0;
                        iIndex = GetExcelIdex();
                        if (iIndex > 0)
                            using (SqlConnection conn1 = new SqlConnection(strConnectString))
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
                                        string strInsertColumn="", strInsertData="";
                                        //if (strImportType=="Stitch")
                                        //{
                                        //    //strInsertColumn = ",[QCQty],[ErrorQty],[ErrorUnreturnQty],[OnlineDay]";
                                        //    //strInsertData = ",@QCQty ,@ErrorQty,@ErrorUnreturnQty,@OnlineDay";
                                        //    strInsertColumn = ",[QCQty],[ErrorQty],[OnlineDay],[ErrorRate]";
                                        //    strInsertData = ",@QCQty ,@ErrorQty,@OnlineDay,@ErrorRate";
                                        //}
                                        //TypeLB.Text = i.ToString();
                                        command1.CommandText = string.Format(@"INSERT INTO [dbo].[Productivity_Line]
                                                       ([uid]
                                                       ,[SheetName]
                                                       ,[Dept]
                                                       ,[Customer]
                                                       ,[StyleNo]
                                                       ,[OrderQty]
                                                       ,[OrderShipDate]
                                                       ,[OnlineDate]
                                                       ,[StandardProductivity]
                                                       ,[TeamProductivity]
                                                       ,[GoalProductivity]
                                                       ,[DayProductivity]
                                                       ,[PreProductivity]
                                                       ,[TotalProductivity]
                                                       ,[Person]
                                                       ,[TotalTime]
                                                       ,[Time]
                                                       ,[Percent]
                                                       ,[Difference]
                                                       ,[Efficiency]
                                                       ,[TotalEfficiency]
                                                       ,[ReturnPercent]
                                                       ,[Rmark1]
                                                       ,[Rmark2]
                                                       ,[DayCost1]
                                                       ,[DayCost2]
                                                       ,[DayCost3]
                                                       ,[DayCost4]
                                                       ,[DayCost5]
                                                       ,[DayCost6]
                                                       ,[DayCost7]
                                                       ,[Creator] {0})
                                                 VALUES
                                                       ({1}
                                                       ,@SheetName
                                                       ,@Dept
                                                       ,@Customer
                                                       ,@StyleNo
                                                       ,@OrderQty
                                                       ,@OrderShipDate
                                                       ,@OnlineDate
                                                       ,@StandardProductivity
                                                       ,@TeamProductivity
                                                       ,@GoalProductivity
                                                       ,@DayProductivity
                                                       ,@PreProductivity
                                                       ,@TotalProductivity
                                                       ,@Person
                                                       ,@TotalTime
                                                       ,@Time
                                                       ,@Percent
                                                       ,@Difference
                                                       ,@Efficiency
                                                       ,@TotalEfficiency
                                                       ,@ReturnPercent
                                                       ,@Rmark1
                                                       ,@Rmark2
                                                       ,@DayCost1
                                                       ,@DayCost2
                                                       ,@DayCost3
                                                       ,@DayCost4
                                                       ,@DayCost5
                                                       ,@DayCost6
                                                       ,@DayCost7
                                                       ,'Program' {2} )
                                                       ", strInsertColumn, iIndex, strInsertData);
                                        command1.Parameters.Add("@SheetName", SqlDbType.NVarChar).Value = dt.Rows[i]["SheetName"].ToString();
                                        command1.Parameters.Add("@Dept", SqlDbType.NVarChar).Value = dt.Rows[i]["Dept"].ToString();
                                        command1.Parameters.Add("@Customer", SqlDbType.NVarChar).Value = dt.Rows[i]["Customer"].ToString();
                                        command1.Parameters.Add("@StyleNo", SqlDbType.NVarChar).Value = dt.Rows[i]["StyleNo"].ToString();
                                        command1.Parameters.Add("@OrderQty", SqlDbType.Int).Value = dt.Rows[i]["OrderQty"].ToString();
                                        command1.Parameters.Add("@OrderShipDate", SqlDbType.NVarChar).Value = dt.Rows[i]["OrderShipDate"].ToString();
                                        command1.Parameters.Add("@OnlineDate", SqlDbType.NVarChar).Value = dt.Rows[i]["OnlineDate"].ToString();
                                        command1.Parameters.Add("@StandardProductivity", SqlDbType.Float).Value = dt.Rows[i]["StandardProductivity"].ToString();
                                        command1.Parameters.Add("@TeamProductivity", SqlDbType.Int).Value = dt.Rows[i]["TeamProductivity"].ToString();
                                        command1.Parameters.Add("@GoalProductivity", SqlDbType.Float).Value = dt.Rows[i]["GoalProductivity"].ToString();
                                        command1.Parameters.Add("@DayProductivity", SqlDbType.Int).Value = dt.Rows[i]["DayProductivity"].ToString();
                                        command1.Parameters.Add("@PreProductivity", SqlDbType.Int).Value = dt.Rows[i]["PreProductivity"].ToString();
                                        command1.Parameters.Add("@TotalProductivity", SqlDbType.Int).Value = dt.Rows[i]["TotalProductivity"].ToString();
                                        command1.Parameters.Add("@Person", SqlDbType.Float).Value = dt.Rows[i]["Person"].ToString();
                                        command1.Parameters.Add("@TotalTime", SqlDbType.Float).Value = dt.Rows[i]["TotalTime"].ToString();
                                        command1.Parameters.Add("@Time", SqlDbType.Float).Value = dt.Rows[i]["Time"].ToString();
                                        command1.Parameters.Add("@Percent", SqlDbType.Float).Value = dt.Rows[i]["Percent"].ToString();
                                        command1.Parameters.Add("@Difference", SqlDbType.Int).Value = dt.Rows[i]["Difference"].ToString();
                                        command1.Parameters.Add("@Efficiency", SqlDbType.Float).Value = dt.Rows[i]["Efficiency"].ToString();
                                        command1.Parameters.Add("@TotalEfficiency", SqlDbType.Float).Value = dt.Rows[i]["TotalEfficiency"].ToString();
                                        command1.Parameters.Add("@ReturnPercent", SqlDbType.Float).Value = dt.Rows[i]["ReturnPercent"].ToString();
                                        command1.Parameters.Add("@Rmark1", SqlDbType.NVarChar).Value = dt.Rows[i]["Rmark1"].ToString();
                                        command1.Parameters.Add("@Rmark2", SqlDbType.NVarChar).Value = dt.Rows[i]["Rmark2"].ToString();
                                        command1.Parameters.Add("@DayCost1", SqlDbType.Float).Value = dt.Rows[i]["DayCost1"].ToString();
                                        command1.Parameters.Add("@DayCost2", SqlDbType.Float).Value = dt.Rows[i]["DayCost2"].ToString();
                                        command1.Parameters.Add("@DayCost3", SqlDbType.Float).Value = dt.Rows[i]["DayCost3"].ToString();
                                        command1.Parameters.Add("@DayCost4", SqlDbType.Float).Value = dt.Rows[i]["DayCost4"].ToString();
                                        command1.Parameters.Add("@DayCost5", SqlDbType.Float).Value = dt.Rows[i]["DayCost5"].ToString();
                                        command1.Parameters.Add("@DayCost6", SqlDbType.Float).Value = dt.Rows[i]["DayCost6"].ToString();
                                        command1.Parameters.Add("@DayCost7", SqlDbType.Float).Value = dt.Rows[i]["DayCost7"].ToString();
                                        //if (strImportType == "Stitch")
                                        //{
                                        //    command1.Parameters.Add("@QCQty", SqlDbType.Int).Value = dt.Rows[i]["QCQty"].ToString();
                                        //    command1.Parameters.Add("@ErrorQty", SqlDbType.Int).Value = dt.Rows[i]["ErrorQty"].ToString();
                                        //    //command1.Parameters.Add("@ErrorUnreturnQty", SqlDbType.Int).Value = dt.Rows[i]["ErrorQty"].ToString();
                                        //    command1.Parameters.Add("@OnlineDay", SqlDbType.Int).Value = dt.Rows[i]["OnlineDay"].ToString();
                                        //    command1.Parameters.Add("@ErrorRate", SqlDbType.Int).Value = dt.Rows[i]["ErrorRate"].ToString();
                                        //}
                                        command1.ExecuteNonQuery();
                                        command1.Parameters.Clear();
                                    }
                                    //上傳成功更新Head狀態
                                    command1.CommandText = string.Format(@"UPDATE [dbo].[Productivity_Head] SET [Flag] = 1 ,[Date] = @Date WHERE uid = {0} ", iIndex);
                                    command1.Parameters.Add("@Date", SqlDbType.NVarChar).Value = Session["Date"].ToString();
                                    command1.ExecuteNonQuery();
                                    transaction1.Commit();
                                }
                                catch (Exception ex1)
                                {
                                    try
                                    {
                                        Log.ErrorLog(ex1, "Import Excel Error :" + Session["FileName"].ToString(), "VN002.aspx");
                                    }
                                    catch (Exception ex2)
                                    {
                                        Log.ErrorLog(ex2, "Insert Error Error:" + Session["FileName"].ToString(), "VN002.aspx");
                                    }
                                    finally
                                    {
                                        transaction1.Rollback();
                                        Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('匯入失敗請連絡MIS');</script>");
                                    }
                                }
                                finally
                                {
                                    conn1.Close();
                                    conn1.Dispose();
                                    command1.Dispose();
                                    Session.RemoveAll();
                                    Label1.Text = "資料上傳成功";
                                }
                            }
                        else
                            Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('單頭匯入失敗請連絡MIS');</script>");
                    }
            }
            else
            {
                if(F_CheckData()==false)
                    Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('當日已有匯入資料');</script>");
                else
                    Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('請選擇匯入日期');</script>");
            }
        }


        protected void DeleteBT_Click(object sender, EventArgs e)
        {

                using (SqlConnection conn1 = new SqlConnection(strConnectString))
                {
                    SqlCommand command1 = conn1.CreateCommand();
                    SqlTransaction transaction1;
                    conn1.Open();
                    transaction1 = conn1.BeginTransaction("createExcelImport");

                    command1.Connection = conn1;
                    command1.Transaction = transaction1;
                    try
                    {
                        command1.CommandText = string.Format(@"UPDATE [dbo].[Productivity_Head] SET [Flag] = 2,[ModifyDate]=GETDATE()  WHERE Team = @Team and [Date] = @Date and [Flag] = 1 ");
                        command1.Parameters.Add("@Date", SqlDbType.NVarChar).Value = SearchTB.Text;
                        //command1.Parameters.Add("@Team", SqlDbType.NVarChar).Value = strImportType;
                        command1.ExecuteNonQuery();
                        transaction1.Commit();
                        Label1.Text = "刪除完畢，請再次夾檔";
                    }
                    catch (Exception ex1)
                    {
                        try
                        {
                            Log.ErrorLog(ex1, "Delete Error :" , "VN002.aspx");
                        }
                        catch (Exception ex2)
                        {
                            Log.ErrorLog(ex2, "Delete Error Error:" , "VN002.aspx");
                        }
                        finally
                        {
                            transaction1.Rollback();
                            Page.ClientScript.RegisterStartupScript(Page.GetType(), "", "<script>alert('刪除失敗請連絡MIS');</script>");
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

                using (SqlConnection conn = new SqlConnection(strConnectString1))
                {
                    SqlCommand command = new SqlCommand();
                    command.Connection = conn;
                    command.CommandText = @"SELECT [uid]
                                              ,[Team]
                                              ,[FileName]
                                              ,[Date]
                                              ,[Area]
                                              ,[Flag]
                                              ,[CreateDate]
                                              ,[Creator]
                                            FROM [dbo].[Productivity_Head]
                                            where Team = @Team and Date = @Date and Flag = 1";
                    command.CommandType = CommandType.Text;
                    //command.Parameters.Add("@Team", SqlDbType.NVarChar).Value =strImportType ;
                    command.Parameters.Add("@Date", SqlDbType.NVarChar).Value = SearchTB.Text;
                    conn.Open();
                    SqlDataReader reader = command.ExecuteReader();

                    if (reader.HasRows)
                    {
                        bcheck = false;
                        DataTable dt = new DataTable();
                        dt.Load(reader);
                    }
                    reader.Close();
                }

            return bcheck;
        }
        
    }
}