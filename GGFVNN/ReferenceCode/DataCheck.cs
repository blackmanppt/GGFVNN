using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace GGFGAMA.ReferenceCode
{
    public class DataCheck
    {
        static string strConnectString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["DBConnectionString"].ToString();
        static string strConnectString1 = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["GGFConnectionString"].ToString();
        /// <summary>
        /// 
        /// </summary>
        /// <param name="代工廠"></param>
        /// <param name="月份"></param>
        /// <returns></returns>
        public Boolean Check工時Lock(string str代工廠, string str月份)
        {
            bool bcheck = true;
            using (SqlConnection conn = new SqlConnection(strConnectString))
            {
                SqlCommand command = new SqlCommand();
                command.Connection = conn;
                command.CommandText = @"SELECT *
                                        FROM [dbo].[ProductivityCost]
                                        where VendorId = @VendorId and Year = @Year and Month = @Month and Flag = 1";
                command.CommandType = CommandType.Text;
                command.Parameters.Add("@VendorId", SqlDbType.NVarChar).Value = str代工廠;
                command.Parameters.Add("@Year", SqlDbType.NVarChar).Value = str月份.Substring(0, 4);
                command.Parameters.Add("@Month", SqlDbType.NVarChar).Value = str月份.Substring(4, 2);
                conn.Open();
                SqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    bcheck = false;
                    //DataTable dt = new DataTable();
                    //dt.Load(reader);
                }
                reader.Close();
            }
            return bcheck;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="D_dataRow"></param>
        /// <param name="berror"></param>
        /// <param name="sError"></param>
        /// <param name="i"></param>
        /// <param name="j"></param>
        /// <param name="x"></param>
        /// <param name="ColumnName"></param>
        public static void StringData(IRow row, DataRow D_dataRow, ref bool berror, ref string sError, int i, int j, int x, string ColumnName)
        {
            try
            {
                if (x == 1 && row.GetCell(j) == null)
                {
                    berror = true;
                    sError += ColumnName + "必填欄位未填資料。";
                    D_dataRow[j + 2] = (row.GetCell(j) == null) ? "" : row.GetCell(j).ToString();  //--每一個欄位，都加入同一列 DataRow

                }
                else
                {
                    string strString = "";
                    strString = (string.IsNullOrEmpty(row.GetCell(j).ToString())) ? "" : row.GetCell(j).ToString().Trim();   //--每一個欄位，都加入同一列 DataRow
                    D_dataRow[j + 2] = strString;
                    if (j == 2)
                    {
                        using (SqlConnection conn = new SqlConnection(strConnectString1))
                        {
                            // Create the command and set its properties.
                            SqlCommand command = new SqlCommand();
                            command.Connection = conn;
                            command.CommandText = "select * from ordc_bah1 where bah_status<> 'CA' and cus_item_no = @cus_item_no";
                            command.CommandType = CommandType.Text;

                            // Add the input parameter and set its properties.
                            SqlParameter parameter = new SqlParameter();
                            parameter.ParameterName = "@cus_item_no";
                            parameter.SqlDbType = SqlDbType.NVarChar;
                            parameter.Direction = ParameterDirection.Input;
                            parameter.Value = strString;

                            // Add the parameter to the Parameters collection. 
                            command.Parameters.Add(parameter);

                            // Open the connection and execute the reader.
                            conn.Open();
                            SqlDataReader reader = command.ExecuteReader();

                            if (reader.HasRows == false)
                            {
                                berror = true;
                                sError += "Style_no(" + strString + ")：Không có tài liệu";
                            }
                            reader.Close();
                        }
                    }
                }

            }
            catch
            {
                //x==1代表需要檢查資料
                if (x == 1)
                {
                    berror = true;
                    sError += ColumnName + "欄匯入失敗。";
                }
                //else
                //    sError += "第" + i.ToString() + "行、第" + (j).ToString() + "欄匯入失敗(非必要資料不影響匯入)。";
                D_dataRow[j + 2] = (row.GetCell(j) == null) ? "" : row.GetCell(j).ToString();  //--每一個欄位，都加入同一列 DataRow
                //throw;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="D_dataRow"></param>
        /// <param name="berror"></param>
        /// <param name="sError"></param>
        /// <param name="i"></param>
        /// <param name="j"></param>
        /// <param name="x"></param>
        /// <param name="ColumnName"></param>
        public static void IntData(IRow row, DataRow D_dataRow, ref bool berror, ref string sError, int i, int j, int x, string ColumnName)
        {
            //x==1代表需要檢查資料
            if ((row.GetCell(j) != null) && (row.GetCell(j).CellType == CellType.Formula))  //== v.1.2.4版修改。2.x版只是修正英文大小寫。
            {
                //D_dataRow[j] = row.GetCell(j).NumericCellValue.ToString();
                ////-- 表示格子裡面，公式運算後的「值」，是數字（Numeric）。
                try
                {
                    D_dataRow[j + 2] = row.GetCell(j).NumericCellValue.ToString();
                }
                catch
                {
                    if (x == 1)
                    {
                        berror = true;
                        sError += ColumnName + "公式錯誤1。";
                    }
                    //else
                    //    sError += "第" + i.ToString() + "行、第" + (j).ToString() + "欄公式錯誤(非必要資料不影響匯入)。";
                    D_dataRow[j + 2] = (row.GetCell(j) == null) ? "" : "0";  //--每一個欄位，都加入同一列 DataRow
                }

            }
            else
            {

                try
                {
                    int iout = 0;
                    if (string.IsNullOrEmpty(row.GetCell(j).ToString()))
                    {
                        if (x == 1)
                        {
                            berror = true;
                            sError += ColumnName + "未填資料。";
                        }
                        D_dataRow[j + 2] = "0";  //--每一個欄位，都加入同一列 DataRow
                    }
                    else
                    {
                        if (int.TryParse(row.GetCell(j).ToString(), out iout) == false)
                        {
                            if (x == 1)
                            {
                                berror = true;
                                sError += ColumnName + "非數字：" + row.GetCell(j).ToString() + "。";
                            }
                            //else
                            //    sError += "第" + i.ToString() + "行、第" + (j).ToString() + "非數字(非必要資料不影響匯入)。";
                            D_dataRow[j + 2] = (row.GetCell(j) == null) ? "0" : row.GetCell(j).ToString();  //--每一個欄位，都加入同一列 DataRow

                        }
                        else
                            D_dataRow[j + 2] = (row.GetCell(j) == null) ? "0" : row.GetCell(j).ToString();  //--每一個欄位，都加入同一列 DataRow 
                    }
                    //D_dataRow[j + 2] = (string.IsNullOrEmpty(row.GetCell(j).ToString())) ? "0" : (int.TryParse( row.GetCell(j).ToString(),out iout)==false)?"0": row.GetCell(j).ToString();   //--每一個欄位，都加入同一列 DataRow
                }
                catch
                {
                    if (x == 1)
                    {
                        berror = true;
                        sError += ColumnName + "公式錯誤2。";
                    }
                    //else
                    //    sError += "第" + i.ToString() + "行、第" + (j).ToString() + "欄公式錯誤(非必要資料不影響匯入)。";
                    D_dataRow[j + 2] = (row.GetCell(j) == null) ? "0" : row.GetCell(j).ToString();  //--每一個欄位，都加入同一列 DataRow
                    //throw;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="D_dataRow"></param>
        /// <param name="berror"></param>
        /// <param name="sError"></param>
        /// <param name="i"></param>
        /// <param name="j"></param>
        /// <param name="x"></param>
        /// <param name="ColumnName"></param>
        public static void FloatData(IRow row, DataRow D_dataRow, ref bool berror, ref string sError, int i, int j, int x, string ColumnName)
        {
            //x==1代表需要檢查資料
            if ((row.GetCell(j) != null) && (row.GetCell(j).CellType == CellType.Formula))  //== v.1.2.4版修改。2.x版只是修正英文大小寫。
            {
                //D_dataRow[j] = row.GetCell(j).NumericCellValue.ToString();
                ////-- 表示格子裡面，公式運算後的「值」，是數字（Numeric）。
                try
                {
                    D_dataRow[j + 2] = row.GetCell(j).NumericCellValue.ToString();
                }
                catch
                {
                    if (x == 1)
                    {
                        berror = true;
                        sError += ColumnName + "公式錯誤3。";
                    }
                    //else 
                    //    sError += "第" + i.ToString() + "行、第" + (j).ToString() + "欄公式錯誤(非必要資料不影響匯入)。";
                    D_dataRow[j + 2] = (row.GetCell(j) == null) ? "" : "0";  //--每一個欄位，都加入同一列 DataRow
                }

            }
            else
            {

                try
                {

                    if (string.IsNullOrEmpty(row.GetCell(j).ToString()))
                    {
                        if (x == 1)
                        {
                            berror = true;
                            sError += ColumnName + "必填欄位未填資料。";
                        }
                        D_dataRow[j + 2] = "0";  //--每一個欄位，都加入同一列 DataRow
                    }
                    else
                    {
                        double dout = 0;
                        if (double.TryParse(row.GetCell(j).ToString(), out dout) == false)
                        {
                            if (x == 1)
                            {
                                berror = true;
                                sError += ColumnName + "非數字。";
                            }
                            //else
                            //    sError += "第" + i.ToString() + "行、第" + (j).ToString() + "非數字(非必要資料不影響匯入)。";
                            D_dataRow[j + 2] = (row.GetCell(j) == null) ? "0" : row.GetCell(j).ToString();  //--每一個欄位，都加入同一列 DataRow
                        }
                        else
                            D_dataRow[j + 2] = (row.GetCell(j) == null) ? "0" : row.GetCell(j).ToString();  //--每一個欄位，都加入同一列 DataRow 
                    }
                }
                catch
                {
                    if (x == 1)
                    {
                        berror = true;
                        sError += ColumnName + "非數字。";
                    }
                    //else
                    //    sError += "第" + i.ToString() + "行、第" + (j).ToString() + "非數字(非必要資料不影響匯入)。";
                    D_dataRow[j + 2] = (row.GetCell(j) == null) ? "0" : row.GetCell(j).ToString();  //--每一個欄位，都加入同一列 DataRow
                    //throw;
                }
            }
        }
    }

}