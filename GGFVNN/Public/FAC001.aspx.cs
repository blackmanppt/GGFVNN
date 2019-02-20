using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace GGFGAMA.Public
{
    public partial class A001 : System.Web.UI.Page
    {
        static string strConnectString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["GGFConnectionString"].ToString();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["username"] == null)
                Response.Redirect("~/LoginIndex.aspx");
            DBInit();

        }
        private void DBInit()
        {
            string strsql,strhtmltitle;
            StringBuilder sbbody = new StringBuilder();
            DataSet ds = new DataSet();
            strsql = string.Format(@"
                        select 
                        a.site,a.ord_nbr
                        ,b.cus_item_no 
                        ,b.ord_qty 
                        ,dbo.F_start_dateDate(a.site,a.mps_nbr,'MA01') as CutDate
                        ,dbo.F_mpsc_ctl_mQty(a.site,a.mps_nbr,'MA01') as Cut
                        ,dbo.F_start_dateDate(a.site,a.mps_nbr,'MA02') as SewDate
                        ,dbo.F_mpsc_ctl_mQty(a.site,a.mps_nbr,'MA02') as Sew
                        ,dbo.F_start_dateDate(a.site,a.mps_nbr,'MA03') as QCDate
                        ,dbo.F_mpsc_ctl_mQty(a.site,a.mps_nbr,'MA03') as QC
                        ,dbo.F_start_dateDate(a.site,a.mps_nbr,'MA04') as IrionDate
                        ,dbo.F_mpsc_ctl_mQty(a.site,a.mps_nbr,'MA04') as Irion
                        ,dbo.F_start_dateDate(a.site,a.mps_nbr,'MA05') as PackDate
                        ,dbo.F_mpsc_ctl_mQty(a.site,a.mps_nbr,'MA05') as Pack
                        ,dbo.F_mpsc_ctl_mQty(a.site,a.mps_nbr,'MP01') as FinishedQty
                        ,dbo.F_mpsc_cut_finishQty(a.site,a.mps_nbr,a.ord_nbr) as CutAccumulation
                        ,dbo.F_mpsc_seam_finishQty(a.site,a.ord_nbr) as SewAccumulation
                        ,dbo.F_mpsc_chkQty(a.site,a.ord_nbr) as QCAccumulation
                        ,dbo.F_mpsc_iron_finishQty(a.site,a.ord_nbr) as IrionAccumulation
                        ,dbo.F_mpsc_boxQty(a.site,a.ord_nbr) as PackAccumulation
                            from mpsc_ctl_m a
                        left join ordc_bah1 b on a.site =b.site and a.ord_nbr=b.ord_nbr and b.bah_status<>'CA'
                        where a.assign_close ='N' 
                        and a.item_no like '{0}%' 
                        and a.status<>'CA' and b.cus_item_no is not null and a.ord_nbr in (select  ord_nbr from mpsc_manuf mpsc_manuf where upper(status)  in ('NA','OP'))  
                        order by a.modify_date desc", Session["username"].ToString());
            using (SqlConnection connection = new SqlConnection(strConnectString))
            {
                SqlDataAdapter myAdapter = new SqlDataAdapter(strsql, connection);
                myAdapter.Fill(ds, "FAC");
            }
            
            
            foreach (DataRow dr in ds.Tables["FAC"].Rows)
            {
                string strStyleNO,strShippingDate="",strFinishQty;
                string strCutOnLineDate, strSewOnLineDate, strQCOnLineDate, strIrionOnLineDate, strPackOnLineDate;
                int iCutPreQty, iSewPreQty, iQCPreQty, iIrionPreQty, iPackPreQty, iOrderQty;
                //string strCutEfficient, strSewEfficient, strQCEfficient, strIrionEfficient, strPackEfficient;
                int iCutAccumulation, iSewAccumulation, iQCAccumulation, iIrionAccumulation, iPackAccumulation;
                

                strStyleNO = (dr["cus_item_no"] == DBNull.Value) ? "" : dr["cus_item_no"].ToString();
                iOrderQty = (dr["ord_qty"] == DBNull.Value) ? 0 : Convert.ToInt32((decimal)dr["ord_qty"]);
                //strShippingDate = (dr["plan_shp_date"] == DBNull.Value) ? "" : dr["plan_shp_date"].ToString();
                strFinishQty = (dr["FinishedQty"] == DBNull.Value) ? "0" : ((int)dr["FinishedQty"]).ToString();
                strCutOnLineDate = (dr["CutDate"] == DBNull.Value) ? "" : ((DateTime)dr["CutDate"]).ToString("yyyy/MM/dd");
                strSewOnLineDate = (dr["SewDate"] == DBNull.Value) ? "" : ((DateTime)dr["SewDate"]).ToString("yyyy/MM/dd");
                strQCOnLineDate = (dr["QCDate"] == DBNull.Value) ? "" : ((DateTime)dr["QCDate"]).ToString("yyyy/MM/dd");
                strIrionOnLineDate = (dr["IrionDate"] == DBNull.Value) ? "" : ((DateTime)dr["IrionDate"]).ToString("yyyy/MM/dd");
                strPackOnLineDate = (dr["PackDate"] == DBNull.Value) ? "" : ((DateTime)dr["PackDate"]).ToString("yyyy/MM/dd");
                iCutPreQty = (dr["Cut"] == DBNull.Value) ? 0 : (int)dr["Cut"];
                iSewPreQty = (dr["Sew"] == DBNull.Value) ? 0 : (int)dr["Sew"];
                iQCPreQty = (dr["QC"] == DBNull.Value) ? 0 : (int)dr["QC"];
                iIrionPreQty = (dr["Irion"] == DBNull.Value) ? 0 :(int)dr["Irion"];
                iPackPreQty = (dr["Pack"] == DBNull.Value) ? 0 : (int)dr["Pack"];
                iCutAccumulation = (dr["CutAccumulation"] == DBNull.Value) ? 0 : Convert.ToInt32((decimal)dr["CutAccumulation"]);
                iSewAccumulation = (dr["SewAccumulation"] == DBNull.Value) ? 0 : Convert.ToInt32((decimal)dr["SewAccumulation"]);
                iQCAccumulation = (dr["QCAccumulation"] == DBNull.Value) ? 0 : Convert.ToInt32((decimal)dr["QCAccumulation"]);
                iIrionAccumulation = (dr["IrionAccumulation"] == DBNull.Value) ? 0 : Convert.ToInt32((decimal)dr["IrionAccumulation"]);
                iPackAccumulation = (dr["PackAccumulation"] == DBNull.Value) ? 0 : Convert.ToInt32((decimal)dr["PackAccumulation"]);
                
                using (SqlConnection connection = new SqlConnection(strConnectString)) 
                {
                    DataTable DT = new DataTable();
                    SqlDataAdapter myAdapter = new SqlDataAdapter(string.Format(@"select plan_shp_date from mpsc_manuf where site ='{0}' and ord_nbr='{1}'	group by plan_shp_date", dr["site"].ToString(), dr["ord_nbr"].ToString()), connection);
                    myAdapter.Fill(DT);
                    foreach (DataRow dr2 in DT.Rows)
                    {
                        strShippingDate +=(strShippingDate.Length>0)?" </br> "+ ((DateTime)dr2["plan_shp_date"]).ToString("yyyy/MM/dd") : ((DateTime)dr2["plan_shp_date"]).ToString("yyyy/MM/dd");
                    }
                }

                sbbody.AppendFormat(
                #region html
                    @"     
<br>
                    <HR color='#000000' size='1' width='100%'  align='left'>
                    <HR color='#696969' size='5' width='100%'  align='left'>
                    <div id='container'> 
                    <table background='../IMG/01.jpg' style='text-align: left; width: 100%;' border='1' cellpadding='2' cellspacing='2'>
                    <tbody>
                    <tr>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>STYLE<span style=''>&nbsp; 
                    </span>NO</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 10pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>ORDER
                    QTY</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 8pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>IN-LINE
                    DATE</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 8pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>PROCESS</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 8pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>PREVIOUS
                    QTY</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 8pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>ACCUMULATION</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 8pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>EFFICIENT%</span></b></td>
                    </tr>
                    <tr>

                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; ' lang='EN-US'>{0}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{1}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{2}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 10pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>CUTTING</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{3}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{4}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{5}</span></b></td>
                    </tr>

                    <tr>
                    <td style='vertical-align: top;'><br>
                    </td>
                    <td style='vertical-align: top;'><br>
                    </td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{6}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 10pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>SEWING</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{7}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{8}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{9}</span></b></td>
                    </tr>

                    <tr>
                    <td style='vertical-align: top;'><b><span style='font-size: 10pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>SHIPPING DATE</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 10pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>FINISHED QTY</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{20}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 10pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'>PACKING</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{21}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{22}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{23}
                    </span></b></td>
                    </tr>

                    <tr>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: #ffffff;' lang='EN-US'>{14}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>{15}</span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'></span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 10pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'></span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'></span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'></span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'></span></b></td>
                    </tr>

                    <tr>
                    <td style='vertical-align: top;'><br>
                    </td>
                    <td style='vertical-align: top;'><br>
                    </td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'></span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 10pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: black;' lang='EN-US'></span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'></span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'></span></b></td>
                    <td style='vertical-align: top;'><b><span style='font-size: 12pt; font-family: 
                    &quot;&#26032;&#32048;&#26126;&#39636;&quot;,serif; color: red;' lang='EN-US'>
                    </span></b></td>
                    </tr>
                    </tbody>
                    </table>
                    </div>        
                    "
                #endregion
                                    , strStyleNO
                                    , iOrderQty
                                    , strCutOnLineDate
                                    , iCutPreQty
                                    , iCutAccumulation
                                    , ((float)iCutPreQty / (float)iOrderQty * 100).ToString("0.00")
                                    , strSewOnLineDate
                                    , iSewPreQty
                                    , iSewAccumulation
                                    , ((float)iSewPreQty / (float)iOrderQty * 100).ToString("0.00")
                                    , strQCOnLineDate
                                    , iQCPreQty
                                    , iQCAccumulation
                                    , ((float)iQCPreQty / (float)iOrderQty * 100).ToString("0.00")
                                    , strShippingDate
                                    , strFinishQty
                                    , strIrionOnLineDate
                                    , iIrionPreQty
                                    , iIrionAccumulation
                                    , ((float)iIrionPreQty / (float)iOrderQty * 100).ToString("0.00")
                                    , strPackOnLineDate
                                    , iPackPreQty
                                    , iPackAccumulation
                                    , ((float)iPackPreQty / (float)iOrderQty * 100).ToString("0.00")
                                    );

            }
            strhtmltitle = @"<img src='../IMG/GGF.gif' width='80' higth='80'>";
            Literal1.Text = strhtmltitle + sbbody;
        }
    }
}