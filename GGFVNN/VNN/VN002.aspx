  <%@ Page Language="C#" AutoEventWireup="true" CodeBehind="VN002.aspx.cs" Inherits="GGFVNN.VNN.VN002" %>

<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="ajaxToolkit" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>資料上傳</title>
    <style type="text/css">
        .line{
            border: 2px solid black;
        }
        
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <table style="border: 2px solid #000000; width:500px; border-collapse: collapse;border: 2px solid black;"  >
            <tr class="line">
                <td colspan="3">
                    <h1>
                        <asp:Label ID="AreaLB" runat="server"></asp:Label>
                    <asp:Label ID="TypeLB" runat="server"></asp:Label>
                    </h1>
                </td>
            </tr>
            <tr>
                <td class="line">
                    <asp:Label ID="Label4" runat="server" Text="UpDate："></asp:Label>
                    
                </td>
                <td class="line">

                    <asp:TextBox ID="SearchTB" runat="server" ></asp:TextBox>
  <%--                   <ajaxToolkit:TextBoxWatermarkExtender ID="SearchTB_TextBoxWatermarkExtender" runat="server" TargetControlID="SearchTB" />
                   --%>
                   <ajaxToolkit:TextBoxWatermarkExtender ID="SearchTB_TextBoxWatermarkExtender" runat="server" TargetControlID="SearchTB" WatermarkText="請填入資料" />
                    
                    <ajaxToolkit:calendarextender ID="SearchTB_CalendarExtender" runat="server" Format="yyyyMMdd" TargetControlID="SearchTB" />
                </td>
                <td class="line">

                    <asp:Button ID="DeleteBT" runat="server" Text="DeleteData" OnClick="DeleteBT_Click" />
                </td>
            </tr>
            <tr>
                <td  class="line">                   
                    <asp:Label ID="Label3" runat="server" Text="File Update"></asp:Label>
                </td>
                <td class="line">          
                    <asp:FileUpload ID="FileUpload1" runat="server" />
                </td>
                <td>
                    <asp:Button ID="CheckBT" runat="server" Text="Check" OnClick="CheckBT_Click" />
                </td>
            </tr>
            <tr>
                <td class="line"></td>
                <td class="line">

                    <asp:Button ID="TeamCodeBT" runat="server" Text="TeamCode" OnClick="TeamCodeBT_Click" />

                    <asp:Button ID="TempExcel" runat="server" Text="TempExcel" OnClick="TempExcel_Click" />
                </td>
                <td class="line">
                    <asp:Button ID="UpLoadBT" runat="server" Text="UpLoad" OnClick="UpLoadBT_Click" />
                </td>
            </tr>
            <tr>
                <td colspan="3" class="line">
                    <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
                </td>
            </tr>
        </table>
    
    </div>
        <div>
            <asp:GridView ID="ErrorGV" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None">
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:BoundField DataField="SheetName" HeaderText="SheetName" />
                    <asp:BoundField DataField="Dept" HeaderText="部門 Bộ Phận" />
                    <asp:BoundField DataField="Error" HeaderText="Error" />
                </Columns>
                <FooterStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#FFCC66" ForeColor="#333333" HorizontalAlign="Center" />
                <RowStyle BackColor="#FFFBD6" ForeColor="#333333" />
                <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="Navy" />
                <SortedAscendingCellStyle BackColor="#FDF5AC" />
                <SortedAscendingHeaderStyle BackColor="#4D0000" />
                <SortedDescendingCellStyle BackColor="#FCF6C0" />
                <SortedDescendingHeaderStyle BackColor="#820000" />
            </asp:GridView>
            <br />
            <%--<asp:GridView ID="ExcelGV" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" BorderColor="#DEBA84" BorderStyle="None" BorderWidth="1px" CellPadding="3" CellSpacing="2">
                <Columns>
                    <asp:BoundField DataField="SheetName" HeaderText="SheetName" />
                    <asp:BoundField DataField="Dept" HeaderText="部門 Bộ Phận" />
                    <asp:BoundField DataField="Customer" HeaderText="客户 Khách Hàng" />
                    <asp:BoundField DataField="StyleNo" HeaderText="款號 Mã Hàng" />
                    <asp:BoundField DataField="OrderQty" HeaderText="訂單數量 SL đơn hàng" />
                    <asp:BoundField DataField="TeamProductivity" HeaderText="组生产量 sản lượng tổ" />
                    <asp:BoundField DataField="OrderShipDate" HeaderText="訂單交期 Ngày giao hàng" />
                    <asp:BoundField DataField="OnlineDate" HeaderText="上線日期 Ngày lên chuyền" DataFormatString="{0:d}" />
                    <asp:BoundField DataField="StandardProductivity" HeaderText="1/人8H標準產量 M.tiêu 1ng/8H" />
                    <asp:BoundField DataField="Person" HeaderText="實際工作人數 Số cn  làm" />
                    <asp:BoundField DataField="Time" HeaderText="工時 Thời gian làm việc" />
                    <asp:BoundField DataField="TotalTime" HeaderText="總工時 tổng thời gian làm việc của cả tổ" />
                    <asp:BoundField DataField="Percent" HeaderText="百分比 phần trăm" />
                    <asp:BoundField DataField="GoalProductivity" HeaderText="今日目標產量 SL Mụ tiêu ngày" />
                    <asp:BoundField DataField="DayProductivity" HeaderText="今日產量 Sản Lượng Ngày" />
                    <asp:BoundField DataField="PreProductivity" HeaderText="前天累積產量 Tích luỹ trước 1 ngày" />
                    <asp:BoundField DataField="TotalProductivity" HeaderText="累積產量 Sản lượng tích luỹ" />
                    <asp:BoundField DataField="Difference" HeaderText="差異量 Sản lượng  tích luỹ" />
                    <asp:BoundField DataField="Efficiency" HeaderText="組各別效率 hiệu quả của 1 mã hàng trong tổ" />
                    <asp:BoundField DataField="TotalEfficiency" HeaderText="組效率 Tỉ Lệ Hiệu Suất %" />
                    <asp:BoundField DataField="ReturnPercent" HeaderText="返修率 Tỉ lệ hàng sửa" />
                    <asp:BoundField DataField="Rmark1" HeaderText="责任归属及上线天数 (文字備註)" />
                    <asp:BoundField DataField="Rmark2" HeaderText="顏色" />
                    <asp:BoundField DataField="DayCost1" HeaderText="今日各組成本 giá thành các tổ" />
                    <asp:BoundField DataField="DayCost2" HeaderText="今日生產成本/DZ giá thành SP/DZ" />
                    <asp:BoundField DataField="DayCost3" HeaderText="工繳收入/DZ Đơn giá bán/DZ" />
                    <asp:BoundField DataField="DayCost4" HeaderText="今日工繳收入 /DZ Doanh thu" />
                    <asp:BoundField DataField="DayCost5" HeaderText="今日生產損益 USD USD Lãi,lỗ" />
                    <asp:BoundField DataField="DayCost6" HeaderText="(CM-COST)/CM 損 益 % lãi lỗ" />
                    <asp:BoundField DataField="DayCost7" HeaderText="累積損益" />
                    <asp:BoundField DataField="QCQty" HeaderText="QC檢驗數量" />
                    <asp:BoundField DataField="ErrorQty" HeaderText="瑕疵數可返修" />
                    <asp:BoundField DataField="ErrorUnreturnQty" HeaderText="瑕疵數不可返修" />
                    <asp:BoundField DataField="OnlineDay" HeaderText="上線天數" />
                </Columns>
                <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
                <HeaderStyle BackColor="#A55129" Font-Bold="True" ForeColor="White" />
                <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
                <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
                <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
                <SortedAscendingCellStyle BackColor="#FFF1D4" />
                <SortedAscendingHeaderStyle BackColor="#B95C30" />
                <SortedDescendingCellStyle BackColor="#F1E5CE" />
                <SortedDescendingHeaderStyle BackColor="#93451F" />
            </asp:GridView>--%>

        </div>
        <asp:GridView ID="ExcelGV" runat="server" BackColor="#DEBA84" BorderColor="#DEBA84" BorderStyle="None" BorderWidth="1px" CellPadding="3" CellSpacing="2">
            <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
            <HeaderStyle BackColor="#A55129" Font-Bold="True" ForeColor="White" />
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#FFF1D4" />
            <SortedAscendingHeaderStyle BackColor="#B95C30" />
            <SortedDescendingCellStyle BackColor="#F1E5CE" />
            <SortedDescendingHeaderStyle BackColor="#93451F" />
        </asp:GridView>
    </form>
</body>
</html>
