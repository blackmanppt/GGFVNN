<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="GM003.aspx.cs" Inherits="GG.GAMA.GM003" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>GAMA Import</title>
         <script src="../Scripts/jquery-3.4.1.min.js"></script>

    <script src="../Scripts/bootstrap.min.js"></script>
    <link href="../Content/bootstrap.min.css" rel="stylesheet" />
    <style type="text/css">
        .line{
            border: 2px solid black;
        }
        
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
    <div>
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

                    <asp:TextBox ID="SearchTB" runat="server"  CssClass="form-control" ></asp:TextBox>

                    

                                       

                    

                                       <ajaxToolkit:TextBoxWatermarkExtender ID="SearchTB_TextBoxWatermarkExtender" runat="server" TargetControlID="SearchTB" WatermarkText="請選擇匯入日期" />
                    
                    <ajaxToolkit:calendarextender ID="SearchTB_CalendarExtender" runat="server" Format="yyyy-MM-dd" TargetControlID="SearchTB" />
                    
                </td>
                <td class="line">

                    <asp:Button ID="DeleteBT" runat="server" Text="DeleteData" OnClick="DeleteBT_Click" CssClass="btn btn-outline-danger" />
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
                    <asp:Button ID="CheckBT" runat="server" Text="Check" OnClick="CheckBT_Click" CssClass="btn btn-primary" />
                </td>
            </tr>
            <tr>
                <td class="line"></td>
                <td class="line">
                </td>
                <td class="line">
                    <asp:Button ID="UpLoadBT" runat="server" Text="UpLoad" OnClick="UpLoadBT_Click" CssClass="btn btn-secondary" />
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
            <asp:GridView ID="ErrorGV" runat="server"  CellPadding="4" ForeColor="#333333" GridLines="None" CssClass="table">
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
            <asp:GridView ID="ErrorGV2" runat="server" CssClass="table table-danger">
            </asp:GridView>
            <br />
           

        </div>
        <asp:GridView ID="ExcelGV" runat="server" BackColor="#DEBA84" BorderColor="#DEBA84" BorderStyle="None" BorderWidth="1px" CellPadding="3" CellSpacing="2" CssClass="table">
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
