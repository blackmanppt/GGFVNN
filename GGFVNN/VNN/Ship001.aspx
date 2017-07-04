<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Ship001.aspx.cs" Inherits="GGFPortal.VNN.Ship001" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=12.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="ajaxToolkit" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>出口大表</title>
    <link href="../Content/bootstrap-theme.min.css" rel="stylesheet" />
    <link href="../Content/bootstrap.min.css" rel="stylesheet" />
    <link href="../Content/style.css" rel="stylesheet" />
    <script src="../scripts/bootstrap.min.js"></script>
    <script src="../scripts/jquery-3.1.1.min.js"></script>
    <script src="../scripts/scripts.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <div class="container-fluid">
            <div class="row">
                <div class="col-md-2">
                    <nav class="navbar navbar-default" role="navigation">
                        <h3 class="text-info text-center">出口大表
                        </h3>
                        <div class="collapse navbar-collapse " id="bs-example-navbar-collapse-1">
<%--                            <div class="form-group">
                            <h4>年</h4>
<asp:DropDownList ID="YearDDL" runat="server" class="form-control"></asp:DropDownList>
						</div> --%>
                    <h4>訂單日期</h4>
                    <div class="form-group">
                        <asp:TextBox ID="StartTB" runat="server" class="form-control"></asp:TextBox>
                        <ajaxToolkit:calendarextender ID="StartTB_CalendarExtender" runat="server" BehaviorID="StartTB_CalendarExtender" TargetControlID="StartTB" Format="yyyy-MM-dd"/>
                        <asp:TextBox ID="EndTB" runat="server" class="form-control"></asp:TextBox>
                        <ajaxToolkit:calendarextender ID="EndTB_CalendarExtender" runat="server" BehaviorID="EndTB_CalendarExtender" TargetControlID="EndTB" Format="yyyy-MM-dd" />
						</div> 
<%--                            <h4>地區</h4>
                            <div class="form-group">
                                <asp:CheckBoxList ID="DeptCBL" runat="server" DataSourceID="SqlDataSource1" DataTextField="部門" DataValueField="部門" CssClass="checkbox checkbox-inline"></asp:CheckBoxList>
                                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:DBConnectionString %>" SelectCommand="SELECT DISTINCT [部門] FROM [View訂單資料] ORDER BY [部門]"></asp:SqlDataSource>
                            </div>--%>

                            <div class="form-group">
                            <asp:Button ID="SearchBT" runat="server" Text="Search" class="btn btn-default" OnClick="SearchBT_Click" />
                            <asp:Button ID="ClearBT" runat="server" Text="Clear" class="btn btn-default" OnClick="ClearBT_Click" />

                            </div>
                            <asp:Literal ID="MessageLT" runat="server"></asp:Literal>


                        </div>

                    </nav>
                </div>
                <div class="col-md-10">
                    <rsweb:ReportViewer ID="ReportViewer1" runat="server" Font-Names="Verdana" Font-Size="8pt" WaitMessageFont-Names="Verdana" WaitMessageFont-Size="14pt" Height="768px" Width="1024px" Visible="False" >
                        <LocalReport ReportPath="VNN\Report\Report1.rdlc" DisplayName="訂單資料">
                        </LocalReport>
                    </rsweb:ReportViewer>
                </div>
            </div>
        </div>
    </form>
</body>
</html>
