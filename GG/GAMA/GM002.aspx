<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="GM002.aspx.cs" Inherits="GG.GAMA.GM002" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=12.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>GAMA工段表</title>
    <script src="../Scripts/jquery-3.4.1.min.js"></script>

    <script src="../scripts/bootstrap-4.3.1/site/docs/4.3/examples/dashboard/dashboard.js"></script>
    <link href="../scripts/bootstrap-4.3.1/site/docs/4.3/examples/dashboard/dashboard.css" rel="stylesheet" />
    <script src="../scripts/bootstrap-4.3.1/dist/js/bootstrap.min.js"></script>
    <link href="../scripts/bootstrap-4.3.1/dist/css/bootstrap.min.css" rel="stylesheet" />


</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <nav class="navbar navbar-dark fixed-top bg-dark flex-md-nowrap p-0 shadow">

            <asp:Label ID="BrandLB" runat="server" Text="工段表" CssClass="navbar-brand col-sm-3 col-md-2 mr-0"></asp:Label>

        </nav>
        <div class="container-fluid">
            <div class="row">
                <nav class="col-md-2 d-none d-md-block bg-light sidebar">
                    

                        <div >
                            <h4>出貨日期</h4>
                            <div class="form-group">
                                <asp:TextBox ID="StartTB" runat="server" class="form-control"></asp:TextBox>
                                <ajaxToolkit:CalendarExtender ID="StartTB_CalendarExtender" runat="server" BehaviorID="StartTB_CalendarExtender" TargetControlID="StartTB" Format="yyyy/MM/dd" />
                                <asp:TextBox ID="EndTB" runat="server" class="form-control"></asp:TextBox>
                                <ajaxToolkit:CalendarExtender ID="EndTB_CalendarExtender" runat="server" BehaviorID="EndTB_CalendarExtender" TargetControlID="EndTB" Format="yyyy/MM/dd" />
                            </div>
                            <div class="form-group">
                                <asp:Button ID="SearchBT" runat="server" Text="Search" class="btn btn-secondary" OnClick="SearchBT_Click" />
                                <asp:Button ID="ClearBT" runat="server" Text="Clear" class="btn btn-outline-secondary" OnClick="ClearBT_Click" />

                            </div>
                            <asp:Literal ID="MessageLT" runat="server"></asp:Literal>


                        </div>
                </nav>

                <main role="main" class="col-md-9 ml-sm-auto col-lg-10 px-4">
                    <rsweb:ReportViewer ID="ReportViewer1" runat="server" Font-Names="Verdana" Font-Size="8pt" WaitMessageFont-Names="Verdana" WaitMessageFont-Size="14pt" Height="768px" Width="1024px" Visible="False">
                        <LocalReport ReportPath="GAMA\Report\工段匯入.rdlc" DisplayName="ImportEXCEL">
                        </LocalReport>
                    </rsweb:ReportViewer>
                </main>
                </div>
        </div>
    </form>
</body>
</html>
