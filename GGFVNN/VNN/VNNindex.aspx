<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="VNNindex.aspx.cs" Inherits="GGFGAMA.VNN.VNNindex" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <%--<link href="../Scripts/bootstraptemp.css" rel="stylesheet" />--%>
    
    
    <link href="../Scripts/style.css" rel="stylesheet" />
    <link href="../Scripts/bootstrap.min.css" rel="stylesheet" />
   <script src="../Scripts/jquery.min.js"></script>
    <script src="../Scripts/scripts.js"></script>
    <script src="../Scripts/bootstrap.min.js"></script>
   <%--  <link href="../Scripts/bootstrap-theme.min.css" rel="stylesheet" />--%>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table class="table">
                <tr>
      <th><h3 class="text-success">越南程式首頁</h3></th>
        </tr>
        <tr>
      <th>
          <asp:HyperLink ID="HyperLink1" runat="server" CssClass="btn btn-default btn-primary" NavigateUrl="~/VNN/Ship001.aspx">越南出口大表</asp:HyperLink>

      </th>
      
  </tr>

</table>
    </div>
    
    </form>
</body>
</html>
