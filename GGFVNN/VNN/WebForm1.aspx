<%@ Page Title="" Language="C#" MasterPageFile="~/VNN/VNNMaster.Master" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="GGFGAMA.VNN.WebForm1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
 <table class="table">
  <tr>
      <th>123</th>
      <td>456</td>
  </tr>
          <tr>
      <th>123</th>
      <td>456</td>
  </tr>
</table>
    </div>
        <div class="container">
  <h2>Collapsible Panel</h2>
  <p>Click on the collapsible panel to open and close it.</p>
  <div class="panel-group">
    <div class="panel panel-default">
      <div class="panel-heading">
        <h4 class="panel-title">
          <a data-toggle="collapse" href="#collapse1">Collapsible panel</a>
        </h4>
      </div>
      <div id="collapse1" class="panel-collapse collapse">
        <div class="panel-body">Panel Body</div>
        <div class="panel-footer">Panel Footer</div>
      </div>
    </div>
  </div>
</asp:Content>
