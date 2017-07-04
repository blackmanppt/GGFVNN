<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="LoginTest.aspx.cs" Inherits="GGFVNN.LoginTest" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">

    </style>
</head>
<body >
    <form id="form1" runat="server" >
        <div>
        </div>
        <div style="margin: 0px auto;">
            <table>
                <tr>
                    <td><img src="IMG/GGF.gif" style="height: 110px; width: 217px" /></td>
                    <td>
                        <asp:Login ID="Login" runat="server" BackColor="#EFF3FB" BorderColor="#B5C7DE" BorderPadding="4" BorderStyle="Solid" BorderWidth="1px" FailureText="Login faire" Font-Names="Verdana" Font-Size="0.8em" ForeColor="#333333" Height="107px" LoginButtonText="Login" PasswordLabelText="Password:" PasswordRequiredErrorMessage="Password" RememberMeText="Remember Password" TitleText="" UserNameLabelText="User Name:" UserNameRequiredErrorMessage="User Name" Width="315px" OnAuthenticate="Login_Authenticate">
                            <CheckBoxStyle Font-Size="Larger" />
                            <HyperLinkStyle Font-Size="Larger" />
                            <InstructionTextStyle Font-Italic="True" Font-Size="Larger" ForeColor="Black" />
                            <LabelStyle Font-Size="Larger" />
                            <FailureTextStyle Font-Size="Larger" />
                            <LoginButtonStyle BackColor="White" BorderColor="#507CD1" BorderStyle="Solid" BorderWidth="1px" Font-Names="Verdana" Font-Size="Larger" ForeColor="#284E98" />
                            <TextBoxStyle Font-Size="Larger" />
                            <TitleTextStyle BackColor="#507CD1" Font-Bold="True" Font-Size="0.9em" ForeColor="White" />
                        </asp:Login>
                    </td>
                </tr>
            </table>




        </div>
    </form>
</body>
</html>
