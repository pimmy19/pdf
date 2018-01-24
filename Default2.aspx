<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default2.aspx.cs" Inherits="Default2" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        
      
      <div>
    <asp:Panel ID="pnlPerson" runat="server">
      <!--  <asp:Image ID="Image2" runat="server" /></asp:Image>-->
        
        <br />
        
    <table border="1" style="font-family: Arial; font-size: 10pt; width: 200px">
       
        <tr>
            <td><b>Name:</b></td>
            <td><asp:Label ID="lblName" runat="server"></asp:Label></td>
        </tr>
        <tr>
            <td><b>Age:</b></td>
            <td><asp:Label ID="lblAge" runat="server"></asp:Label></td>
        </tr>
        <tr>
            <td><b>City:</b></td>
            <td><asp:Label ID="lblCity" runat="server"></asp:Label></td>
        </tr>
        <tr>
            <td><b>Country:</b></td>
            <td><asp:Label ID="lblCountry" runat="server"></asp:Label></td>
        </tr>
    </table>
</asp:Panel>
<asp:Button ID="btnExport" runat="server" Text="Export" OnClick="btnExport_Click" />
    </div>
    </form>
</body>
</html>
