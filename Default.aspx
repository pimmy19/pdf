<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div id="employeelistDiv" runat="server">

    <table border="1">

        <tr>

            <td colspan="2">

                <b>Employee Detail</b>

            </td>

        </tr>

        <tr>

            <td><b>EmployeeID:</b></td>

            <td><asp:Label ID="lblEmployeeId" runat="server"></asp:Label></td>

        </tr>

        <tr>

            <td><b>FirstName:</b></td>

            <td><asp:Label ID="lblFirstName" runat="server"></asp:Label></td>

        </tr>

        <tr>

            <td><b>LastName:</b></td>

            <td><asp:Label ID="lblLastName" runat="server"></asp:Label></td>

        </tr>

        <tr>

            <td><b>City:</b></td>

            <td><asp:Label ID="lblCity" runat="server"></asp:Label></td>

        </tr>

           <tr>

            <td><b>Region:</b></td>

            <td><asp:Label ID="lblState" runat="server"></asp:Label></td>

        </tr>

         <tr>

            <td><b>Postal Code:</b></td>

            <td><asp:Label ID="lblPostalCode" runat="server"></asp:Label></td>

        </tr>

           <tr>

            <td><b>Country:</b></td>

            <td><asp:Label ID="lblCountry" runat="server"></asp:Label></td>

        </tr>      

    </table>

</div>
    <div>
    <asp:Button ID="btnExport" runat="server" Text="Export" OnClick="btnExport_Click" />
    </div>
    </form>
</body>
</html>
