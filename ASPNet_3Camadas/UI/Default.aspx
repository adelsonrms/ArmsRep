<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="UI.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .auto-style2 {
            height: 24px;
            width: 498px;
        }
        .auto-style3 {
            text-align: right;
            width: 127px;
        }
        .auto-style4 {
            height: 24px;
            text-align: right;
            width: 127px;
        }
        .auto-style6 {
            text-align: right;
            width: 127px;
            height: 23px;
        }
        .newStyle1 {
            font-family: "Segoe UI Light";
            width: 100%;
        }
        .auto-style7 {
            height: 23px;
            width: 498px;
        }
        .auto-style8 {
            text-align: right;
            width: 127px;
            height: 25px;
        }
        .auto-style9 {
            height: 25px;
            width: 498px;
        }
        .auto-style10 {
            color: #FFFFCC;
            background-color: #6666FF;
        }
        .auto-style11 {
            width: 498px;
        }
        .auto-style13 {
            text-align: right;
            background-color: #009999;
        }
        .auto-style14 {
            text-align: right;
            width: 127px;
            height: 34px;
        }
        .auto-style15 {
            width: 498px;
            height: 34px;
        }
        .auto-style16 {
            font-family: "Segoe UI Light";
            width: 100%;
            height: 630px;
        }
        .auto-style17 {
            text-align: right;
            width: 127px;
            height: 165px;
        }
        .auto-style18 {
            width: 498px;
            height: 165px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <table class="auto-style16">
            <tr>
                <td class="auto-style10" colspan="2"><strong>ARMSTECH - CADASTRO DE CLIENTES</strong></td>
            </tr>
            <tr>
                <td class="auto-style3"><strong>Código&nbsp; : </strong></td>
                <td class="auto-style11">
                    <asp:TextBox ID="txtID" runat="server" style="margin-left: 6px" Width="148px"></asp:TextBox>
                    <asp:Button ID="btnLocalizar" runat="server" Text="Button" Width="138px" OnClick="btnLocalizar_Click" />
                </td>
            </tr>
            <tr>
                <td class="auto-style4"><strong>Nome : </strong></td>
                <td class="auto-style2">
                    <asp:TextBox ID="txtNome" runat="server" style="margin-left: 6px" Width="906px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="auto-style6"><strong>Endereço : </strong></td>
                <td class="auto-style7">
                    <asp:TextBox ID="txtEndereco" runat="server" style="margin-left: 6px" Width="906px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="auto-style8"><strong>Telefone : </strong></td>
                <td class="auto-style9">
                    <asp:TextBox ID="txtTelefone" runat="server" style="margin-left: 6px" Width="147px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="auto-style4"><strong>Email : </strong></td>
                <td class="auto-style11">
                    <asp:TextBox ID="txtEmail" runat="server" style="margin-left: 6px" Width="906px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="auto-style6"><strong>Observações : </strong></td>
                <td class="auto-style11">
                    <asp:TextBox ID="txtObservacao" runat="server" style="margin-left: 6px" Width="906px" Height="63px" TextMode="MultiLine"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="auto-style13" colspan="2">&nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style3">&nbsp;</td>
                <td class="auto-style11">
                    <asp:GridView ID="gvClientes" runat="server" BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="4" ForeColor="Black" GridLines="Horizontal" Height="223px" Width="917px">
                        <FooterStyle BackColor="#CCCC99" ForeColor="Black" />
                        <HeaderStyle BackColor="#333333" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="White" ForeColor="Black" HorizontalAlign="Right" />
                        <SelectedRowStyle BackColor="#CC3333" Font-Bold="True" ForeColor="White" />
                        <SortedAscendingCellStyle BackColor="#F7F7F7" />
                        <SortedAscendingHeaderStyle BackColor="#4B4B4B" />
                        <SortedDescendingCellStyle BackColor="#E5E5E5" />
                        <SortedDescendingHeaderStyle BackColor="#242121" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td class="auto-style13" colspan="2">&nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style14"></td>
                <td class="auto-style15">
                    <asp:Button ID="btnIncluir" runat="server" Text="Incluir" Width="138px" />
                &nbsp;<asp:Button ID="btnAlterar" runat="server" Text="Alterar" Width="138px" />
                &nbsp;<asp:Button ID="btnExcluir" runat="server" Text="Excluir" Width="138px" />
                </td>
            </tr>
            <tr>
                <td class="auto-style14">Informações : </td>
                <td class="auto-style15">
                    <asp:Label ID="lblmsg" runat="server" Text="Label"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="auto-style17">Detalhes&nbsp; Adicionais</td>
                <td class="auto-style18">
                    <asp:Label ID="lblDetalhes" runat="server" Text="Label"></asp:Label>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
