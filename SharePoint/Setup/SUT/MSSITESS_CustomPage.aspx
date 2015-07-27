<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Page language="C#" %>
<%@ Register tagprefix="SharePoint" namespace="Microsoft.SharePoint.WebControls" assembly="Microsoft.SharePoint, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>TestPage</title>

<%@Import Namespace="Microsoft.SharePoint"%>
<script language="c#" runat="server" type="text/c#">
    protected void Button1_Click(object sender, EventArgs e)
    {
        try
        {
            SPContext.Current.Web.Lists["MSSITESSList"].Delete();
        }
        catch { }
        Guid list = SPContext.Current.Web.Lists.Add("MSSITESSList", "MSSITESS form digest test", SPListTemplateType.DocumentLibrary);        
        SPContext.Current.Web.Lists["MSSITESSList"].Delete();
        
        Response.Write("Your input is" + TextBox1.Text + "." );
        Response.End();
    }
</script>


</head>

<body>
    <form id="form1" runat="server">
    <div>
        <asp:Label ID="Label1" runat="server" Text="InputText"></asp:Label>
        <asp:TextBox ID="TextBox1" runat="server">abc</asp:TextBox>
        <asp:Button ID="Button1" runat="server" Text="Submit" OnClick="Button1_Click"/>
    </div>
    <SharePoint:FormDigest runat="server" id="FormDigest1"/>
    </form>
</body>
</html>