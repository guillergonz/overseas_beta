<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="validate_access.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML><!-- InstanceBegin template="/Templates/index.dwt.asp" codeOutsideHTMLIsLocked="false" --><HEAD>
<!-- InstanceBeginEditable name="doctitle" -->
<TITLE>Online Inventory 2007</TITLE>
<!-- InstanceEndEditable -->
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style4 {font-size: 12px}
-->
</style>
<!-- InstanceEndEditable -->
<link href="../fonts.css" rel="stylesheet" type="text/css">

<script src="../SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<script language="JavaScript">
<!--
function mmLoadMenus() {
  if (window.mm_menu_0708133328_0) return;
                window.mm_menu_0708133328_0 = new Menu("root",128,17,"Arial, Helvetica, sans-serif",11,"#000000","#FFFFFF","#CCCCCC","#000084","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0708133328_0.addMenuItem("Search&nbsp;Parts","location='search_parts.asp'");
  mm_menu_0708133328_0.addMenuItem("Parts&nbsp;Request","location='parts_request.asp'");
  mm_menu_0708133328_0.addMenuItem("Business&nbsp;Partners","location='business_partners.asp'");
   mm_menu_0708133328_0.hideOnMouseOut=true;
   mm_menu_0708133328_0.bgColor='#555555';
   mm_menu_0708133328_0.menuBorder=1;
   mm_menu_0708133328_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0708133328_0.menuBorderBgColor='#777777';
  window.mm_menu_0708134000_0 = new Menu("root",113,17,"Arial, Helvetica, sans-serif",11,"#000000","#FFFFFF","#CCCCCC","#000084","left","middle",3,0,1000,-5,7,true,true,true,0,true,true);
  mm_menu_0708134000_0.addMenuItem("My&nbsp;Cart","location='my_cart.asp'");
  mm_menu_0708134000_0.addMenuItem("My&nbsp;Hot&nbsp;Parts","location='my_hot_parts.asp'");
  mm_menu_0708134000_0.addMenuItem("Order&nbsp;Status","location='my_order_status.asp'");
  mm_menu_0708134000_0.addMenuItem("Orders&nbsp;History","location='my_orders_history.asp'");
   mm_menu_0708134000_0.hideOnMouseOut=true;
   mm_menu_0708134000_0.bgColor='#555555';
   mm_menu_0708134000_0.menuBorder=1;
   mm_menu_0708134000_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0708134000_0.menuBorderBgColor='#777777';

mm_menu_0708134000_0.writeMenus();
} // mmLoadMenus()
//-->
</script>
<script language="JavaScript" src="mm_menu.js"></script>
<style type="text/css">
<!--
.style6 {
	font-family: "Times New Roman", Times, serif;
	font-size: 10px;
	font-style: italic;
}
-->
</style>

<style type="text/css">
<!--
body {
	background-image: url(theme2MENU.png);
	background-repeat: repeat-y;
}
-->
</style>
<link href="../SpryAssets/SpryMenuBarHorizontal.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style7 {font-size: small}
.style9 {font-size: small; color: #0000FF; }
-->
</style>
</HEAD>
<BODY aLink=#000000 leftMargin=0 link=#000000  
onload="if (self != top) top.location = self.location;" text=#000000 topMargin=0 
vLink=#496989>
<script language="JavaScript1.2">mmLoadMenus();</script>
<TABLE border=0 cellPadding=0 cellSpacing=0 class="style6">
  <TBODY>
  <TR>
    <TD height="39" align=left valign="top"></div>
      <table width="770" border="0" bordercolor="#FFFFFF">
        <tr>
          <td width="150" class="header1"></td>
          <td height="17" class="header1"><div align="left"><img src="oiclogo2.gif" width="256" height="31"></div></td>
          </tr>
        <tr>
          <td height="25" class="big2">&nbsp;</td>
		  <td width="614"><ul id="MenuBar1" class="MenuBarHorizontal">
                <li><a href="../index.asp">Home</a></li>
                <li><a href="../Catalog/index.asp">Soportes</a></li> 
                <li><a href="../search2.asp">Specials</a> </li>
                <!--<li><a href="../search_parts.asp">Search Catalog</a></li> -->
				<li><a href="../search_parts_beta.asp">Search Catalog</a></li>
                <!--<li>
                  <div align="center"><a href="../search_parts_beta.asp">Advanced Search </a><span class="style9">New!</span> 
                    </div> 
                    </div>
                </li>	-->	
                <li><a href="../account_statement.asp">Account Statement</a>                </li>
            <li><a href=""> Log Out </a></li>
          </ul>  		    </td>
  		  </tr>
      </table>      </TD>
  </TR></TBODY></TABLE>
<TABLE width=100% border=0 cellPadding=0 cellSpacing=0>
  <TBODY>
  <TR vAlign=top>
    <TD width="15%" rowspan="2" align="left" vAlign=top>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width=151>
        <TBODY>
        <TR>
          <TD width=170>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width=150>
              <TBODY>
              <TR>
                <TD colSpan=3><FONT face=Verdana size=1><img 
                  border=0 height=29 src="blah2.gif" 
                  width=150></FONT></TD>
              </TR>
              <TR>
                <TD width=17 height="56" background=blah4.gif></TD>
                <TD bgColor=#ffffff vAlign=top width=115><FONT face=Verdana 
                  size=1><IMG height=9 src="arrow.gif" width=9>                  <FONT class=content color=#363636><B>Main Menu</B><BR>
                        <IMG 
                  height=8 src="bulletb.gif" width=8></FONT></FONT> <span class="style7"><FONT face=Verdana><FONT color=#363636><a href="http://www.overseaspr.com/index.asp" class="big2">Home</a><BR>
                        <IMG height=8 src="bulletb.gif" 
                  width=8> <a href="http://www.overseaspr.com/logout.asp" class="big2">Logout</a> <BR>
                        <IMG height=8 src="bulletb.gif" 
                  width=8> </FONT></FONT></span><span class="big2"><FONT face=Verdana><FONT color=#363636><a href="http://www.overseaspr.com/my_account.asp">My Account</a></FONT></FONT> </span></TD>
                <TD background=blah3.gif width=19></TD></TR>
              <TR>
                <TD colSpan=3><FONT face=Verdana size=1><IMG 
                  border=0 height=21 src="blah5.gif" 
                  width=150></FONT></TD></TR></TBODY>
            </TABLE>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width=138>
              <TBODY>
              <TR>
                <TD colSpan=3 width=150><FONT face=Verdana size=1><IMG 
                  border=0 height=29 src="blah2.gif" 
                  width=150></FONT></TD></TR>
              <TR>
                <TD background=blah4.gif width=17></TD>
                <TD bgColor=#ffffff vAlign=top width=114><FONT face=Verdana 
                  size=1><IMG height=9 src="arrow.gif" width=9>                  <FONT class=content 
                  color=#363636><B>Features</B><BR>
                        </FONT></FONT><FONT 
                  class=content color=#363636 face=Verdana size=1><IMG 
                  height=8 src="bulletb.gif" width=8> </FONT> <span class="style7"><FONT color="#363636" face=Verdana><a href="../search_parts_beta.asp" class="big2">Search Parts</a><BR>
                        <IMG 
                  height=8 src="bulletb.gif" width=8> </FONT></span><span class="big2"><FONT color="#363636" face=Verdana><a href="http://www.overseaspr.com/parts_request.asp"> Request</a><A href="#"></A></FONT></span><span class="style7"><FONT color="#363636" face=Verdana><A href="#"><BR>
                        </A><IMG height=8 src="bulletb.gif" 
                  width=8> <a href="http://www.overseaspr.com/business_partners.asp" class="big2">Business</a></FONT></span></TD>
                <TD background=blah3.gif width=19></TD></TR>
              <TR>
                <TD colSpan=3 width=150><FONT face=Verdana size=1><IMG 
                  border=0 height=21 src="blah5.gif" 
                  width=150></FONT></TD></TR></TBODY></TABLE>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width=147>
              <TBODY>
              <TR>
                <TD colSpan=3 width=150><FONT face=Verdana size=1><IMG 
                  border=0 height=29 src="blah2.gif" 
                  width=150></FONT></TD>
              </TR>
              <TR>
                <TD background=blah4.gif width=17></TD>
                <TD bgColor=#ffffff vAlign=top width=114><FONT face=Verdana 
                  size=1><IMG height=9 
                  src="arrow(1).gif" 
                  width=9> <FONT class=content 
                  color=#363636><B>Shopping Cart </B><BR>
                        <IMG height=8 
                  src="bulletb(1).gif" 
                  width=8></FONT></FONT><FONT face=Verdana> <span class="style7"><FONT color=#000080><a href="http://www.overseaspr.com/my_cart.asp" class="big2">My Cart </a><br>
                        </FONT>
                        <FONT 
                  class=content color=#363636><IMG height=8 
                  src="bulletb(1).gif" 
                  width=8> </FONT><FONT 
                  class=content color=#363636 face=Verdana><a href="http://www.overseaspr.com/my_order_status.asp" class="big2">Order Status</a> </FONT><FONT 
                  class=content color=#363636><BR>
                        <IMG height=8 
                  src="bulletb(1).gif" 
                  width=8></FONT></span> <span class="big2"><FONT color=#363636 face=Verdana><a href="http://www.overseaspr.com/my_orders_history.asp">History</a></FONT></span></FONT></TD>
                <TD background=blah3.gif width=19></TD></TR>
              <TR>
                <TD colSpan=3 width=150><FONT face=Verdana size=1><IMG 
                  border=0 height=21 src="blah5.gif" 
                  width=150></FONT></TD></TR>
              <TR>
                <TD colSpan=3 width=150><FONT face=Verdana size=1><BR><IMG 
                  border=0 height=29 src="blah2.gif" 
                  width=150></FONT></TD></TR>
              <TR>
                <TD background=blah4.gif width=17></TD>
                      <TD bgColor=#ffffff vAlign=top width=114><FONT face=Verdana 
                  size=1>
				  
				  <FONT face=Verdana 
                  size=1><IMG 
                  src="arrow(1).gif" alt="a" 
                  width=9 height=9></FONT><FONT 
                  class=content color=#363636>
                        <B>Your Cart!</B><br>
                        </FONT></FONT><FONT face=Verdana><font face=Verdana><font color=#363636>
                        <%
   dim retVal,browser, connStr,dwMine_action, dwMine_context, selfLink, selfLinkArgs, statement
   dim rc, mod_string, original_select, where_clause, client_code
   client_code = Session("MM_Username")
   set dwMine = Server.CreateObject("PowerBuilder.HTMLDataWindow")
   dwMine.SetDWObject "C:\\apps\\zorrilla\\zorrilla\\pb_workspace\\com_framework.pbl", "dw_cart_resume"
   connStr = "ConnectString='DSN=zorrilla;UID=dba;PWD=sql',ConnectOption='SQL_DRIVER_CONNECT,SQL_DRIVER_NOPROMPT'"
   dwMine.setTrans "ODBC", connStr, "", "", "", "", ""
   dwMine.retrieveex (client_code)
   dwMine.SetHTMLObjectName("dwMine")
   dwMine_action = Request.Form("dwMine_action")
   dwMine_context = Request.Form("dwMine_context")
   if dwMine_action <> "undefined" then
       if dwMine.SetAction (dwMine_action, dwMine_context) < 0  then
		  Response.Write ("Error on SetAction(): " + retVal + dwMine.GetLastErrorString() + "<br>")
       end if
   end if
   if dwMine.RowCount() > 0 then 
'	   if dwMine.RowCount() > 3 then
		   Response.Write dwMine.Generate + "<br><br><br><br><br><br><br><br><br><br><br><br>"
		   '<FONT class=content face=Verdana size=1 color=red>Loggedon: " + client_code   + "</FONT>"
'	   else
'		   Response.Write dwMine.Generate + "<br>" + client_code "<br><br><br><br><br><br><br><br><br>"
'	   end if
   else
   	   Response.Write " </FONT><FONT class=content face=Verdana size=1 color=red><br>Your Cart Is Empty"
   end if
   set dwMine = nothing
%>
                        </font></font></FONT></TD>
                    <TD background=blah3.gif width=19></TD></TR>
              <TR>
                <TD colSpan=3 width=150><FONT face=Verdana size=1><IMG 
                  border=0 height=21 src="blah5.gif" 
                  width=150></FONT></TD></TR></TBODY></TABLE>
      </TD></TR></TBODY></TABLE></TD>
    <TD width="85%" align="left">      
        <!-- InstanceBeginEditable name="main_page" -->
<script>
function obj_dw1_ButtonClicked(row, objName)
{
	if (objName == "b_add")
	{
		alert(objName + "\n" + obj_dw1.GetItem(row, 1) + "\n" + row);
		//obj_dw1_submitForm.obj_dw1_action.value = "InsertRow" ;
		//obj_dw1_submitForm.obj_dw1_context.value = obj_dw1_submitForm.obj_dw1_context.value ;
	}
}
</script>	
<table class="short"><tr><td>
<%
dim resultado,prt1,prt2,prt3,prt4,prt5,prt6,prt7,prt8,prt9,prt10
dim qty1,qty2,qty3,qty4,qty5,qty6,qty7,qty8,qty9,qty10
dim dw_context, dw_action
dim user_id

user_id = Session("MM_Username")

prt1 = UCase(Request.Form("part1"))

if len(prt1) = 0 then
	prt1 = Request.Form("arg_part")
end if

qty1 = UCase(Request.Form("part1_qty"))

dw_context = Request.Form("obj_dw1_context")
dw_aciton = Request.Form("obj_dw1_action")

set zorrilla = server.CreateObject("zorrilla.framework")
resultado = zorrilla.f_retrieve2(user_id,client_code,dw_action,dw_context,prt1,qty1)
response.Write resultado + "<br><Br><br><br><br><br><br>" + user_id + "<br>" + client_code + "<Br>" + dw_action + "<br>" + dw_context + "<br>" + prt1 + "<br>" + qty1
set zorrilla = nothing
resultado = ""
%>
</td></tr></table>
		<!-- InstanceEndEditable -->
<!--         <div align="center"><img src="../images/prius-starter.jpg" alt="" width="450" height="302"><br>
          <br>      
        </div> -->   
  </TR>
  <TR vAlign=top>
    <TD align="center">    
  </TR>
</TBODY></TABLE>

<script type="text/javascript">
<!--
var MenuBar1 = new Spry.Widget.MenuBar("MenuBar1", {imgDown:"../SpryAssets/SpryMenuBarDownHover.gif", imgRight:"../SpryAssets/SpryMenuBarRightHover.gif"});
//-->
</script>
</BODY><!-- InstanceEnd --></HTML>L>
