<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/zorrilla.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername=UCase(CStr(Request.Form("u_id")))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization="user_priority"
  MM_redirectLoginSuccess="index.asp"
  MM_redirectLoginFailed="incorrect_information.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_zorrilla_STRING
  'MM_rsUser.Source = "SELECT user_id, user_pwd"
  MM_rsUser.Source = "SELECT user_id, user_pwd, user_client_code, user_last_visit, user_auto_id, user_email"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM DBA.users WHERE user_id='" & Replace(MM_valUsername,"'","''") &"' AND user_pwd='" & Replace(Request.Form("pwd"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
	  Session("MM_UserClientCode") = CStr(MM_rsUser.Fields.Item("user_client_code").Value)
	  Session("MM_UserUnicID") = CStr(MM_rsUser.Fields.Item("user_auto_id").Value)
	  Session("MM_UserEmail") = CStr(MM_rsUser.Fields.Item("user_email").Value)

	  Dim objConn, strConnection
	  Set objConn = Server.CreateObject("ADODB.Connection")
	  strConnection = MM_zorrilla_STRING
	  objConn.Open strConnection
	  ls_months = DateDiff("M", CDate(MM_rsUser.Fields.Item("user_last_visit").Value), Date)
	  ls_pwd = CStr(MM_rsUser.Fields.Item("user_pwd").Value)

		<!-- Do this before redirecting to resetpwd.asp to update user_last_visit -->
	  objConn.Execute "UPDATE users SET user_last_visit = '" + CStr(Now()) + "' WHERE (user_id = '" + MM_valUsername + "') AND (user_client_code = '" + CStr(MM_rsUser.Fields.Item("user_client_code").Value) + "') ;"
	  objConn.Execute "INSERT INTO user_audit (userid,lastchange,mode_type) VALUES ('" + MM_valUsername + "','" + CStr(Now()) + "','1');"
	  objConn.Close()
	  Set objConn = Nothing

	  if ls_months > "1" or ls_pwd = "11111" then
		session("change_pwd") = "true"
	  	response.Redirect("resetpwd.asp")
	  end if
	  
	  if MM_rsUser.Fields.Item("user_email").Value = "oic@overseasimport.com" then 
	  	response.Write(MM_UserEmail)
	  	response.redirect("my_account.asp")
	  end if
	  
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/special_messages.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>

<!-- InstanceBeginEditable name="doctitle" -->
<title>Inventory Manager 2008 All Rights Reserved ITEC Inc.</title>
<!-- InstanceEndEditable --><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="fonts.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="head" --><style type="text/css">
<!--
body {
	background-image: url();
	background-repeat: repeat-x;
}
.style18 {font-size: 18px; color: #000000; font-family: Tahoma, fantasy;}
.style22 {
	font-size: 18px;
	font-weight: bold;
	font-style: italic;
}
.style25 {color: #CC0000}
-->
</style><!-- InstanceEndEditable -->
<style type="text/css">
<!--
.style16 {font-family: Tahoma, fantasy}
-->
</style>
</head>

<body topmargin="0" leftmargin="0">
<TABLE border=0 cellPadding=0 cellSpacing=0 height=17 width=1430>
  <TBODY>
    <TR>
      <TD width="1430" height=1 bgColor=#b8babc>
        <P align=right> <IMG border=0 height=5 src="images/greybit.jpg"
      width=34></P></TD>
    </TR>
  </TBODY>
</TABLE> 
<!-- InstanceBeginEditable name="main_message" --><div align="left">
  <p align="center" class="style18 ">Welcome to Overseas Imports Corporation's internet ordering System.</p>
  <p align="center" class="style18 style25"><em><strong>787-751-4036</strong></em></p>
</div>

  
    
      <div align="center">
        <table width="275" height="205" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="187" align="center" valign="top" nowrap>	<form ACTION="<%=MM_LoginAction%>" METHOD="POST" name="user_validation" >
              <div align="center">
	            <table width="248" height="189" border="0" align="center" class="tables_top">
                    <tr>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr background="images/blah4.gif">
                      <td colspan="3"><TABLE width=150 border=0 align="center" cellPadding=0 cellSpacing=0 class="short">
                        <TBODY>
                          <TR>
                            <TD colSpan=4><font size=1><img
                  border=0 height=29 src="images/blah2.gif"
                  width=150></font></TD>
                          </TR>
                          <TR>
                            <TD height="20" background=images/blah4.gif></TD>
                            <TD colspan="2" vAlign=top bgColor=#ffffff><font
                  size=1><img height=9 src="images/arrow.gif" width=9></font> <font color=#363636 class=big1 style19><b>Login</b></font></TD>
                            <TD background=images/blah3.gif></TD>
                          </TR>
                          <TR>
                            <TD height="19" background=images/blah4.gif></TD>
                            <TD bgColor=#ffffff vAlign=top>&nbsp;</TD>
                            <TD bgColor=#ffffff vAlign=top>&nbsp;</TD>
                            <TD background=images/blah3.gif></TD>
                          </TR>
                          <TR>
                            <TD height="19" background=images/blah4.gif></TD>
                            <TD bgColor=#ffffff vAlign=top><font
                  size=1><font class=short color=#363636> User Id </font></font></TD>
                            <TD bgColor=#ffffff vAlign=top><input name="u_id" type="text" class="parts" id="u_id2" style="width:50px;text-transform:uppercase"></TD>
                            <TD background=images/blah3.gif></TD>
                          </TR>
                          <TR>
                            <TD height="21" background=images/blah4.gif></TD>
                            <TD vAlign=top bgColor=#ffffff><font
                  size=1><font class=short color=#363636> Password</font></font></TD>
                            <TD vAlign=top bgColor=#ffffff><input name="pwd" type="password" class="parts" id="pwd" style="width:50px" ></TD>
                            <TD background=images/blah3.gif></TD>
                          </TR>
                          <TR>
                            <TD width=17 height="19" background=images/blah4.gif></TD>
                            <TD bgColor=#ffffff vAlign=top width=60><FONT
                  size=1><FONT class=content color=#363636> <BR>
                            </FONT></FONT></TD>
                            <TD bgColor=#ffffff vAlign=top width=54><input name="Submit" type="submit" class="short" value="Search"></TD>
                            <TD background=images/blah3.gif width=19></TD>
                          </TR>
                          <TR>
                            <TD colSpan=4><FONT size=1><IMG
                  border=0 height=21 src="images/blah5.gif"
                  width=150></FONT></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="3">&nbsp;</td>
                    </tr>
                </table>
	            <p class="style22">&nbsp;</p>
              </div>
	          </form>	</td>
          </tr>
        </table>
      </div>
      <table width="608" border="0" align="center">
   <tr>
     <td colspan="2"><span class="style18 style17 style16"></span></td>
   </tr>
   <tr>
     <td width="602" colspan="2"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="600" height="400">
       <param name="movie" value="thumbnail_scroller.swf">
       <param name=quality value=high>
       <embed src="thumbnail_scroller.swf" width="600" height="400" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed>
     </object></td>
   </tr>
</table>
 <form ACTION="<%=MM_LoginAction%>" METHOD="POST" name="user_validation" >
   <div align="center">
     <p class="tables_top">Designed for Windows Internet Explorer</p>
   </div>
 </form>
<p>&nbsp;</p>
<!-- InstanceEndEditable -->
</body>
<!-- InstanceEnd --></html>
