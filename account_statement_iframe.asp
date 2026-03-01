<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
	Response.Expires = -1
	valid = false
	psw = Trim(Request("psw"))
	if Len(psw) > 0 then
		  Set oRS2 = Server.CreateObject ("ADODB.Command")
		  oRS2.ActiveConnection = MM_overseaspr_STRING
		  oRS2.CommandText = "SELECT accountStatementPassword FROM dbo.Users WHERE user_auto_id = '" + Session("MM_Username") + "' AND accountStatementPassword = '" + psw + "'"
		  oRS2.Prepared = true
		 Set checker = oRS2.Execute
		 
		 if Not checker.EOF then
			if trim(checker.Fields.Item("accountStatementPassword").Value) = psw Then
				valid = true
			end if
		 end if
	end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Overseas Import Corporation - Order Detail</title>
<link href="overseas.css" rel="stylesheet" type="text/css" />

<style type="text/css">
body,td,th {
	font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
	font-size: 12px;
}
</style>


</head>

<body>



<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="index.asp"
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

<div class="display_no" id="loader"><img src="images/ajax-loader.gif" width="16"  /></div>

<form id="order_submit" method="post" action="/account_statement_iframe.asp">



<% If not valid and (Session("MM_Username") = "B004" or Session("MM_Username") = "Z099") then %>
<div class="well">

	<p>
	<label>Please enter password to access account statement</label>
	<input type="password" id="psw" name="psw" pattern="(?=.*\d)(?=.*[a-z])(?=.*[A-Z]).{8,}" title="Must contain at least one number and one uppercase and lowercase letter, and at least 8 or more characters" required>

    <input type="submit" value="Submit">
	</p>
	
</div>
<% else %>
<table width='100%' border="0" cellpadding="0" cellspacing="0" >
  <tr>
   <td align="center"><div class="display_yes" id="shopping_cart">
       <% AccountStatementDisplay() %>
	 </div>
      
      <p>&nbsp;</p></td>
  </tr>
</table>
<% end if %>

</form>
</body>
</html>
