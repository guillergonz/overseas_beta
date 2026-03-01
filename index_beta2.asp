
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("usr"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "security_proxy.asp"
  MM_redirectLoginFailed = "index.asp?e=1"

'  MM_loginSQL = "SELECT user_name, user_pwd"
  MM_loginSQL = "SELECT user_auto_id, user_pwd"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
'  MM_loginSQL = MM_loginSQL & " FROM dbo.users WHERE user_name = ? AND user_pwd = ?"
  MM_loginSQL = MM_loginSQL & " FROM dbo.users WHERE user_auto_id = ? AND user_pwd = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_overseaspr_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
'  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 40, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 20, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 10, Request.Form("pwd")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
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

<!DOCTYPE HTML>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Overseas Import Corporation - Login</title>
<link href="overseas.css" rel="stylesheet" type="text/css" />
<style type="text/css">
body,td,th {
	color: #FFF;
}
body {
	background-color: #000;
	background-repeat: no-repeat;
	margin-top: 5px;
}
.footer {
	font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
	font-size: 9px;
	text-align: center;
	padding: 10px;
}
</style>
</head>

<body>
<table height="100%" width="109%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left"><div><img src="images/oiclogo2.gif" height="31"></div>
    </td>
    <td align="center">&nbsp;</td>
    <td height="40" align="center">&nbsp;</td>
  </tr>
  <tr>
    <td>
      <div class="oiclogo">
        <table width="250px" border="0" cellpadding="0" cellspacing="0" class="grayback" >
          <tr>
            <td ><table border="0" cellpadding="0" cellspacing="0" class="grayback">
              <tr>
                <td align="center"><p>Overseas Import Corporation <br>internet catalog.<br>
                    <br>Can't find what you are looking for ?<br>Call us at 787 751-4036<br>
                    Send us an email oic@overseaspr.com<br><% If Request("e") > 0 Then
If Request("e") = 1 Then
 Response.Write "<font class='credenciales_de_acceso'>Verificar credenciales de acceso / Check password </font>"
					End if
				End if
			    %>
                  </p></td>
                </tr>
              </table></td>
            </tr>
          <tr>
            <td align="center"><form action="<%=MM_LoginAction%>" method="POST" name="validacion" id="validacion">
              <table align = "center" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td align="right" ><div align="right"><strong>User ID :</strong></div></td>
                  <td align="left" ><label for="usr"></label>
                    <input name="usr" type="text" class="login_fields" id="usr" /></td>
                  </tr>
                <tr>
                  <td align="right"><div align="right"><strong>Password :</strong></div></td>
                  <td align="left"><input name="pwd" type="password" class="login_fields" id="pwd" /></td>
                  </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td align="right">&nbsp;&nbsp;&nbsp;&nbsp;</td>
                  <td><input name="Entrar" type="submit" class="login_fields" id="Entrar" value="Login" /></td>
                  </tr>
                <tr>
                </tr>
                </table>
              </form></td>
            </tr>
        </table>
      </div>
    </td>
    <td width="100%" height="351" rowspan="2" ><div class="youtube">
      <iframe title="YouTube video player" width="560" height="349" src="http://www.youtube.com/embed/5tFPF6M5PLo" frameborder="0" allowfullscreen></iframe>
    </div></td>
    <td width="69%" height="351" rowspan="2" align="center">&nbsp;</td>
  </tr>
  <tr>
    <td><div class="oiclogo"><img src="images/2011-Bugatti-Veyron-Super-Sport-Front-Side-Top-View.jpg" alt="Veyron" width="267" height="162"></div></td>
  </tr>
  <tr>
    <td><img src="images/Rinspeed-zaZen_Concept.jpg" width="267" height="162" class="oiclogo"></td>
    <td><img src="images/used_auto_parts.png" alt="Veyron" width="584" height="235"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td rowspan="2"><div class="footer">Overseas Import Corporation All rights reserved TM 2011</div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
