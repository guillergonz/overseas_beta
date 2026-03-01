
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
</head>

<body>
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><table width="400" border="0" cellpadding="0" cellspacing="0" bgcolor="#F0F0F0" >
      <tr>
        <td width="398px" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td width="50%" align="center"><div class="image_placement">
              <h1><strong><img src="images/oiclogo2.gif" width="256" height="31" class="image_placement"></strong></h1>
            </div>
              <h2><strong class="tablas_font12">LOGIN</strong><br>
                <span>
                <% 
				If Request("e") > 0 Then
					If Request("e") = 1 Then
						Response.Write "<font class='credenciales_de_acceso'>Verificar credenciales de acceso</font>"
					End if
				End if
			    %>
                </span> <br />
              </h2></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td align="center"><form action="<%=MM_LoginAction%>" method="POST" name="validacion" id="validacion">
          <table width="400" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="40%" align="right" class="tablas_font12"><div align="right"><strong>USUARIO</strong>:</div></td>
              <td width="60%" align="left"><label for="usr"></label>
                <input name="usr" type="text" class="login_fields" id="usr" /></td>
            </tr>
            <tr>
              <td align="right" class="tablas_font12"></td>
              <td align="left">&nbsp;</td>
            </tr>
            <tr>
              <td align="right" class="tablas_font12"><div align="right"><strong>CONTRASEÑA</strong>:</div></td>
              <td align="left"><input name="pwd" type="password" class="login_fields" id="pwd" /></td>
            </tr>
            <tr>
              <td align="right" class="tablas_font12">&nbsp;</td>
              <td align="left">&nbsp;</td>
            </tr>
            <tr>
              <td align="right">&nbsp;</td>
              <td align="left"><input type="submit" name="Entrar" id="Entrar" value="Entrar" /></td>
            </tr>
          </table>
        </form></td>
      </tr>
      <tr>
        <td align="center">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
