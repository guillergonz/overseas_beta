
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="Connections/overseaspr.asp" -->

<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = UCASE(CStr(Request.Form("usr")))
MM_valUserpwd = UCASE(CStr(Request.Form("pwd")))

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
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 20, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 10, MM_valUserpwd) ' adVarChar
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
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation - Login</title>

<link rel="icon" href="images/favicon.ico" type="image/x-icon" /> 
<link rel="icon" href="images/favicon.ico" type="image/x-icon" />

<link href="overseas.css" rel="stylesheet" type="text/css"><!--[if lt IE 9]>
<script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
<![endif]-->

<script type="text/javascript" src="jquery-ui/jquery-ui-1.8.16-sunny.custom/js/jquery-1.6.2.min.js"></script>
<script type="text/javascript" src="jquery-ui/jquery-ui-1.8.16-sunny.custom/js/jquery-ui-1.8.16.custom.min.js"></script>
<link type="text/css" href="jquery-ui/jquery-ui-1.8.16-sunny.custom/css/sunny/jquery-ui-1.8.16.custom.css" rel="stylesheet" />

<link rel="stylesheet" href="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/themes/base/jquery.ui.all.css">
<script src="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/jquery-1.8.2.js"></script>
<script src="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/ui/jquery.ui.core.js"></script>
<script src="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/ui/jquery.ui.widget.js"></script>
<script src="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/ui/jquery.ui.button.js"></script>
<script src="jquery-ui/jquery-ui-1.8.16-sunny.custom/development-bundle/ui/jquery.ui.datepicker.js"></script>
<link rel="stylesheet" href="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/demos/demos.css">


<script type="text/javascript" charset="utf-8">
$( "input[type=submit], input[type=button], #gobacktosource " )
		.button()
		.click(function( event ) {
			//event.preventDefault();
		});
</script>

<style type="text/css">
body {
    background-color: #FAFAFA;
    background-repeat: repeat;
}

body,td,th {
	font-family: "Segoe UI", "Segoe UI Light", "Segoe UI Semibold", "Lucida Console";
	font-size: 16px;
	padding-top: 2px;
	padding-right: 5px;
	padding-bottom: 2px;
	padding-left: 5px;
}
body {
	background-image: url();
	background-color: #000;
}
.footer {
	font-size: 9px;
	text-align: center;
	padding: 10px;
}
</style>

</head>

<body>

  <table width="90%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="35%" height="504"><form action="<%=MM_LoginAction%>" method="POST" name="validacion" id="validacion" class="borderaround_login" >
        <br>
        <center>
          <img src="images/oiclogo2.gif" alt="overseas_logo" width="256" height="31" align="top">
        </center>
        <br>
        <center>
          <label>Looking for wholesale auto part? &nbsp;&nbsp;We specialize in spare parts for Japanese and Korean cars. &nbsp;&nbsp;We offer a large selection of wholesale car parts through this easy to use website.&nbsp;&nbsp;Find wholesale parts at the lowest prices ...</label>
        </center>
        <p>
          <label>&nbsp;&nbsp;&nbsp;&nbsp;Enter a valid user id:&nbsp;
            <input name="usr" type="text" size="10" maxlength="10" class="login_fields" id="usr" />
          </label>
          <br>
          <label>Enter a valid password:&nbsp;
            <input name="pwd" size="10"  maxlength="10" type="password" class="login_fields" id="pwd" />
          </label>
        </p>
        <table cellpadding="0" cellspacing="0" border="0" width="90%">
          <tr>
            <td width="50%"><label style="color:yellow">
              <% If Request("e") = 1 Then
           Response.Write "&nbsp;&nbsp;INVALID PASSWORD!&nbsp;&nbsp;"
       End if
    %>
            </label></td>
            <td width="50%"><input name="Entrar"  type="submit"  id="Entrar" value="Login" width="100%" /></td>
          </tr>
        </table>
        <center>
          <label style="color:yellow">Contact us&nbsp;:&nbsp;(T) 787-751-4036<br>
            &nbsp;(F) 787-765-6735</label>
          <br>
          Correo Electrónico/Email: oic@overseaspr.com <br>
          Physical Address: Urb. El Paraiso<br>
          Calle Ganges #9 <br>
          San Juan P.R. 00921
        </center>
        </p>
        <center style="font-size:10px">
          Overseas Import Corporation<br>
          All rights reserved TM 2011
        </center>
      </form></td>
      <td width="65%">&nbsp;</td>
    </tr>
    <tr>
      <td>  
  <param name="movie" value="http://www.youtube.com/v/4TshFWSsrn8?version=3">
  <param name="allowFullScreen" value="true">
  <param name="allowScriptAccess" value="always">
  <embed src="http://www.youtube.com/v/4TshFWSsrn8?version=3" type="application/x-shockwave-flash" allowfullscreen="true" allowscriptaccess="always" width="219" height="202">
    </object></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <br clear="all">
  
 
</div>
    

</body>
</html>
