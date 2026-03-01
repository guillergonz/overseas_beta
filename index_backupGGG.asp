<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
Response.Expires = -1

Set oRS = Server.CreateObject("ADODB.Recordset")
' Check websote on maintenance
strSQL = "SELECT forceinventoryupd FROM dbo.counters"
oRS.Open strSQL, MM_overseaspr_STRING
If Not oRS.EOF AND Not IsNull(oRS("forceinventoryupd")) Then
	onmaintenance = oRS("forceinventoryupd")	
Else
	onmaintenance = "0"
End If		
		
		
		
Session("MM_Multi_Username") = ""

' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = UCASE(CStr(Request.Form("usr")))
MM_valUserpwd = UCASE(CStr(Request.Form("pwd")))

If MM_valUsername <> "" Then

  	if Not MM_valUsername = "Z099" AND 	onmaintenance = "1" then
		Response.Redirect("index.asp?maint=1")
	end if		
		
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
  MM_loginSQL = MM_loginSQL & " FROM dbo.users WHERE user_auto_id = ? AND user_pwd = ? AND activestatus = 'A' "
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
	
	
	if MM_valUsername = "B001" then
		Response.Redirect("multilogin.asp")
	else
		Response.Redirect(MM_redirectLoginSuccess)
	end if
    
	
	
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>

<!doctype html>
<head>
<meta charset="utf-8"> 
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" >
<link rel="icon" href="images/favicon.ico" type="image/x-icon" >
<title>Overseas Import Corporation</title>

<!-- Bootstrap -->
<link rel="stylesheet" type="text/css" href="bootstrap-3.3.6-dist/css/bootstrap.min.css">

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script type="text/javascript" charset="utf-8" src="bootstrap-3.3.6-dist/jquery.min.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->

<link href="overseas.css" rel="stylesheet" type="text/css" >

<style type="text/css">
body {
	margin:0;
	background-color: #FFF;
	background-repeat: repeat;
}
.input focus{
	border: #FC0;
}
.red {
	color:red;
}
body,td,th {
font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
font-size: 14px;
color: #000;
}
</style>

	
    
<script type="application/javascript">

//if($("#usr").val() == "B001") {
//	window.location.href = "multilogin.asp";
//}
			
$(document).ready(function() {

	$("#usr").val("");
	$("#pwd").val("");
	
	$("#effect").css("display", "inline");

	$("#usr").change(function() {
		var usr = $("#usr").val(); 
		if(usr == 'B001' || usr == 'B0011' || usr == 'B0012' || usr == 'B0013' || usr == 'B0014') {
			window.location.href = "multilogin.asp";
		}
	});
		
});
</script>
</head>

<body  >

<div class="container"  >
	
    <div class="jumbotron" style="padding:20px">
	    <img src="images/oiclogo2.gif" alt="overseas_logo" width="256" height="31" align="top">
        <h3>Welcome<br>
            Looking for wholesale auto parts? &nbsp;
            We specialize in spare parts for Japanese and Korean cars. We offer a large selection of wholesale car parts through this easy to use website.&nbsp;Find wholesale parts at the lowest prices ...
            </h3>
    </div>
    
    <form class="form-horizontal" action="<%=MM_LoginAction%>" method="POST" name="validacion" id="validacion" >
    
	<div class="row">
    
		<div class="col-lg-6 col-xs-12">	
                         
            <div class="form-group">
                <label for="useridl" class="control-label col-xs-4 col-xs-offset-1">User ID</label>
                <div class="col-xs-6">
                    <input  name="usr" type="text" maxlength="10" class="form-control" id="usr" value=""  >
                </div>
            </div>
            <br><br>
            
             <div class="form-group">
                <label for="passwordl" class="control-label col-xs-4 col-xs-offset-1">Password</label>
                <div class="col-xs-6">
                    <input autocomplete="new-password" name="pwd"  maxlength="10" type="password" class="form-control" id="pwd" value=""  >
                </div>
            </div>    
            
             <div class="form-group">
                <div class="col-xs-offset-5 col-xs-4">
                    <%
                    Time1 = TimeValue("9:59:00 PM")
                    Time2 = TimeValue("10:10:00 PM") 
                    '					FormatDateTime(Now)                = 2/29/2016 1:02:03 PM
                    '					FormatDateTime(Now, vbGeneralDate) = 2/29/2016 1:02:03 PM
                    '					FormatDateTime(Now, vbLongDate)    = Monday, February 29, 2016
                    '					FormatDateTime(Now, vbShortDate)   = 2/29/2016
                    '					FormatDateTime(Now, vbLongTime)    = 1:02:03 PM
                    TimeCheck = TimeValue(Now)
                    If (Time1 < TimeCheck AND TimeCheck < Time2) Then
					%>
                    
                    <label class="red">Data loading<br>Please wait...</label>
                    
					<% Else %>
                    
                    <input style="width:120px" name="Entrar" type="submit" class="btn btn-primary" id="Entrar" value="Login" >
                    
					<% End If %>
                
                </div>
            </div>
            
            <div class="form-group">
               <div class="col-xs-offset-5 col-xs-6">
                    
                    <% 
					
                    If Request("e") > 0 Then 
                        If Request("e") = 1 Then
                            Response.Write "<label class='red'>INVALID PASSWORD!</label>"
                        End if
                    End if
					
					 If Request("maint") > 0 Then 
                        If Request("maint") = 1 Then
                            Response.Write "<label class='red'>Periodo de Mantenimiento<br>Website on maintenance</label>"
                        End if
                    End if
                    %>
                    
                </div>
            </div>
                     
           
            
             
		</div>
                      
		
        <div class="col-lg-6 col-xs-12 hidden-xs ">
    
    		<img  src="logooic.png" width="568" height="322" alt="logo3">
           
    	</div>
    
	</div>
    
    <footer class="panel-footer col-xs-12">
        <p>Contact us:(T) 787-751-4036&nbsp;(F) 787-765-6735	&nbsp;Correo Electrónico/Email: oic@overseaspr.com<br>
        Physical Address: Urb. El Paraiso Calle Ganges #9 San Juan P.R. 00926</p>
    </footer>
                    
    <footer class="panel-footer col-xs-12">Overseas Import Corporation  All rights reserved TM 2016</footer>
    </form>    
         
</div>
</body>
</html>
