<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
Response.Expires = -1

MM_Logout = GetSecureVal(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_UserID")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "index.asp"
	' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In (Request.QueryString)
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(GetSecureVal(Request.QueryString(Item)))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = UCASE(CStr(Request("usr2")))
MM_valUserpwd = UCASE(CStr(Request("pwd2")))

Session("MM_Username") = "B001"
Session("MM_UserId_Multi") = MM_valUsername

If MM_valUsername <> "" AND MM_valUserpwd <> "" Then

  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "security_proxy.asp"
  MM_redirectLoginFailed = "multilogin.asp?e=1"
  
  MM_loginSQL = "SELECT * FROM dbo.multiuser WHERE userid = ? AND password = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_overseaspr_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 20, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 10, MM_valUserpwd) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute
  
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Multi_Username") = trim(MM_rsUser("username"))
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
	'if Not MM_valUsername = "B0011" then
    	MM_rsUser.Close
    	Response.Redirect(MM_redirectLoginSuccess)
  	'end if
	
 
  End If

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

	
<script>
$(document).ready(function() {

	$("#submit").click(function() {
		$("#submit").hide();
	});
	
	$("#usr2").change(function() {
		$("#pwd2").val() == "";
	});
		
});
</script>

</head>
<body>

<div class="container" >
	
	<div class="jumbotron">
	    <img src="images/oiclogo2.gif" alt="overseas_logo" width="256" height="31" align="top">
        <h3>Welcome B & V Auto Parts</h3>
    </div>
    
    
     <form class="form-horizontal" action="<%=MM_LoginAction%>" method="POST" name="validacion" id="validacion" autocomplete="false" >
    
	<div class="row">
    
    	
		<div class="col-lg-6 col-xs-12">	
                         
            <div class="form-group">
                <label for="useridl" class="control-label col-xs-3 col-xs-offset-1">User ID</label>
                <div class="col-xs-6">
                   	<select name="usr2" size="5" class="form-control" id="usr2" autocomplete="false"  >
                    <option value="B0011">CARLE BETANCOURT</option>
                    <option value="B0012">OFICINA 2</option>
                    <option value="B0013">OFICINA 3</option>
                    <option value="B0014">COUNTER</option> 
					<option value="B0015">SAINT JUST STORE</option> 
                    </select>  
                </div>
            </div>
            
             <div class="form-group">
                <label for="passwordl" class="control-label col-xs-3 col-xs-offset-1">Password</label>
                <div class="col-xs-6">
                    <input autocomplete="new-password" name="pwd2"  maxlength="10" type="password" class="form-control" id="pwd2" value=""  >
                </div>
            </div>    
    
 			<div class="form-group">
                <div class="col-xs-offset-4 col-xs-4">
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
                    
                    <input name="Entrar" type="submit" class="btn btn-primary col-xs-12" id="Entrar" value="Login" >
                    
					<% End If %>
                
                </div>
            </div>
            
            <div class="form-group">
               <div class="col-xs-offset-4 col-xs-8">
                    
                    <% 
					
                    If Request("e") > 0 Then 
                        If Request("e") = 1 Then
                            Response.Write "<label class='red'>INVALID PASSWORD!</label>"
                        End if
                    End if
                    %>
                    
                </div>
            </div>
      
	</div>
          
    <div class="row">
    
     	<div class="col-xs-offset-1">
    		<a class="btn btn-danger" href="https://support.google.com/chrome/answer/95606?hl=en" >Como eliminar contraseñas guardadas en google chrome</a>
   	 	</div>
     
    </div>
    <br>
    
   <!-- <div class="row">        
           < if LEN(MM_valUserpwd) > 0 Then %>
           <div class="col-xs-offset-1">
               
               	<a class="btn btn-primary" href="multipass.asp">Cambiar contraseña</a>
               
               
               	<a class="btn btn-primary" href="security_proxy.asp">Continuar</a>
               
           </div> 
           < end if %>
    </div>
    <br>-->
    
   <!-- <div class="row">          
              
           <div class="form-group">
                <div class="col-xs-offset-3 col-xs-9">
                    <div class="badge">
                        <p>Contact us:(T) 787-751-4036&nbsp;(F) 787-765-6735<br>
                        Correo Electrónico/Email: oic@overseaspr.com<br>
                        Physical Address:
                        Urb. El Paraiso<br> Calle Ganges #9<br>
                        San Juan P.R. 00926</p>
                    </div>
                    
                </div>
           </div>     
           
             
		</div>   
    
	</div>-->
    
    <footer class="panel-footer col-xs-12">
        <p>Contact us:(T) 787-751-4036&nbsp;(F) 787-765-6735	&nbsp;Correo Electrónico/Email: oic@overseaspr.com<br>
        Physical Address: Urb. El Paraiso Calle Ganges #9 San Juan P.R. 00926</p>
    </footer>
    
    <footer class="panel-footer col-xs-12">Overseas Import Corporation  All rights reserved TM 2016</footer>
    
    
    </form>    
         
</div>    
</body>
</html>
