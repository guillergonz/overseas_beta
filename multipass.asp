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
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="../index.asp"
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
<%
If GetSecureVal(Request("id") <> "") Then

	userid = GetSecureVal(Request.Querystring("id"))
	If LEN(userid) > 0 Then 
		
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open MM_overseaspr_STRING
		sql="UPDATE dbo.multiuser SET password = '11111' WHERE userid = '" + CStr(userid) + "'" 
		objConn.Execute sql
		sError = err.description
		If len(sError) > 0 Then
			Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
		End if
		objConn.Close
		Set objConn = Nothing		
		response.redirect("part_search.asp")
				
	End If
	
End if	
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_overseaspr_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.multiuser order by userid" 
Recordset1_cmd.Prepared = true
Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>

<!DOCTYPE HTML>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

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


<!--<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation - Login</title>
<link rel="icon" href="images/favicon.ico" type="image/x-icon" /> 
<link rel="icon" href="images/favicon.ico" type="image/x-icon" />
<link href="overseas.css" rel="stylesheet" type="text/css"><!--[if lt IE 9]>
<script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
<![endif]-->
<!--<style type="text/css">
body {
	background-image: url(AnimatedFrom/images/bg.gif);
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
</style>-->


<!--<link rel="stylesheet" href="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/themes/base/jquery.ui.all.css">
<script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/jquery-1.6.2.js"></script>
<script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/ui/jquery.ui.core.js"></script>
<script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/ui/jquery.ui.widget.js"></script>
<script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/ui/jquery.ui.button.js"></script>
<link href="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/demos/demos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<link href="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/css/ui-darkness/jquery-ui-1.10.4.custom.css" rel="stylesheet">
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-1.10.2.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-ui-1.10.4.custom.js"></script>

<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/jquery-1.10.2.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.core.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.widget.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.button.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.menu.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.position.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.tooltip.min.js"></script>-->

<!--<style type="text/css" title="currentStyle">
@import "DataTables-1.9.4/media/css/demo_page.css";
@import "DataTables-1.9.4/media/css/demo_table_jui.css";

body,td,th {
font-size: 14px;
color: #000;
font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
background-color: silver;
}
body {
background-image: url();
margin-left: 10px;
margin-top: 10px;
margin-right: 10px;
margin-bottom: 10px;
color: black;
background-color: #000;
}
a , h1,h2 {
margin-left: 5px;
margin-right: 5px;
margin-top: 12px;
margin-bottom: 8px;
padding: 5px;
text-decoration: none;
}

.button-link {
    padding: 5px 7px;
    background: #4479BA;
    color: #FFF;
    -webkit-border-radius: 4px;
    -moz-border-radius: 4px;
    border-radius: 4px;
    border: solid 1px #20538D;
    text-shadow: 0 -1px 0 rgba(0, 0, 0, 0.4);
    -webkit-box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.4), 0 1px 1px rgba(0, 0, 0, 0.2);
    -moz-box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.4), 0 1px 1px rgba(0, 0, 0, 0.2);
    box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.4), 0 1px 1px rgba(0, 0, 0, 0.2);
    -webkit-transition-duration: 0.2s;
    -moz-transition-duration: 0.2s;
    transition-duration: 0.2s;
    -webkit-user-select:none;
    -moz-user-select:none;
    -ms-user-select:none;
    user-select:none;
}
.button-link:hover {
    background: #356094;
    border: solid 1px #2A4E77;
    text-decoration: none;
}
.button-link:active {
    -webkit-box-shadow: inset 0 1px 4px rgba(0, 0, 0, 0.6);
    -moz-box-shadow: inset 0 1px 4px rgba(0, 0, 0, 0.6);
    box-shadow: inset 0 1px 4px rgba(0, 0, 0, 0.6);
    background: #2E5481;
    border: solid 1px #203E5F;
}
body,td,th {
	color: #000;
}
h2 {
	height:14px;
	padding:2px;
	-webkit-border-radius: 4px;
    -moz-border-radius: 4px;
    border-radius: 4px;
}
</style>-->
<!--<script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/ui/jquery.effects.core.js"></script>
<script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/ui/jquery.effects.fade.js"></script>
<script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/ui/jquery.effects.pulsate.js"></script><script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/ui/jquery.effects.slide.js"></script> -->   
    
<script type="application/javascript">

 function updatepassword(userid) {
	var opwd = "pwd" + (userid);
	var fieldx = encodeURIComponent(document.getElementById(opwd).value);
	//alert(fieldx);
	$.ajax({
		type:"POST",
		url: "ajaxupdate.asp",
		context: document.body,
		data: "userid=" + userid + "&pwd=" + fieldx,
		success: function(outputhtml){
			$("#ajaxDiv").html(outputhtml);
			
			//$("#ajaxmessage").html('');	
			//$("#ajaxDiv").fadeIn("3000");
			//$("#ajaxDiv").fadeOut("6000");
			
	   },
	   error: function(xhr, textStatus, errorThrown){
		$("#ajaxDiv").html(textStatus + errorThrown );
			alert(textStatus);
	   }  
	});
	
  };
  
$(document).ready(function() {
	
  $("#effect").css("display", "inline");

  $("#ajaxmessage").html('');	
  			
});
</script>
</head>

<body>


<div class="container"  >
	<br>

	
    
    <div class="row">
		<h4 class="col-xs-12" >Mantenimiento de usuarios de <em>B&V Auto Parts</em></h4>
	</div>
    
    <div class="row">
    <div class="col-xs-12">
    	<br>
        <div id="ajaxDiv" name="ajaxDiv" ></div>
        <br>
    </div>

	<div class="row">
        <div class="responsive">
            <table class="table-striped"   >
            <tr>
            <td  >ID de usuario</td>
            <td  >Contraseña</td>
            <td  >Nombre</td>
            <td  >&nbsp;</td>
            </tr>
            
            <% Do While not Recordset1.eof %>
            
            <tr>
            <td width="100px" height="50px"><input class="form-control"  type="text" id="<%= Recordset1("userid")%>" name="<%= Recordset1("userid")%>" value="<%= Recordset1("userid")%>"></td>
            <td height="50px"><input class="form-control"  type="password" id="pwd<%= Recordset1("userid")%>" name="pwd<%= Recordset1("userid")%>" value="<%= Recordset1("password")%>"></td>
            <td height="50px"><input class="form-control"  type="text" id="name<%= Recordset1("userid")%>" name="name<%= Recordset1("userid")%>" value="<%= Recordset1("username")%>"></td>
            <td height="50px"><input class="btn btn-primary small"  class="button-link" type="button" id="guardar<%= Recordset1("userid") %>" name="guardar<%= Recordset1("userid") %>" value="Guardar" onClick="javascript:updatepassword('<%= Recordset1("userid")%>');" ></td> 
            </tr>
            
            <% 
                Recordset1.MoveNext
            Loop
            %>
                
            </table>
        </div>
	</div>
    <br>
	<div class="row">
	<a class="btn btn-primary" href="part_search.asp" >Ir al catálogo</a>
	</div>
    
</div>
</body>
</html>


<%
Recordset1.close
Set Recordset1 = Nothing
%>
