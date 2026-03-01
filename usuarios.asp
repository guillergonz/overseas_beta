<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
' *** Logout the current user.
Response.Expires = -1

MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "index.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
if Session("MM_UserName") <> "Z099" then 
	MM_authFailedURL="index.asp"
	Response.Redirect(MM_authFailedURL)
End if

Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_overseaspr_STRING
	Recordset1_cmd.CommandText = "SELECT user_level,user_auto_id,user_name,user_pwd,lang,activestatus FROM dbo.users WHERE activestatus in ('A','I') ORDER BY activestatus,user_name; "

Recordset1_cmd.Prepared = true
Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0

%>
<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation</title>

<!-- Bootstrap -->
<link rel="stylesheet" type="text/css" href="bootstrap-3.3.6-dist/css/bootstrap.min.css">

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script type="text/javascript" charset="utf-8" src="bootstrap-3.3.6-dist/jquery.min.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->



<link href="overseas.css" rel="stylesheet" type="text/css" >

<link rel="stylesheet" href="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/themes/base/jquery.ui.all.css">
<link href="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/css/ui-darkness/jquery-ui-1.10.4.custom.css" rel="stylesheet">
<!--<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-1.10.2.js"></script>-->
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-ui-1.10.4.custom.js"></script>

<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/jquery-1.10.2.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.core.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.widget.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.button.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.menu.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.position.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.tooltip.min.js"></script>

<style type="text/css" title="currentStyle">
	@import "DataTables-1.9.4/media/css/demo_page.css";
	@import "DataTables-1.9.4/media/css/demo_table_jui.css";
	
body {
margin:0;
background-color: #FFF;
background-repeat: repeat;
}
.navbar {
    margin-bottom: 0px;
}
.input focus{
	border: #FC0;
}
body,td,th {
font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
font-size: 14px;
color: #000;
}

#dt_example {
	color: #555555;
	font-size:14px;
	font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
}
.ui-widget-content {
	background-color: #555555;
}

.ui-widget-content  input{
	color: #555555;
}

/*.ui-widget-content a {
	color: white;
}*/

button, input, textarea, select {
	color:black;
}

#example_info {
	color:white;
}

.dataTables, .example_paginate, .example_first, .example_previous  {
	color:white;
}
.printbuttonfloatright {
	float: right;
	margin-top: 0px;
	margin-right: 50px;
	margin-bottom: 0px;
	margin-left: 0px;
}
	
	
</style>

<script type="text/javascript" language="javascript" src="DataTables-1.9.4/media/js/jquery.js"></script>
<script type="text/javascript" language="javascript" src="DataTables-1.9.4/media/js/jquery.dataTables.js"></script>

        
<script type="text/javascript" charset="utf-8">

function callfunction(id) {
	
	var isGood=confirm('Esta seguro que quiere cambiar la contraseña a un status que obliga al usuario a cambiar el password. El usuario podra entrar con la contraseña "11111" ');
	if (isGood) {
	  callfunction3(id);	
	  //alert('true');
	} else {
	  return;
	}
	window.parent.location.href = 'resetpwd.asp?id=' + id.toString();
	$("#demo").hide(); 
}

function callfunction2(id) {
	window.parent.location.href = 'lockoutuser.asp?id=' + id.toString() + '&action=I';
	$("#demo").hide();
}

function callfunction3(id) {
	window.parent.location.href = 'lockoutuser.asp?id=' + id.toString() + '&action=A';
	$("#demo").hide();
}

//Spanish
function callfunction4(id) {
	var lang = $("#group1").val();
	var checkbox = "S";
	var url = 'lockoutuser.asp?id=' + id.toString() + '&action=' + lang + '&check=' + checkbox;
	window.parent.location.href = url;
	$("#demo").hide();
}
// English
function callfunction5(id) {
	var lang = $("#group2").val();
	var checkbox = "E";
	var url = 'lockoutuser.asp?id=' + id.toString() + '&action=' + lang + '&check=' + checkbox;
	window.parent.location.href = url;
	$("#demo").hide();
}


	
jq = jQuery.noConflict();
$(document).ready(function () {
	
	$("#demo").hide();
	
   	jq('#example').dataTable({
		//"bDestroy": true,
		"bJQueryUI": true,
		//"bDeferRender": false,
		"bAutoWidth": true,
		"aLengthMenu":[[5,10,50,100,-1], [5,10,50,100, "All"]],
		"iDisplayLength":-1,
		"sPaginationType": "full_numbers",
		"bLengthChange":true,
		"bInfo":true,
		"bSort":true
	});	
	
	$("#demo").show();

});
</script>
</head>

<body id="dt_example">
<form id="commentForm" name="commentForm" method="post" action="usuarios.asp" > 
    
<div class="container">	
  
	<nav class="navbar navbar-inverse navbar-static-top" role="navigation">
        <div class="container">
            <div class="navbar-header">
                <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">
                    <span class="sr-only">Toggle navigation</span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                </button>
                
            </div>
    
            <!-- Collect the nav links, forms, and other content for toggling -->
            <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
                <ul class="nav navbar-nav">
                     <li ><a href="part_search.asp" title="<%= Lang("buscar") %>" target="_self"><%= Lang("buscar") %></a></li>
                
                     <li><a id="completar" href="cart.asp"  title="<%= Lang("completar_orden")%>" target="_self" ><%= Lang("completar_orden") %></a></li>
                     <li><a href="catalog_maint_CAT.asp" title="<%= Lang("catalogo") %> Catalog" target="new"><%= Lang("catalogo") %></a></li>
                    
                    
                    <% If LEN(Session("MM_Multi_Username")) = 0 or Session("MM_Multi_Username") = "CARLE BETANCOURT" then %>
                    
                    	<li><a href="account_statement_iframe.asp" title="<%= Lang("estado_de_cuenta_actual") %>" target="_self"><%= Lang("estado_de_cuenta_actual") %></a></li>
                        
                    <% End IF %>
                    
                    <% if ucase(Session("MM_UserName")) = "Z099" then %>
                        <li><a href="uploadDataDaily.asp" title="Import parts" target="_self">Import parts</a></li>
                        <li><a href="monitor_beta.asp" title="Monitor" target="_self">Monitor  <span class="badge"><%= GetCurrentSales(Session("MM_UserName")) %> </span></a></li>
                        <li class="active"><a href="usuarios.asp" title="Usuarios" target="_self">Usuarios</a></li>
                    <% End If %>
                </ul>
            </div>
        </div>
    </nav>
    

    <a class="navbar-brand" rel="home" href="#" title="Overseas Import Corporation">
    	<img style="max-width:256px; margin-top: -7px;" src="/images/oiclogo2.gif">
    </a>
    
  
    <br clear="all">
   
    <div class="row" >
	
     <br>  
     <div class="table-responsive" id="demo"  style="display:none">
        <table class="table table-striped" id="example" >
        <thead>
        <tr>
        <th>User ID</th>
        <th>Level</th>
        <th>Last Purchase</th>
        <th>Name</th>
        <th>Password</th>
        <th>Language</th>
        <th>Active </th>
        <th>Action</th>
        </tr>
        </thead>
        
        <% While  (NOT Recordset1.EOF) %>
        <tr>
        <td align="center">
        <%= trim(Recordset1.Fields.Item("user_auto_id").Value)%>
        </td>
        <td align="center">
        <%= trim(Recordset1.Fields.Item("user_level").Value)%>
        </td>
        
        <td>
        <%= GetLastPurchase(Recordset1.Fields.Item("user_auto_id").Value)%>
        </td>
        
        <td >         
        <%= trim(Recordset1.Fields.Item("user_name").Value)%>
        </td>
       <td align="right">         
        <% 'if trim(Recordset1.Fields.Item("user_pwd").Value) <> "11111" then %>
        <%= trim(Recordset1.Fields.Item("user_pwd").Value) %>
        &nbsp;&nbsp;
        <input class="ui-state-active ui-corner-all" id="reset" onClick="callfunction('<%=(Recordset1.Fields.Item("user_auto_id").Value)%>')"  type="button" value="reset" >
        <% 'End IF %>
        
        </td>
        
        <td align="center">
      
         
        <div class="btn-group" data-toggle="buttons-checkbox">
        
          <input <% if (trim(Recordset1.Fields.Item("lang").Value) = "S") then response.write("checked"):response.write("")%> type="checkbox" id="group1"  onClick="callfunction4('<%=trim(Recordset1("user_auto_id"))%>')" name="language" value="1"	 >&nbsp;<%= ucase("Español")%>&nbsp;  
          
          <input <% if (trim(Recordset1.Fields.Item("lang").Value) = "E") then response.write("checked"):response.write("")%>  type="checkbox" id="group2"  onClick="callfunction5('<%=trim(Recordset1("user_auto_id"))%>')" name="language" value="1" >&nbsp;<%= ucase("Inglés")%>&nbsp;
          
  		</div>
                                      
        
        </td>
        
        <td align="center">
        [<%= trim(Recordset1.Fields.Item("activestatus").Value) %>]
        <% if trim(Recordset1.Fields.Item("activestatus").Value) = "A" then %>
         <label>Activo</label>
        <% ElseIf Not IsNull(Recordset1.Fields.Item("activestatus").Value) and trim(Recordset1.Fields.Item("activestatus").Value) <> "" then %>
         <label>Inactivo</label>
        <% End If %> 
         
        </td>
        
        <td align="center">
        <% 
        
        
        
        if trim(Recordset1.Fields.Item("activestatus").Value) = "A" then %>
         
            <input class="ui-state-default ui-corner-all" id="reset" onClick="callfunction2('<%=(Recordset1.Fields.Item("user_auto_id").Value)%>')"  type="button" value=" Inactivar Acceso " >
      
        <% Else %>
       
            <input class="ui-state-active ui-corner-all" id="reset" onClick="callfunction3('<%=(Recordset1.Fields.Item("user_auto_id").Value)%>')"  type="button" value=" Activar Acceso " >
      
        <% End IF %>
        </td>
        </tr>
        
        <% 
              Recordset1.MoveNext()
            Wend
        %>
       
        </table>
		
        </div>        
	</div>        
      
</div>

	<!--<input id="stodos" type="hidden" value="<= stodos %>" >
    <input id="sinac" type="hidden" value="<= sinac %>" >
    <input id="sact" type="hidden" value="<= sact %>" >
    <input id="aspa" type="hidden" value="<= aspa %>" >
    <input id="seng" type="hidden" value="<= seng %>" >-->
    
</form>
</body>
</html>


<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
