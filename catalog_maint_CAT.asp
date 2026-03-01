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
  If InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Not(Request.QueryString = "") Then
    MM_newQS = "?"
    For Each Item In (Request.QueryString)
      If Not (Item = "MM_Logoutnow") Then
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
If Not (Session("MM_Username") = "") Then
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
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Dim vcatid

If IsNull(Session("MM_CatID")) or IsEmpty(Session("MM_CatID")) then
	vcatid = 1
	Session("MM_CatID") = 1
Else	
	vcatid = CInt(Session("MM_CatID"))
End If
if vcatid > 0 Then
	Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
	Recordset1_cmd.ActiveConnection = MM_overseaspr_STRING
	Recordset1_cmd.CommandText = "SELECT * FROM dbo.OIC_Category ORDER BY category"
	Recordset1_cmd.Prepared = true
	Set Recordset1 = Recordset1_cmd.Execute
	Recordset1_numRows = 0
Else
	response.redirect("index.asp")
End if
%>
<%
Dim RecordsetCatalogos
Dim RecordsetCatalogos_cmd
Dim RecordsetCatalogos_numRows

Set RecordsetCatalogos_cmd = Server.CreateObject ("ADODB.Command")
RecordsetCatalogos_cmd.ActiveConnection = MM_overseaspr_STRING
RecordsetCatalogos_cmd.CommandText = "SELECT * FROM dbo.OIC_Catalog" 
RecordsetCatalogos_cmd.Prepared = true

Set RecordsetCatalogos = RecordsetCatalogos_cmd.Execute
RecordsetCatalogos_numRows = 0
%>
<!DOCTYPE HTML>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation</title>

<!-- Bootstrap -->
<link rel="stylesheet" type="text/css" href="bootstrap-3.3.6-dist/css/bootstrap.min.css">

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script type="text/javascript" charset="utf-8" src="bootstrap-3.3.6-dist/jquery.min.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->



<link href="overseas.css" rel="stylesheet" type="text/css" >





<!--<link rel="stylesheet" href="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/themes/base/jquery.ui.all.css">
<link href="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/css/ui-darkness/jquery-ui-1.10.4.custom.css" rel="stylesheet">-->

<!--<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-ui-1.10.4.custom.js"></script>-->
<script src="jquery-ui-1.12.0-rc.1/external/jquery/jquery.js"></script>

<link rel="stylesheet" type="text/css" href="jquery-ui-1.12.0-rc.1/jquery-ui.theme.min.css">
<link rel="stylesheet" type="text/css" href="jquery-ui-1.12.0-rc.1/jquery-ui.structure.min.css">

<script src="jquery-ui-1.12.0-rc.1/jquery-ui.js"></script>
<link rel="stylesheet" type="text/css" href="jquery-ui-1.12.0-rc.1/jquery-ui.min.css">


<link rel="stylesheet" type="text/css" href="jquery-ui-themes-1.12.0-rc.1/themes/redmond/jquery-ui.min.css">

<!--<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/jquery-1.10.2.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.core.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.widget.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.button.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.menu.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.position.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.tooltip.min.js"></script>-->

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
	}
	/*.ui-widget-content a {
		color: #555555;
	}
	
	.ui-widget-header {
		color: #555555;
	}*/
	
</style>

<script type="text/javascript" language="javascript" src="DataTables-1.9.4/media/js/jquery.js"></script>
<script type="text/javascript" language="javascript" src="DataTables-1.9.4/media/js/jquery.dataTables.js"></script>

        
<script type="text/javascript" charset="utf-8">

function callfunctionDelCat(id) {
	  var isGood=confirm('Esta seguro que quiere eliminar esta categoría?');
		if (isGood) {
		  //alert('true');
		} else {
		  return;
		}

		var tagname = "nombrecat" ;
		var newvalue = $('#'+tagname).val();
		
		window.parent.location.href = 'deletecat.asp?id=' + id.toString() 
	}
	
function callfunctionAddCat(id) {
		var tagname = "nombrecat" ;
		var newvalue = $('#'+tagname).val();
		if (newvalue > '') {
			window.parent.location.href = 'AddCatalogCategory.asp?id=0' + '&newval=' + newvalue;
		} else { 
			$("#errormessage").val('Entre nombre de categoría nueva!')
			$("#errormessage").show();
		}
	}
	
function callfunction(id) {
		var tagname = "newcat" + id.toString();
		var newvalue = $('#'+tagname).val();
		if (newvalue > '') {
			window.parent.location.href = 'renamecat.asp?id=' + id.toString() + '&newval=' + newvalue;
		} else { 
			$("#errormessage").val('Entre nombre de categoría nueva!')
			$("#errormessage").show();
		}
	}
	
jq = jQuery.noConflict();
$(document).ready(function () {
   	
	jq('#example').dataTable({
		"bDestroy": true,
		"bJQueryUI": true,
		"bDeferRender": false,
		"bAutoWidth": true,
		"aLengthMenu":[[5,10,50,100,-1], [5,10,50,100, "All"]],
		"iDisplayLength":50,
		"sPaginationType": "full_numbers",
		"bLengthChange":true,
		"bInfo":true,
		"bSort":true
	});	
	//oTable.fnSort( [ [1,'asc'], [4,'asc'], [3,'asc'] ] );	
	$('#print').button();
	$('#back').button();
	$('#oic').button();
	$('#catalog').button();
	$("#addcatb").button();
	$('.Rename').button();
	$('.Delete').button();
	//$('a, button').button();
	
});
</script>
</head>

<body id="dt_example">
<div class="container" >
	
   
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
                     <li class="active"><a href="catalog_maint_CAT.asp" title="<%= Lang("catalogo") %> Catalog" target="new"><%= Lang("catalogo") %></a></li>
                    
                    
                    <% If LEN(Session("MM_Multi_Username")) = 0 or Session("MM_Multi_Username") = "CARLE BETANCOURT" then %>
                    
                    <li><a href="account_statement_iframe.asp" title="<%= Lang("estado_de_cuenta_actual") %>" target="_self"><%= Lang("estado_de_cuenta_actual") %></a></li>
                        
                    <% End IF %>
            
            
                    
                    <% if ucase(Session("MM_UserName")) = "Z099" then %>
                    <li><a href="uploadDataDaily.asp" title="Import parts" target="_self">Import parts</a></li>
                    <li><a href="monitor_beta.asp" title="Monitor" target="_self">Monitor  <span class="badge"><%= GetCurrentSales(Session("MM_UserName")) %> </span></a></li>
                    <li><a href="usuarios.asp" title="Usuarios" target="_self">Usuarios</a></li>
                    <% End If %>
                </ul>
            </div>
        </div>
    </nav>
    

    <a class="navbar-brand" rel="home" href="#" title="Overseas Import Corporation">
        			<img style="max-width:256px; margin-top: -7px;" src="/images/oiclogo2.gif"></a>
                    
                    
    
    <br clear="all">
    <h4>Mantenimiento de Catálogo de <%= GetCatalog(CInt(Session("MM_CatalogID")))%> </h4>
	
    
    <% if Session("MM_UserName") = "Z099" then %>    
    <div class="ui-widget-content ui-corner-all"   >
    	<br>
		<h4> &nbsp;&nbsp;Añadir marcas de autos ...&nbsp;&nbsp;</h4>	
		
		<div style="margin:10px; padding:10px" >
		    <div style="height:40px; vertical-align:baseline">
            <label  style="margin:0px">Nombre de marca nueva (Brand)</label>
            &nbsp;&nbsp;<input title="Para entrar una nueva marca entre el nombre y oprima Añadir marca " class="ui-widget-content" id="nombresubcat" name="nombresubcat" type="text" size="30" maxlength="50" >
            &nbsp;&nbsp;
            <input id="addcatb" type="button" value="Añadir marca" onClick="callfunctionAddCat(0)" style="margin-left:5px" >        	
            &nbsp;&nbsp;<input disabled class="ui-state-error" id="errormessage" style="display:none;width:200px;background-color:transparent;margin:5px" value="<%= Session("MM_Update") %>" >
        </div>
       
	</div>
    <% End If %>
   
   	
    
    <table id="example" class="table table-striped"  >
    <thead>
    <tr>
    <th >Marcas</th>
   	
    <% if Session("MM_UserName") = "Z099" then %> 
    <th>Nombre</th>
    
    <th></th>
    <% End If %>
    
    </tr>
    </thead>
    <tbody>
	
	<% While  (NOT Recordset1.EOF) %>
    
    <tr>
	<td align="center" valign="middle" >
    <a style="width:90%" href="routecat4.asp?id=<%=(Recordset1.Fields.Item("categoryid").Value)%>">
    <%= UCase(Recordset1.Fields.Item("category").Value)%>
    <br>
    <img src="images/<%= UCase(Recordset1.Fields.Item("category").Value)%>.png" alt="Logo" height="140" width="143" onClick="href='routecat4.asp?id=<%=(Recordset1.Fields.Item("categoryid").Value)%>'"> 
    </a>
    </td>
  
    
    
  
   
    
    <% if Session("MM_UserName") = "Z099" then %> 
    
        <td align="center" valign="middle">
        	<input class="form-control"  id="newcat<%=(Recordset1.Fields.Item("categoryid").Value)%>" name="newcat<%=(Recordset1.Fields.Item("categoryid").Value)%>"  type="text" value="<%= trim(Recordset1.Fields.Item("category").Value)%>" size="20">
    	</td>
    	<td>
      		<input  type="button" class="btn btn-primary" id="Rename<%=(Recordset1.Fields.Item("categoryid").Value)%>" value="Cambiar nombre" onClick="callfunction(<%=(Recordset1.Fields.Item("categoryid").Value)%>)" >
   &nbsp;&nbsp;
   
     <input title="Eliminar categoría"  type="image" src="/images/del.png"  alt="Submit" onClick="callfunctionDelCat(<%=(Recordset1.Fields.Item("categoryid").Value)%>)" >
	
    </td>
    
    
    <% End If %>
    
    </tr>
	<% 
		  Recordset1.MoveNext()
		Wend
	%>
    </tbody>
    </table>


</div>
</div>

</body>
</html>
<%
RecordsetCatalogos.Close()
Set RecordsetCatalogos = Nothing
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
