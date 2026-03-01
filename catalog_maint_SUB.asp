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
'if Session("MM_UserName") <> "Z099" then 
'	MM_authFailedURL="index.asp"
'	Response.Redirect(MM_authFailedURL)
'End if



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
	
	
	Recordset1_cmd.CommandText = "SELECT * FROM dbo.OIC_Subcategory WHERE categoryid = " + CStr(vcatid) + " ORDER BY category"
	Recordset1_cmd.Prepared = true
	Set Recordset1 = Recordset1_cmd.Execute
	Recordset1_numRows = 0
Else
	response.redirect("index.asp")
End if
%>
<%
Set Recordsetsub_cmd = Server.CreateObject ("ADODB.Command")
Recordsetsub_cmd.ActiveConnection = MM_overseaspr_STRING
Recordsetsub_cmd.CommandText = "SELECT * FROM dbo.OIC_Subcategory ORDER BY category"
Recordsetsub_cmd.Prepared = true
Set Recordsetsub = Recordsetsub_cmd.Execute
Recordsetsub_numRows = 0
%>
<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation - Order Detail</title>

<!-- Bootstrap -->
<link rel="stylesheet" type="text/css" href="bootstrap-3.3.6-dist/css/bootstrap.min.css">

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script type="text/javascript" charset="utf-8" src="bootstrap-3.3.6-dist/jquery.min.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->

<link href="overseas.css" rel="stylesheet" type="text/css" >
<script src="jquery-ui-1.12.0-rc.1/external/jquery/jquery.js"></script>
<link rel="stylesheet" type="text/css" href="jquery-ui-1.12.0-rc.1/jquery-ui.theme.min.css">
<link rel="stylesheet" type="text/css" href="jquery-ui-1.12.0-rc.1/jquery-ui.structure.min.css">
<script src="jquery-ui-1.12.0-rc.1/jquery-ui.js"></script>
<link rel="stylesheet" type="text/css" href="jquery-ui-1.12.0-rc.1/jquery-ui.min.css">
<link rel="stylesheet" type="text/css" href="jquery-ui-themes-1.12.0-rc.1/themes/redmond/jquery-ui.min.css">

<!--<link href="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/css/ui-darkness/jquery-ui-1.10.4.custom.css" rel="stylesheet">
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-1.10.2.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-ui-1.10.4.custom.js"></script>

<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/jquery-1.10.2.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.core.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.widget.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.button.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.menu.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.position.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.tooltip.min.js"></script>
-->

<style type="text/css" title="currentStyle">
	/*@import "DataTables-1.9.4/media/css/demo_page.css";
	@import "DataTables-1.9.4/media/css/demo_table_jui.css";*/

	/*body {
	margin:0;
	background-color: #FFF;
	background-repeat: repeat;
	}
	.input focus{
		border: #FC0;
	}
    
	body,td,th {
	font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
	font-size: 14px;
	color: #000;
	}*/
	
	/*#dt_example {
		color: #555555;
		font-size:14px;
		font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
	}
	.ui-widget-content a {
		color: #555555;
		
	}
	
	.ui-widget-header {
		color: #555555;
	}*/
	
</style>

<!--<style type="text/css" title="currentStyle">
	@import "DataTables-1.9.4/media/css/demo_page.css";
	@import "DataTables-1.9.4/media/css/demo_table_jui.css";

	
	body,td,th {
	font-size: 14px;
	color: #000;
	font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
	}
    .printbuttonfloatright {
		float: right;
		margin-top: 0px;
		margin-right: 50px;
		margin-bottom: 0px;
		margin-left: 0px;
	}
	
	body {
	background-image: url();
	margin-left: 10px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
	color: black;
	}
	a , h1,h2 {
	margin-left: 5px;
	margin-right: 5px;
	margin-top: 12px;
	margin-bottom: 8px;
	padding: 5px;
	text-decoration: none;
	}
</style>
-->




        
<script type="text/javascript" charset="utf-8">

function callfunctionDelCat(id) {
	  var isGood=confirm('Esta seguro que quiere eliminar esta subcategoría?');
		if (isGood) {
		  //alert('true');
		} else {
		  return;
		}

		var tagname = "nombresubcat" ;
		var newvalue = $('#'+tagname).val();
		
		window.parent.location.href = 'deletesub.asp?id=' + id.toString() 
	}
	
function callfunctionAddSubCat(id) {
		var tagname = "nombresubcat" ;
		var newvalue = $('#'+tagname).val();
		if (newvalue > '') {
			window.parent.location.href = 'AddCatalogSubCategory.asp?id=0' + '&newval=' + newvalue;
		} else { 
			$("#errormessage").val('Entre nombre de categoría nueva!')
			$("#errormessage").show();
		}
	}
	
function callfunction(id) {
		var tagname = "newcat" + id.toString();
		var newvalue = $('#'+tagname).val();
		if (newvalue > '') {
			window.parent.location.href = 'renamesub.asp?id=' + id.toString() + '&newval=' + newvalue;
		} else { 
			$("#errormessage").val('Entre nombre de subcategoría nueva!')
			$("#errormessage").show();
		}
	}
	
jq = jQuery.noConflict();
$(document).ready(function () {
   	
	
	
	$('#print').button();
	$('#back').button();
	$("#addcatb").button();
	$('.Rename').button();
	$('.Delete').button();
	$('#example a, button').button();
	
});
</script>
</head>

<body id="dt_example">
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
        
    
    <h4>Mantenimiento de Catálogo ::&nbsp;<%= GetCatalog(CInt(Session("MM_CatalogID")))%> </h4>
	
    
    
    <div class="breadcrumb"  >&nbsp;
	<a title="Regresar a lista de marcas" href="/catalog_maint_cat.asp" style="font-size:16px" ><%= GetNombreCategoria(Session("MM_CatID"))%></a>
    </div>
    
    
    
    <% if Session("MM_UserName") = "Z099" then %>   
    
    <h4> &nbsp;&nbsp;Añadir modelos de <%= GetCategoryName(CInt(vcatid))%></h4>	
    
    <div class="row">
     
      <div class="col-xs-4">
            <label>Nombre de marca seleccionada (Brand)
                &nbsp;&nbsp;<input title="Seleccione para ir a esta marca de auto (Click)" onClick="document.location.href='catalog_maint_CAT.asp'" name="nombrecat" type="text" class="btn btn-default" id="nombrecat" value="<%= GetCategoryName(CInt(vcatid))%>"  maxlength="50" >
            </label>
       </div>
       <div class="col-xs-4">
            <label>Nombre de modelo a crear en catálogo (Model)
                &nbsp;&nbsp;<input title="Para entrar un nuevo modelo entre el nombre y oprima Añadir modelo " class="form-control"  id="nombresubcat" name="nombresubcat" type="text"  maxlength="50" >
            </label>
      </div>
      <div class="col-xs-2">
      	<br>
      	<input id="addcatb" type="button" value="Añadir modelo" onClick="callfunctionAddSubCat(0)" class="btn btn-primary" >     
      </div>
      
      

    </div>
      
   <div class="row">
        
      		<input disabled class="col-xs-12 text-warning"  id="errormessage"  value="<%= Session("MM_Update") %>" >     
      	    
        </div>
       
	
    <% End If %>
    <div class="table-responsive">
      <table  class="table table-striped" id="example"   >
        <thead>
          <tr>
            <th>Modelos</th>
            
            <% if Session("MM_UserName") = "Z099" then %>    
            <th>Nombre</th>
            <% End IF %>
            
          </tr>
        </thead>
        <tbody>
          <% While  (NOT Recordset1.EOF) %>
          <tr>
            <td align="center" valign="middle">
              
              <% if Session("MM_UserName") = "Z099" then %>    
              <a class="btn btn-primary") href="routecat3.asp?id=<%=(Recordset1.Fields.Item("subcategoryid").Value)%>"  style="width:200px" >
              <%= UCase(Recordset1.Fields.Item("category").Value)%>
              </a>
              <% else %>
              <a class="btn btn-primary" href="routecat2.asp?id=<%=(Recordset1.Fields.Item("subcategoryid").Value)%>"  style="width:200px" >
              <%= UCase(Recordset1.Fields.Item("category").Value)%>
              </a>
              
              <% end if %>  
            </td>
            
             
            <td align="center" valign="middle">
              <input class="form-control" style="text-align:center; margin-left:5px" id="newcat<%=(Recordset1.Fields.Item("subcategoryid").Value)%>" name="newcat<%=(Recordset1.Fields.Item("subcategoryid").Value)%>"  type="text" value="<%= trim(Recordset1.Fields.Item("category").Value)%>" size="20">
          	</td>
            
             <% if Session("MM_UserName") = "Z099" then %>  
            <td>
            
              <input class="btn btn-primary" type="button"  id="Rename<%=(Recordset1.Fields.Item("subcategoryid").Value)%>" value="Cambiar nombre"
       onClick="callfunction(<%=(Recordset1.Fields.Item("subcategoryid").Value)%>)" >
              &nbsp;&nbsp;
              
              <input class="btn btn-danger" title="Eliminar Subcategoría"  type="image" src="/images/del.png"  alt="Submit" onClick="callfunctionDelCat(<%=(Recordset1.Fields.Item("subcategoryid").Value)%>)" >
            </td>
            <% End IF %>
            
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
RecordsetSub.Close()
Set RecordsetSub = Nothing
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
