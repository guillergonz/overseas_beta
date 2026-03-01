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
	Recordset1_cmd.CommandText = "SELECT * FROM dbo.OIC_Catalog ;"
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
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation - Order Detail</title>

<link href="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/css/ui-darkness/jquery-ui-1.10.4.custom.css" rel="stylesheet">
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-1.10.2.js"></script>
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

	
	body,td,th {
	font-size: 12px;
	color: #000;
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
</style>

<script type="text/javascript" language="javascript" src="DataTables-1.9.4/media/js/jquery.js"></script>
<script type="text/javascript" language="javascript" src="DataTables-1.9.4/media/js/jquery.dataTables.js"></script>

        
<script type="text/javascript" charset="utf-8">
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
		"bLengthChange":false,
		"bInfo":false,
		"bSort":false
	});	
	
	
	//oTable.fnSort( [ [1,'asc'], [4,'asc'], [3,'asc'] ] );

	
	$('#print').button();
	$('#back').button();

	$('a, button').button();
	
});
</script>
</head>

<body id="dt_example">
<div class="ui-widget" id="container" >
	
  
	 <div class="css_left"  >
	  
      <a id="back" name="back" href="part_search.asp">Catálogo de Internet OIC</a>       
        <!--<a id="catalog" name="catalog" href="catalog.asp">Catálogo de Imágenes</a>-->
    </div>
   
    
    <br clear="all">
   <h2 class="ui-widget-header" style="height:24px" >Seleccione Catálogo</h2>
   
  <hr>
  
    
    <% if Session("MM_UserName") = "Z099" then %>    
    <div class="ui-widget-content ui-corner-all" style="padding:20px;margin-left:10px;margin-right:10px"  ><br>
      
		<h2> Añadir catálogo ...&nbsp;&nbsp;</h2>	
		<div class="ui-widget" style="margin-left:20px; margin-right:20px; padding:10px">       
            <div style="height:40px; vertical-align:baseline">
            <legend style="width:250px; margin:0px">Nombre del catálogo</legend>
            &nbsp;&nbsp;<input class="ui-widget-header" name="nombrecat" type="text" size="30" maxlength="50">
            &nbsp;&nbsp;<a class="ui-state-disabled" id="addcatalog" name="addcatalog" >Añadir Catálogo</a>
            
        </div>
        <br><br>
	</div>
    <% End If %>
        
    <br clear="all">
    <table id="example" cellpadding="5" cellspacing="5" width="100%" >
    <thead>
    <tr>
    <th>Lista de Catálogos</th>
    <th>Orden</th>
    <th>Activo</th>
    </tr>
    </thead>
    <tbody>
	<% While  (NOT Recordset1.EOF) %>
    <tr>
	<td align="center" valign="middle">
    <a style="width:90%" href="routecat.asp?id=<%=(Recordset1.Fields.Item("idcatalog").Value)%>">
    <%= trim(Recordset1.Fields.Item("catalogname").Value)%>
    </a></td>
    <td align="center" valign="middle">         
    <label ><%=(Recordset1.Fields.Item("orden").Value)%></label>
    </td>
    <td align="center" valign="middle">
		<% if (Recordset1.Fields.Item("activo").Value = 1) then %>
        <h3>Sí</h3>
        <% else %>
		<h3>No</h3>)
        <% End If %>
        </td>
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
