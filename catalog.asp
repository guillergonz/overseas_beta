<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
' *** Logout the current user.
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
Recordset1_cmd.CommandText = "SELECT * FROM dbo.partmst1_distinct WHERE image_exist = '1' ORDER BY fam_make_item,parts_unique_id DESC" 
Recordset1_cmd.Prepared = true
Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
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
	font-family: "Segoe UI", "Segoe UI Light", "Segoe UI Semibold", "Lucida Console";
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
		"bAutoWidth": false,
		"aLengthMenu":[[5,10,50,100,-1], [5,10,50,100, "All"]],
		"iDisplayLength":10,
		"sPaginationType": "full_numbers",
		"bLengthChange":true,
		"bInfo":true,
		"bSort":false
	});	
	
	
	
	$('#print').button();
	$('#back').button();

	$('a, button').button();
	
});
</script>
</head>

<body id="dt_example" >
<div class="ui-widget" id="container" >

	
  
 	<div class="css_left"  >
    	
          <a id="back" name="back" href="part_search.asp">Catálogo de Internet OIC</a>
   		<a id="catalog" name="catalog" href="catalog_maint.asp">Mantenimiento de Catálogos</a>
        
    </div>
   <br clear="all">
   <h2 class="ui-widget-header" style="height:24px" >Catálogo de Imágenes</h2><br>
   
 	
    
    <!--<div class="css_right" >
	<script type="text/javascript" charset="utf-8">
    if (window.print) {
        document.write('<form  style="margin:5px" ><input align="center" class="print_friendly" type=button name="print" id="print" value="   I M P R I M I R   " onClick="window.print()">')
	}
 	</script>
 	</div>-->
   
	<hr>    
  
	
 	<table id="example" cellpadding="5" cellspacing="5" width="100%" >
    <thead>
    <tr>
    <th>Category / <br>Categoría</th>
    <th>Description / <br>Descripción</th>
    <th>Part Number / <br>No. Pieza</th>
    <th>Image / <br>Imágen</th>
    </tr>
    </thead>
    <tbody>
        
   
     
    
      
      <% While  (NOT Recordset1.EOF) %>
      
        <tr>
          <td align="center" valign="middle">       
          <label style="color:blue"><%=(Recordset1.Fields.Item("family_description").Value)%></label>
          <br>
		  <%=(Recordset1.Fields.Item("familia_descripcion").Value)%>'
          </td>

          <td align="center" valign="middle">         
          <label style="color:blue"><%=(Recordset1.Fields.Item("english_version").Value)%></label>
          <br>
		  <%=(Recordset1.Fields.Item("field_2").Value)%>
          </td>

          <td align="center" valign="middle"><%=(Recordset1.Fields.Item("parts_unique_id").Value)%></td>
          
         
          <td align="center" valign="middle">
          
          <label><a href="parts_images/<%=(Recordset1.Fields.Item("field_1").Value)%>.jpg"  target="_parent"><img src="parts_images/<%=(Recordset1.Fields.Item("field_1").Value)%>.jpg" width="200" height="200"><br><%=(Recordset1.Fields.Item("field_1").Value)%></a>
          </label>
          
          </td>
          
        </tr>
        
        <% 
		 
		  Recordset1.MoveNext()
		Wend
		%>
        
    </tbody>
   
    </table>
	
  
</div>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
