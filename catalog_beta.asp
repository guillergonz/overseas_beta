<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->

<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="index.asp"
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
 if Session("MM_UserName") <> "Z099" then 
 	MM_authFailedURL="index.asp"
    Response.Redirect(MM_authFailedURL)
 End if
%>

<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
  Response.Write( MM_editAction )
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>

<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" or IsNull(condition) Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function

%>

<%
If (CStr(Request("MM_update")) = "catalog") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_overseaspr_STRING
    MM_editCmd.CommandText = "UPDATE dbo.especiales SET titulo1 = ?, mensaje1 = ?, titulo2 = ?, mensaje2 = ?, titulo3 = ?, mensaje3 = ? WHERE uniqueidcol = ? " 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 20 , MM_IIF(Request.Form("titulo1"), (Request.Form("titulo1")),"")) 
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 150 , MM_IIF(Request.Form("mensaje1"), (Request.Form("mensaje1")),""))

    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 20 , MM_IIF(Request.Form("titulo2"), (Request.Form("titulo2")),"")) 
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 150 , MM_IIF(Request.Form("mensaje2"), (Request.Form("mensaje2")),""))

    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 20 , MM_IIF(Request.Form("titulo3"), (Request.Form("titulo3")),"")) 
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 150 , MM_IIF(Request.Form("mensaje3"), (Request.Form("mensaje3")),"")) 	
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1 , 1 )
	
	MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

'     append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "catalog_beta.asp"
	'Response.Write( Request.QueryString )
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If

    Response.Redirect(MM_editRedirectUrl)

  End If
End If
%>


<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_overseaspr_STRING
Recordset2_cmd.CommandText = "SELECT titulo1,mensaje1,titulo2,mensaje2,titulo3,mensaje3 FROM dbo.especiales WHERE uniqueidcol = 1 " 
Recordset2_cmd.Prepared = true
Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>


<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_overseaspr_STRING
Recordset1_cmd.CommandText = "SELECT image_exist,field_1,field_2,familia_descripcion, fam_make_item, parts_unique_id, field_3 FROM dbo.partmst1_distinct WHERE image_exist in ('1','2') ORDER BY image_exist DESC" 
Recordset1_cmd.Prepared = true
Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>

<!doctype html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation</title>
<link rel="icon" href="images/favicon.ico" type="image/x-icon" /> 
<link rel="icon" href="images/favicon.ico" type="image/x-icon" />

		
<style type="text/css" title="currentStyle">
	@import "DataTables-1.9.0/media/css/demo_page.css";
	@import "DataTables-1.9.0/media/css/demo_table.css";
body,td,th {
font-family: "Segoe UI", "Segoe UI Light", "Segoe UI Semibold", "Lucida Console";
font-size: 12px;
}
.printbuttonfloatright {
	float: right;
	margin-top: 0px;
	margin-right: 50px;
	margin-bottom: 0px;
	margin-left: 0px;
}
.startsUgly { display: none; }
</style>

<link rel="stylesheet" href="nivozoom.pack1_.0/nivo-zoom.css" type="text/css" media="screen" />

<link rel="stylesheet" href="nivozoom.pack1_.0/custom-nivo-zoom.css" type="text/css" media="screen" />

<script type="text/javascript" src="jquery-ui/jquery-ui-1.8.16-sunny.custom/js/jquery-1.6.2.min.js"></script>
<script type="text/javascript" src="jquery-ui/jquery-ui-1.8.16-sunny.custom/js/jquery-ui-1.8.16.custom.min.js"></script>
<link type="text/css" href="jquery-ui/jquery-ui-1.8.16-sunny.custom/css/sunny/jquery-ui-1.8.16.custom.css" rel="stylesheet" />
<style type="text/css">
body {
    background-color: #FAFAFA;
    background-repeat: repeat;
    margin-left: 10px;
    margin-top: 10px;
    margin-right: 10px;
    margin-bottom: 10px;
}
</style>

<!--<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.4.2/jquery.min.js" type="text/javascript"></script>-->
<script src="nivozoom.pack1_.0/jquery.nivo.zoom.pack.js" type="text/javascript"></script>




<script type="text/javascript" language="javascript" src="DataTables-1.9.0/media/js/jquery.js"></script>
<script type="text/javascript" language="javascript" src="DataTables-1.9.0/media/js/jquery.dataTables.js"></script>

<script src="jquery-ui/jquery-ui-1.8.16-sunny.custom/development-bundle/ui/jquery.ui.core.js"></script>
<script src="jquery-ui/jquery-ui-1.8.16-sunny.custom/development-bundle/ui/jquery.ui.widget.js"></script>

<script type="text/javascript" charset="utf-8">
var $j = jQuery.noConflict();

	$j(document).ready(function() {
		
	   $("input:submit, input[type=button], a, .button" ).button();
	 
		
		$(".startsUgly").show();

		
	});
</script>
    
<script type="text/javascript" charset="utf-8">
$(window).load(function() {
	$('body').nivoZoom();
});
</script>

    <!--Then to handle those JavaScript disabled users, some <noscript> magic:-->
    <noscript>
      <style type="text/css">.startsUgly { display: block; }</style>
    </noscript>
 
	</head>
	<body id="dt_example">
	<div id="container" class="startsUgly">
      
      <img src="images/oiclogo2.gif" alt="Insert Logo Here" name="Insert_logo" width="256" height="31" id="Insert_logo" style="background: #C6D580; display:inline;" />
      
        <div>
          	
        <h3>&nbsp;Overseas Import Corporation - Catálogo de Internet</h3>
       
        
         <form action="catalog_beta.asp" target="_self" method="post" >
         
         
    <h1 style="font-weight:700;width:89%">Manejo de Especiales</h1>
         
        <a href="image_upload.asp" title="Upload Imagenes al catálogo web" target="_self" style="float:right;margin-left:10px;margin-right:150px;margin-top:20px;margin-bottom:20px" >Upload Imagenes al catálogo web</a>
         <a href="index_ggg.asp" title="Test page" target="_self">TEST PAGE </a><br clear="all">
        <br clear="all">
       
         
         
         
         <table width="972" border="1" cellpadding="2" cellspacing="2" class="ui-widget-content">
         <tr class="ui-state-default">
           <td height="29" align="center"><span class="ui-state-default">Imagen (Preview)</td>
           <td width="413" align="center"><span class="ui-state-default">Mensaje</td>
           </tr>
         
     
	
	
         <tr>
         <td width="590"><img height="267" width="400" src="parts_images/Especial1.jpg" alt="Especial1"></td>
         <td align="center" valign="middle">
           <input name="titulo1" type="text" class="ui-state-disabled" id="titulo1" value="Especial1.jpg" readonly>
           <label for="part"></label>
           <br>
           <a target="_self" title="Upload Imagen de Especial #1" onClick="window.location='urlroutetouploader.asp?part=especial1'">Upload Imagen #1</a>
           <br><br>
           <label>Mensaje del especial 1</label><br>
           <textarea name="Mensaje1" cols="50" rows="8" id="Mensaje1"><%= recordset2("mensaje1") %></textarea>
           <br>
           <input style="float:right;margin-left:10px;margin-right:50px;margin-top:20px;margin-bottom:20px;width:70%" name="submit<%= vloop2 %>" type="submit" value="Guardar Cambios" >
           
           
         </td>
         </tr>
         

 <tr>
         <td width="590"><img height="232" width="400" src="parts_images/Especial2.jpg" alt="Especial2"></td>
         <td align="center" valign="middle">
           <input name="titulo2" type="text" class="ui-state-disabled" id="titulo2" value="Especial2.jpg" readonly>
           <br>
           <a target="_self" title="Upload Imagen de Especial #2" onClick="window.location='urlroutetouploader.asp?part=especial2'">Upload Imagen #2</a>
           <br><br>
           <label>Mensaje del especial 2</label><br>
           <textarea name="Mensaje2" cols="50" rows="8" id="Mensaje2"><%= recordset2("mensaje2") %></textarea>
           
           <br>
           
           <input style="float:right;margin-left:10px;margin-right:50px;margin-top:20px;margin-bottom:20px;width:70%" name="submit<%= vloop2 %>" type="submit" value="Guardar Cambios" >
           
         </td>
         </tr>
         
 <tr>
         <td width="590"><img height="232" width="400" src="parts_images/Especial3.jpg" alt="Especial3"></td>
         <td align="center" valign="middle">
           <input name="titulo3" type="text" class="ui-state-disabled" id="titulo3" value="Especial3.jpg" readonly>
           <label for="part"></label>
           <br>
           <a target="_self" title="Upload Imagen de Especial #3" onClick="window.location='urlroutetouploader.asp?part=especial3'">Upload Imagen #3</a>
           <br><br>
           <label>Mensaje del especial 3</label><br>
          <textarea name="Mensaje3" cols="50" rows="8" id="Mensaje3"><%= recordset2("mensaje3") %></textarea>
          <br>
          <input style="float:right;margin-left:10px;margin-right:50px;margin-top:20px;margin-bottom:20px;width:70%" name="submit<%= vloop2 %>" type="submit" value="Guardar Cambios" >
          </td>
         </tr>
         
                  		
         
        
         
         </table>         
         </p>
         
	    <input type="hidden" id="MM_update" name="MM_update" value="catalog" />
		</form>
        	
         <br>           

        <hr>
        <br>
     	<br>    
    
    <br>
        	
</body>
</html>

<%
Recordset2.Close()
Set Recordset2 = Nothing
%>

<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
