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
<script src="jquery-ui/jquery-ui-1.8.16-sunny.custom/development-bundle/ui/jquery.ui.datepicker.js"></script>

	<script type="text/javascript" charset="utf-8">
    var $j = jQuery.noConflict();
    
        $j(document).ready(function() {
            
           $("input:submit, input[type=button], a, .button" ).button();
            
           // $('#bannerintro').button();
            //$('#back').button();
            
            var oTable = $j('#example').dataTable({
		        "aLengthMenu": [
					[10, 20, -1],
					[10, 20, "All"]
				],
				"iDisplayLength": 10
		    });
	
			
             oTable.fnSort( [ [5,'desc'] ] );
             
            //$j('#fecha').datepicker({inline: true});
            
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
         
         
         <h1 style="font-weight:700;width:89%">Subir imágenes al catálogo de internet</h1>
         
         <label>Enter part number:</label><br>
         <label for="part"></label>
         
         &nbsp;&nbsp;
         <label>Upload Image - Utilizar nombre de la pieza para crear una imágen asociada a la pieza</label>&nbsp;&nbsp;
         <input name="part" id="part" type="text" size="10" maxlength="14">
         <input class="button" id="uploadproc" name="uploadproc" type="submit" value="Subir Imágen para que se vea en el web site">
		 
         </form>
         
		<br>
        	                        
       

            
        
            
			<h4>&nbsp;&nbsp;&nbsp;Lista de imágenes disponibles</h4>

        
          
        
        </div>
                
        <div id="demo">
          <table border="0" cellspacing="2" cellpadding="3" id="example" width="100%" >
            <thead>
              <tr>
               	<th width="10%">Categoria</th>
        		<th width="20%">Descripción</th>
		        <th width="10%">No.Pieza</th>
                <th width="10%">Oferta</th>
                <th width="40%">Imagen</th>
                <th width="10%">&nbsp;</th>
              </tr>
            </thead>
            <tbody>
            
  <% 
Dim vloop
Dim vfam
Dim vclass


vloop = 0
if Not Recordset1.EOF then
 vfam = Recordset1("familia_descripcion")
else
 vfam = 0	
end if
vclass = "gradeA"
	
While Not Recordset1.Eof
	vloop = vloop + 1
	if (vfam <> Recordset1("familia_descripcion")) then 
		  vfam = Recordset1("familia_descripcion") 	
		  if (vclass = "gradeA") then
		  	vclass = "gradeC"
            Response.Write( "<tr class='gradeC'>") 
		  elseif (vclass = "gradeC") then 	
		  	vclass = "gradeU"
            Response.Write( "<tr class='gradeU'>") 
		  elseif (vclass = "gradeU") then 	
		  	vclass = "gradeX"
            Response.Write( "<tr class='gradeX'>") 
		  else
		  	vclass = "gradeA"
            Response.Write( "<tr class='gradeA'>") 
		  end if
	else
		  if (vclass = "gradeA") then
            Response.Write( "<tr class='gradeA'>") 
		  elseif (vclass = "gradeC") then 	
            Response.Write( "<tr class='gradeC'>") 
		  elseif (vclass = "gradeU") then 	
            Response.Write( "<tr class='gradeU'>") 
		  elseif (vclass = "gradeX") then 	
            Response.Write( "<tr class='gradeX'>") 
		  end if
	end if
%>
            
            <td><%= Recordset1("familia_descripcion") %></td>
            <td><%= Recordset1("field_2") %></td>
            <td><%= Recordset1("field_1") %></td>

			<td valign="middle" align="right"><%=(Recordset1.Fields.Item("field_3").Value)%></td>
            
            <td align="center" valign="middle">
                                      
           <a href="/parts_images/<%=(Recordset1.Fields.Item("field_1").Value)%>.jpg" target="_self" class="nivoZoom center">
	<img src="/parts_images/<%=(Recordset1.Fields.Item("field_1").Value)%>.jpg" alt="" width="32" height="32" /></a>
	<div class="nivoCaption"><a href="/overseaspr/part_search.asp" title="prueba" target="_self">prueba</a></div>

	</td>
    <td align="center" valign="middle">	
   
   		
            </td>
            
            </tr>
            
			<%
            Recordset1.MoveNext
            wend
            %>
            

            </tbody>
            
            <tfoot>
           		<td width="10%">Categoria</td>
        		<td width="40%">Descripción</td>
		        <td width="10%">No.Pieza</td>
                <td width="10%">Oferta</td>
                <td>Imagen</td>
            </tfoot>
  </table>
  
          
        </div>
        <br>
      </div>
		<br>
        <br>
        <hr>
        <br>
        <br>
        	
</body>
</html>



<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
