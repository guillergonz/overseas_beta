<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>



<!doctype html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation</title>
<link rel="icon" href="images/favicon.ico" type="image/x-icon" /> 
<link rel="icon" href="images/favicon.ico" type="image/x-icon" />


<!--<link rel="stylesheet" type="text/css" href="http://fonts.googleapis.com/css?family=Ultra">-->

<link rel="stylesheet" type="text/css" href="js/vegas/jquery.vegas.css">
<link rel="stylesheet" type="text/css" href="js/jscrollpane/jquery.jscrollpane.css">
<link rel="stylesheet" type="text/css" href="css/styles.css">
<script src="http://code.jquery.com/jquery-1.6.2.min.js"></script>


<script src="js/jquery.easing.js"></script>
<script src="js/vegas/jquery.vegas.js"></script>
<script src="js/jscrollpane/jquery.jscrollpane.min.js"></script>
<script src="js/buzz/buzz.js"></script>
<script src="js/gallery.js"></script>

<%
' if Session("MM_UserName") <> "Z099" then 
' 	MM_authFailedURL="index.asp"
'    Response.Redirect(MM_authFailedURL)
' End if
%>



<%

Dim MM_overseaspr_STRING2
MM_overseaspr_STRING2 = "dsn=overseaspr;uid=overseaspr;pwd=2protectus;"

Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_overseaspr_STRING2
Recordset1_cmd.CommandText = "SELECT field_1,field_2,familia_descripcion, fam_make_item, parts_unique_id, field_3, image_exist FROM dbo.partmst1_distinct WHERE image_exist = '2' ORDER BY familia_descripcion,field_1 ASC" 
Recordset1_cmd.Prepared = true
Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>


<%
FUNCTION GetPartDesc(partno)

	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT field_2 FROM dbo.partmst1_distinct WHERE field_1 = '" + partno + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING2
		
	If Not oRS.EOF Then
		GetPartDesc = CStr(oRS.Fields.Item("field_2"))
	Else
		GetPartDesc = ""
		oRS.Close
		Set oRS = Nothing
	End if

END FUNCTION
%>
        
</head>

<body>

    <div id="flash"></div>
    <div id="title">
        <h1>Overseas Import Corporation</h1>
        <p>Ofertas y Especiales <a href="../../part_search.asp">Overseas Import Corporation Web Catalog - CATALOGO DE PIEZAS</a></p>
    </div>
    <div id="thumbnails">
        <ul>

	<% 
			DIM vpartno
			DIM vloop
			vloop = 0

			
            
            While Not Recordset1.Eof
            
  '          	if CInt(vloop) <= 9 THEN
'                	simage 	= "img/0" + CStr(vloop) + ".jpg"
'					simageb = "img/0" + CStr(vloop) + "b.jpg"
'                else
'                	simage = "img/" + CStr(vloop) + ".jpg"
'                	simageb = "img/" + CStr(vloop) + "b.jpg"
'                end if    

			vloop = vloop + 1
			vpartno = Recordset1("field_1")
			
			simage = "../../parts_images/" + vpartno + ".jpg"
			simageb = "../../parts_images/" + vpartno + "_renamed.jpg"
		'	response.write ("simage " & simage & "<br>")
		'	response.write ("simageB " & simageb & "<br>")
				
			%>
	
            
            
                            
           	<li><a href="<%= simage %>"><img src="<%= simageb %>" title="70 % de Descuento! <%= vpartno + " " + GetPartDesc(vpartno) %>" data-valign="top"  ></a></li>
            
           	<%
           		Recordset1.MoveNext
	        	wend
     		%>
                    	
<!--	
            	<li><a href="img/01.jpg"><img src="img/01b.jpg" title="<= vpartno + " " + GetPartDesc(vpartno) %>" data-valign="top"  ></a></li>
            <li><a href="img/02.jpg"><img src="img/02b.jpg" title="<= GetPartDesc(vpartno) %>" data-valign="bottom"></a></li>
            <li><a href="img/03.jpg"><img src="img/03b.jpg" title="<= GetPartDesc(vpartno)  %>"></a></li>
            <li><a href="img/04.jpg"><img src="img/04b.jpg" title="<= GetPartDesc(vpartno)  %>"></a></li>
            <li><a href="img/05.jpg"><img src="img/05b.jpg" title="Part Number 25721GA443T"></a></li>
            <li><a href="img/06.jpg"><img src="img/06b.jpg" title="Part Number 25721GA443T" data-valign="top"></a></li>
            <li><a href="img/07.jpg"><img src="img/07b.jpg" title="Part Number 25721GA443T"></a></li>
            <li><a href="img/08.jpg"><img src="img/08b.jpg" title="Part Number 25721GA443T" data-valign="top"></a></li>
            <li><a href="img/09.jpg"><img src="img/09b.jpg" title="Part Number 25721GA443T" data-valign="top"></a></li>
            <li><a href="img/10.jpg"><img src="img/10b.jpg" title="Part Number 25721GA443T"></a></li>
            <li><a href="img/11.jpg"><img src="img/11b.jpg" title="Part Number 25721GA443T" data-valign="top"></a></li>
            <li><a href="img/12.jpg"><img src="img/12b.jpg" title="Part Number 25721GA443T"></a></li>
-->     

	   </ul>
    	<div id="pointer"></div>
    </div>
    <div id="pause"><a href="#">Paused</a></div>
    <div id="volume" class="all"><a href="#">Sounds</a></div>

</body>
</html>
