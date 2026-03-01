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
Dim vcatid,vcatalogid




' select_modelo will reload form
if Not IsNull(request("select_modelo")) AND Not IsEmpty(request("select_modelo")) AND LEN(request("select_modelo")) >= 0 Then
	vsubcatid = Clng(request("select_modelo"))
	Session("MM_SubCatID") = vsubcatid
End If

if not isnull(request("select_marca")) and LEN(request("select_marca")) > 0  Then
	vcatid = request("select_marca")
	Session("MM_CatID") = vcatid	
end if

		
If IsNull(Session("MM_CatID")) or IsEmpty(Session("MM_CatID")) or trim(Session("MM_CatID"))="" or trim(Session("MM_CatalogID"))="" or IsNull(Session("MM_CatalogID")) or IsEmpty(Session("MM_CatalogID")) Then

	response.redirect("index.asp")

else

	vsubcatid     = Session("MM_SubCatID")
	vcatid 		= Session("MM_CatID")
	vcatalogid 	= Session("MM_CatalogID")
	
	'response.write("<div style='background-color:white'>C:" + CStr(vcatid) + "<br>")
	'response.write("S:" + CStr(vsubcatid) + "</div>")
			
End If	

if vcatalogid > 0 AND vcatid > 0 Then

	 
	Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
	Recordset1_cmd.ActiveConnection = MM_overseaspr_STRING
	if CInt(vsubcatid) = 0 then	
		Recordset1_cmd.CommandText = "SELECT distinct titulo FROM dbo.OIC_PartsPerCategory WHERE categoryid = " + CStr(vcatid) + " AND idcatalog = " + CStr(vcatalogid) + " ORDER BY titulo"  
	else
		Recordset1_cmd.CommandText = "SELECT distinct titulo FROM dbo.OIC_PartsPerCategory WHERE categoryid = " + CStr(vcatid) + " AND idcatalog = " + CStr(vcatalogid) + " AND (subcategoryid = " + CStr(vsubcatid) + " or subcategoryid is null) ORDER BY titulo"  
	end if
	Recordset1_cmd.Prepared = true
	Set Recordset1 = Recordset1_cmd.Execute
	Recordset1_numRows = 0

Else
	response.redirect("index.asp")
End if
%>
<%
Dim RecordsetCategorias
Dim RecordsetCategorias_cmd
Dim RecordsetCategorias_numRows

Set RecordsetCategorias_cmd = Server.CreateObject ("ADODB.Command")
RecordsetCategorias_cmd.ActiveConnection = MM_overseaspr_STRING
RecordsetCategorias_cmd.CommandText = "SELECT CategoryId, Category FROM dbo.OIC_Category ORDER BY Category" 
RecordsetCategorias_cmd.Prepared = true

Set RecordsetCategorias = RecordsetCategorias_cmd.Execute
RecordsetCategorias_numRows = 0
If (RecordsetCategorias.CursorType > 0) Then
  RecordsetCategorias.MoveFirst
Else
  RecordsetCategorias.Requery
End If
%>
<%
Dim RecordsetSUBCategorias
Dim RecordsetSUBCategorias_cmd
Dim RecordsetSUBCategorias_numRows

Set RecordsetSUBCategorias_cmd = Server.CreateObject ("ADODB.Command")
RecordsetSUBCategorias_cmd.ActiveConnection = MM_overseaspr_STRING
RecordsetSUBCategorias_cmd.CommandText = "SELECT SubCategoryId, Category FROM dbo.OIC_Subcategory ORDER BY Category" 
RecordsetSUBCategorias_cmd.Prepared = true

Set RecordsetSUBCategorias = RecordsetSUBCategorias_cmd.Execute
RecordsetSUBCategorias_numRows = 0
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
function callfunctionDelTitulo(id) {
  var isGood=confirm('Esta seguro que quiere eliminar esta subcategoría?');
	if (isGood) {
	  //alert('true');
	} else {
	  return;
	}

	var tagname = "nombrecat" ;
	var newvalue = $('#'+tagname).val();
	
	window.parent.location.href = 'deletetitulo.asp?uid=' + id.toString() 
}
	
function callfunctionAddTitulo(id) {
	var tagname = "titulo" ;
	var newvalue = $('#'+tagname).val();
	if (newvalue > '') {
		window.parent.location.href = 'AddCatalogTitulo.asp?id=0' + '&newval=' + newvalue;
	} else { 
		$("#errormessage").val('Entre nombre de titulo o sección nueva!')
		$("#errormessage").show();
	}
}
	
function callfunction(loopid) {
	var tagname = "newTitulo" + loopid.toString();
	var newvalue = $('#'+tagname).val();
	var tagname0 = "oldTitulo" + loopid.toString();
	var oldvalue = $('#'+tagname0).val();
	if (newvalue > '') {
		window.parent.location.href = 'renametitulo.asp?newval=' + newvalue + '&oldval=' + oldvalue ;
	} else { 
		$("#errormessage").val('Entre nombre de título o sección nueva!')
		$("#errormessage").show();
	}
}

function updatecategory(loopid) {
	var oldtagname = "oldcat" + loopid.toString();
	var oldcatvalue = $('#'+oldtagname).val();
	var oldtagname0 = "oldsub" + loopid.toString();
	var oldsubvalue = $('#'+oldtagname0).val();

	alert(oldsubvalue)
	
	var tagname = "categoria" + loopid.toString();
	var catvalue = $('#'+tagname).val();
	var tagname0 = "subcategoria" + loopid.toString();
	var subvalue = $('#'+tagname0).val();
	if (catvalue > '') {
		window.parent.location.href = 'updatecategory.asp?cat=' + catvalue + '&sub=' + subvalue + '&oldcat=' + oldcatvalue + '&oldsub=' + oldsubvalue ;
	} else { 
		$("#errormessage").val('Categoria inválida!')
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
		"iDisplayLength":-1,
		"sPaginationType": "full_numbers",
		"bLengthChange":false,
		"bInfo":false,
		"bSort":false
	});	
	
	//oTable.fnSort( [ [1,'asc'], [4,'asc'], [3,'asc'] ] );
	
	$('#print').button();
	$('#back').button();
	
	$('#addcatb').button();
	$('.rename').button();
	
	$('a').button();
	
	$('#select_modelo').change(function()
 	{
    $('#formcategory').submit();
 	});
 
	$('#select_marca').change(function()
 	{
	 $('#select_modelo').val("0");
	 $('#formcategory').submit();
 	});
	
});
</script>
</head>

<body id="dt_example">
<div class="ui-widget" id="container" >
	<div class="css_left"  >
        <a id="back" name="back" href="part_search.asp">Catálogo de Internet OIC</a>
        <a id="catalog" name="catalog" href="catalog.asp">Catálogo de Imágenes</a>
	</div>

	
	<br clear="all">
	<h2 class="ui-widget-header" style="height:24px" > Catálogo ::&nbsp;<%= GetCatalog(CInt(Session("MM_CatalogID")))%> </h2>
	<br>
	
    
	<div class="ui-widget-content ui-corner-all"   >
	
      <form id="formcategory" method="post" >
		<div style="height:40px; vertical-align:baseline">       
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<label>Marca ...&nbsp;&nbsp;&nbsp;&nbsp;</label>
            <select id="select_marca"  name="select_marca" style="margin:5px;padding:5px;width:150px" class="ui-widget-header ui-corner-all" >
            
			<%
            While (NOT RecordsetCategorias.EOF)
            %>
            <option value="<%=(RecordsetCategorias.Fields.Item("CategoryId").Value)%>" 
			<%If(RecordsetCategorias.Fields.Item("CategoryId").Value = vcatid) Then Response.Write("selected=""selected""") : Response.Write("")%> >
			<%=UCASE(RecordsetCategorias.Fields.Item("Category").Value)%></option>
            <%
              RecordsetCategorias.MoveNext()
            Wend
            If (RecordsetCategorias.CursorType > 0) Then
              RecordsetCategorias.MoveFirst
            Else
              RecordsetCategorias.Requery
            End If
            %>
            </select>
		</div>
        <div style="height:40px; vertical-align:baseline">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<label>Modelo ...&nbsp;&nbsp;</label>
            <select id="select_modelo"  name="select_modelo" style="margin:5px;padding:5px;width:150px" class="ui-widget-header ui-corner-all" >
            <option value="0" selected >Todos</option>
			
			
            
			<%
            While (NOT RecordsetSubCategorias.EOF)
			%>
            <option value="<%=(RecordsetSubCategorias.Fields.Item("SubCategoryId").Value)%>" 
			<%If(RecordsetSubCategorias.Fields.Item("SubCategoryId").Value = vsubcatid) Then Response.Write("selected=""selected""") : Response.Write("")%> >
			<%=UCASE(RecordsetSubCategorias.Fields.Item("Category").Value)%></option>
            <%
              RecordsetSubCategorias.MoveNext()
            Wend
            If (RecordsetSubCategorias.CursorType > 0) Then
              RecordsetSubCategorias.MoveFirst
            Else
              RecordsetSubCategorias.Requery
            End If
            %>
            </select>
		</div>
	  	

        </form>
        
  	</div>
    
	<br><br>
       
    <table id="example" cellpadding="5" cellspacing="5" width="100%" >
    <thead>
    <tr>
    
    <th width="28%">Título o Sección</th>
     <th width="27%">Renombrar</th>
    <th width="23%">Marca</th>
    <th width="22%">Modelo</th>
   
    
    </tr>
    </thead>
    <tbody>
      <% 
	  li_loop = 0
	  While  (NOT Recordset1.EOF)
	  	li_loop = li_loop + 1
	   %>
        <tr>
         
         	<%
			stitulo = trim(Recordset1.Fields.Item("titulo").Value)
			%>
            
          <td align="center" valign="middle">
          <a style="font-size:14px;width:90%" href="routesubcat.asp?id=<%=(vcatid)%>&titulo=<%=(stitulo)%>" >
          <%=CStr(Recordset1.Fields.Item("titulo").Value)%>
          </a></td>
          
          <td>
	        <input name="oldTitulo<%=(CStr(li_loop))%>" type="hidden" class="ui-widget-header" id="oldTitulo<%=(CStr(li_loop))%>" value="<%=(stitulo)%>" >
            <input name="newTitulo<%=(CStr(li_loop))%>" type="text" class="ui-widget-header" id="newTitulo<%=(CStr(li_loop))%>" style="text-align:center;margin-left:5px" value="<%=(stitulo)%>" size="20" maxlength="20">
            <br>
            &nbsp;&nbsp;&nbsp;&nbsp;<input style="margin-top:3px" type="button" class="rename" id="Rename<%=(vsubcatid)%>" value="Cambiar nombre" onClick="callfunction('<%=(li_loop)%>')" >
            &nbsp;&nbsp;
            <input title="Eliminar Subcategoría" style="margin-top:3px" type="image" src="/images/del.png"  alt="Submit" onClick="callfunctionDelTitulo('<%=(stitulo)%>')" >

          </td>
          
          
          <td align="center" class="ui-widget-header"  >
          
          <input name="oldcat<%=(CStr(li_loop))%>" type="hidden" class="ui-widget-header" id="oldcat<%=(CStr(li_loop))%>" value="<%=(vcatid)%>" >
          <input name="oldsub<%=(CStr(li_loop))%>" type="hidden" class="ui-widget-header" id="oldsub<%=(CStr(li_loop))%>" value="<%=(vsubcatid)%>" >
                    
          <label>Marca</label>&nbsp;
          <select class="ui-widget-header" id="categoria<%=(CStr(li_loop))%>" name="categoria<%=(CStr(li_loop))%>" style="width:100px" >
            <%
            RecordsetCategorias.MoveFirst
            While (NOT RecordsetCategorias.EOF)
            %>
            <option value="<%=(vcatid)%>" <%If ( CStr(RecordsetCategorias.Fields.Item("CategoryId").Value) = CStr(vcatid)) Then Response.Write("selected=""selected""") : Response.Write("")%> >
				<%=(RecordsetCategorias.Fields.Item("Category").Value)%>
            </option>
            <%
              RecordsetCategorias.MoveNext()
            Wend
            If (RecordsetCategorias.CursorType > 0) Then
              RecordsetCategorias.MoveFirst
            Else
              RecordsetCategorias.Requery
            End If
            %>
          </select>
          
          
          <br><br>
          
         
          <label>Modelo</label>
          <select class="ui-widget-header" id="subcategoria<%=(CStr(li_loop))%>" name="subcategoria<%=(CStr(li_loop))%>" style="width:100px" >
            <%
            RecordsetSUBCategorias.MoveFirst
            While (NOT RecordsetSUBCategorias.EOF)
            %>
            <option value="<%=(vsubcatid)%>" <%if(vsubcatid)=RecordsetSUBCategorias("SubCategoryId")then response.write("selected"):response.write("")%>  >
            <%=(RecordsetSUBCategorias.Fields.Item("Category").Value)%>
            </option>
            <%
              RecordsetSUBCategorias.MoveNext()
            Wend
            If (RecordsetSUBCategorias.CursorType > 0) Then
              RecordsetSUBCategorias.MoveFirst
            Else
              RecordsetSUBCategorias.Requery
            End If
            %>
          </select>
        
          
          </td>
          
          <td class="ui-widget-header" align="center" >
          <input type="button" class="rename" id="Change<%=(li_loop)%>" value="Cambiar marca/modelo" onClick="updatecategory('<%=(li_loop)%>')" >
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
RecordsetCategorias.Close()
Set RecordsetCategorias = Nothing
RecordsetSUBCategorias.Close()
Set RecordsetSUBCategorias = Nothing

%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
