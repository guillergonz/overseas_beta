<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
' *** Logout the current user.
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

Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows
Dim vcatid,vcatalogid




' select_modelo will reload form
if Not IsNull(request("select_modelo")) AND Not IsEmpty(request("select_modelo")) AND LEN(request("select_modelo")) > 0 Then
	vsubcatid = Clng(request("select_modelo"))
	'Session("MM_SubCatID") = vsubcatid
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
<style type="text/css">
body,td,th {
	font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
}
</style>
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

	//alert(oldsubvalue)
	
	var tagname = "categoria" + loopid.toString();
	var catvalue = $('#'+tagname).val();
	var tagname0 = "subcategoria" + loopid.toString();
	var subvalue = $('#'+tagname0).val();

	var tagname = "oldTitulo" + loopid.toString();
	var titulo = $('#'+tagname).val();
	
	
	
	if (catvalue > '') {
		window.parent.location.href = 'updatecategory.asp?cat=' + catvalue + '&sub=' + subvalue + '&oldcat=' + oldcatvalue + '&oldsub=' + oldsubvalue + '&tit=' + titulo ;
	} else { 
		$("#errormessage").val('Categoria inválida!')
		$("#errormessage").show();
	}
}
		
jq = jQuery.noConflict();
$(document).ready(function () {
	
  
	jq('#print, #back, #addcatb, .rename').button();
	
	jq('#select_modelo').change(function()
 	{
    jq('#formcategory').submit();
 	});
 
	jq('#select_marca').change(function()
 	{
	  jq('#select_modelo').val("0");
	  jq('#formcategory').submit();
 	});
	
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
    
<h4>Mantenimiento de Catálogo ::&nbsp;<%= GetCatalog(CInt(Session("MM_CatalogID")))%> </h4>
	
   
    
     <div class="breadcrumb" >
		<a title="Regresar a lista de marcas" href="/catalog_maint_cat.asp" >
		<%= GetNombreCategoria(Session("MM_CatID"))%></a>
     	&nbsp;&nbsp;/&nbsp;&nbsp;
     	<a id="backtomodel" name="backtomodel" title="Regresar a lista de modelos" href="/catalog_maint2.asp" >
        <%= GetNombreSubCategoria(vcatid,vsubcatid)%></a>
        
    </div>
    
  
    
    
	
	
      <form id="formcategory" method="post" >
      <div class="row">
       
		<div class="col-xs-2">      
			<label>Marca ...&nbsp;&nbsp;&nbsp;&nbsp;
            <select disabled id="select_marca"  name="select_marca"  class="form-control"   >
            
            
			<%
            While (NOT RecordsetCategorias.EOF)
            %>
            <option  value="<%=(RecordsetCategorias.Fields.Item("CategoryId").Value)%>" 
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
            </label>
            
		</div>
        <div class="col-xs-2">
			<label>Modelo ...
                &nbsp;&nbsp;<select disabled  id="select_modelo"  name="select_modelo"  class="form-control"  >
                
                <%
                While (NOT RecordsetSubCategorias.EOF)
                %>
                <option  value="<%=(RecordsetSubCategorias.Fields.Item("SubCategoryId").Value)%>" 
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
            </label>
		</div>
      
        
        <% if Session("MM_UserName") = "Z099" then %>    
        
        	<!--<div class="col-xs-2">
	  	    
                <label>&nbsp;&nbsp;Nombre de marca seleccionada (Brand)
                &nbsp;&nbsp;<input title="Seleccione para ir a esta marca de auto (Click)" onClick="document.location.href='catalog_maint_SUB.asp'" name="nombrecat" type="text" readonly class="btn-btn-primary" id="nombrecat" value="<= GetCategoryName(CInt(vcatid))%>"  maxlength="50" >
                </label>
        	</div>-->
        
        	<div class="col-xs-3">
                <label>&nbsp;&nbsp;Nombre o título de nueva sección
                    &nbsp;&nbsp;<input title="Para entrar un nuevo modelo entre el nombre y oprima Añadir modelo " class="form-control" id="titulo" name="titulo" type="text"  maxlength="50" >
                </label>
        	</div>
            
        	<div class="col-xs-2">    
            	<br>
                <input name="addcatb" class="btn btn-primary" type="button" id="addcatb"  onClick="callfunctionAddTitulo(0)" value="Añadir título/sección" >   
        	</div>  
    
  		
                   	
       		
    
    	</div>   
    	
   		
        
   		<% end if %>  
    
	
    	<div class="row">
        
      		<input disabled class="col-xs-12 text-warning"  id="errormessage"  value="<%= Session("MM_Update") %>" >     
      	    
        </div>
    
    </form>
    
    <div class="row">         	
          <br><br>  
            
    </div>  
       
	
    
    <div class="table-responsive">
      <table  class="table table-striped" id="example"   >
    <thead>
    <tr>
    
    <th>Título</th>
    <% if Session("MM_UserName") = "Z099" then %> 
    <th ></th>
    <th >Renombrar</th>
    <th >Marca / Modelo</th>
    <th >&nbsp;</th>
    <% End If %>
    
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
          <a class="btn btn-primary" href="routesubcat.asp?id=<%=(vcatid)%>&titulo=<%=(stitulo)%>" style="width:200px" >
          <%=CStr(Recordset1.Fields.Item("titulo").Value)%>
          </a></td>
          
         
          <td>
	        <input name="oldTitulo<%=(CStr(li_loop))%>" type="hidden" class="form-control" id="oldTitulo<%=(CStr(li_loop))%>" value="<%=(stitulo)%>" >
            <input name="newTitulo<%=(CStr(li_loop))%>" type="text" class="form-control" id="newTitulo<%=(CStr(li_loop))%>"  value="<%=(stitulo)%>"  maxlength="20">
          </td>
          
           <% if Session("MM_UserName") = "Z099" then %> 
          <td>
            
            <input  type="button" class="btn btn-primary" id="Rename<%=(vsubcatid)%>" value="Cambiar nombre" onClick="callfunction('<%=(li_loop)%>')" >
          	<br>
            <input title="Eliminar Subcategoría <%=Recordset1("titulo")%>" class="btn btn-danger" type="image" src="/images/del.png"  alt="Submit"   onClick="callfunctionDelTitulo('<%=Recordset1("titulo")%>')" >
          </td>
                    
          <td align="center"   >
          
          <input name="oldcat<%=(CStr(li_loop))%>" type="hidden" class="form-control" id="oldcat<%=(CStr(li_loop))%>" value="<%=(vcatid)%>" >
          <input name="oldsub<%=(CStr(li_loop))%>" type="hidden" class="form-control" id="oldsub<%=(CStr(li_loop))%>" value="<%=(vsubcatid)%>" >
          
          <select class="form-control" id="categoria<%=(CStr(li_loop))%>" name="categoria<%=(CStr(li_loop))%>"  >
            <%
            While (NOT RecordsetCategorias.EOF)
            %>
            <option value="<%=(RecordsetCategorias("CategoryId"))%>" <%If ( CStr(RecordsetCategorias.Fields.Item("CategoryId").Value) = CStr(vcatid)) Then Response.Write("selected=""selected""") : Response.Write("")%> >
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
          <br>
          
          <select class="form-control" id="subcategoria<%=(CStr(li_loop))%>" name="subcategoria<%=(CStr(li_loop))%>" >
            <%
            While (NOT RecordsetSUBCategorias.EOF)
            %>
            <option value="<%=(RecordsetSUBCategorias("SubCategoryId"))%>" <%if(vsubcatid)=RecordsetSUBCategorias("SubCategoryId")then response.write("selected"):response.write("")%>  >
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
          
          <td  align="center" >
          <input type="button" class="btn btn-primary" id="Change<%=(li_loop)%>" value="Cambiar marca / modelo" onClick="updatecategory('<%=(li_loop)%>')" >
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

'RecordsetCategorias.Close()
'Set RecordsetCategorias = Nothing

'RecordsetSUBCategorias.Close()
'Set RecordsetSUBCategorias = Nothing
%>
