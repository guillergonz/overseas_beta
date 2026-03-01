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


Dim vcatid,vsubcatid,vcatalogid,vtitulo



' siempre viene de catalog_maint2.asp
If IsNull(Session("MM_Titulo")) or IsEmpty(Session("MM_Titulo")) or trim(Session("MM_Titulo"))="" Then
	response.redirect("catalog_maint2.asp")
else
	vtitulo = Session("MM_Titulo")
End IF

'If IsNull(Session("MM_Titulo")) or IsEmpty(Session("MM_Titulo")) or trim(Session("MM_Titulo"))="" Then
'	'response.redirect("index.asp")
'
'else
'
'	if IsNull(Session("MM_Titulo")) or IsEmpty(Session("MM_Titulo")) or LEN(Session("MM_Titulo")) = 0 Then
'		
'		vtitulo = GetTituloDefault(vcatalogid,vcatid)
'		Session("MM_Titulo") = vtitulo
'	else
'		vtitulo = Session("MM_Titulo")
'	end if
'	
'End If

If IsNull(Session("MM_CatID")) or IsEmpty(Session("MM_CatID")) or trim(Session("MM_CatID"))="" or trim(Session("MM_CatalogID"))="" or IsNull(Session("MM_CatalogID")) or IsEmpty(Session("MM_CatalogID")) Then
	
	response.redirect("index.asp")
	
else
	
	vcatalogid 	= Session("MM_CatalogID")
	vcatid 		= Session("MM_CatID")
	vsubcatid 	 = Session("MM_SubCatID") 

End If
	
'If IsNull(Session("MM_PartNo")) or IsEmpty(Session("MM_PartNo"))or trim(Session("MM_PartNo")) = "" Then
	
	vpartno = GetPartNoDefault(vcatalogid,vcatid,vtitulo)
	Session("MM_PartNo") = vpartno

'Else

'	vpartno = Session("MM_PartNo")

'End If

If IsNull(Session("MM_PartNo")) or IsEmpty(Session("MM_PartNo"))or trim(Session("MM_PartNo")) = "" Then
	' NEW ITEM ?
	'response.redirect("index.asp")
	
	
	Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
	Recordset1_cmd.ActiveConnection = MM_overseaspr_STRING
	Recordset1_cmd.CommandText = "SELECT * FROM dbo.OIC_PartsPerCategory WHERE idcatalog = " + CStr(vcatalogid) + " AND categoryid = " + CStr(vcatid) + " AND subcategoryid = " + CStr(vsubcatid)   
	Recordset1_cmd.Prepared = true
	Set Recordset1 = Recordset1_cmd.Execute
	Recordset1_numRows = 0
	
	Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
	Recordset2_cmd.ActiveConnection = MM_overseaspr_STRING
	Recordset2_cmd.CommandText = "SELECT distinct titulo FROM dbo.OIC_PartsPerCategory WHERE idcatalog = " + CStr(vcatalogid) + " AND categoryid = " + CStr(vcatid) + " AND subcategoryid = " + CStr(vsubcatid)  
	Recordset2_cmd.Prepared = true
	Set Recordset2 = Recordset2_cmd.Execute
	Recordset2_numRows = 0
	
	
else
	
	Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
	Recordset1_cmd.ActiveConnection = MM_overseaspr_STRING
	Recordset1_cmd.CommandText = "SELECT * FROM dbo.OIC_PartsPerCategory WHERE idcatalog = " + CStr(vcatalogid) + " AND categoryid = " + CStr(vcatid) + " AND subcategoryid = " + CStr(vsubcatid) + " AND partno = '" + vpartno + "' ;"  
	Recordset1_cmd.Prepared = true
	Set Recordset1 = Recordset1_cmd.Execute
	Recordset1_numRows = 0
	
	Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
	Recordset2_cmd.ActiveConnection = MM_overseaspr_STRING
	Recordset2_cmd.CommandText = "SELECT distinct titulo FROM dbo.OIC_PartsPerCategory WHERE idcatalog = " + CStr(vcatalogid) + " AND categoryid = " + CStr(vcatid) + " AND subcategoryid = " + CStr(vsubcatid)  
	Recordset2_cmd.Prepared = true
	Set Recordset2 = Recordset2_cmd.Execute
	Recordset2_numRows = 0
	
End if
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
<script src="jquery-ui-1.12.0-rc.1/external/jquery/jquery.js"></script>
<link rel="stylesheet" type="text/css" href="jquery-ui-1.12.0-rc.1/jquery-ui.theme.min.css">
<link rel="stylesheet" type="text/css" href="jquery-ui-1.12.0-rc.1/jquery-ui.structure.min.css">
<script src="jquery-ui-1.12.0-rc.1/jquery-ui.js"></script>
<link rel="stylesheet" type="text/css" href="jquery-ui-1.12.0-rc.1/jquery-ui.min.css">
<link rel="stylesheet" type="text/css" href="jquery-ui-themes-1.12.0-rc.1/themes/redmond/jquery-ui.min.css">



 
    
<style type="text/css" title="currentStyle">

	#form1 li {
		border: 1px solid #666666;
		
		
		height:28px;
		list-style: none;
		margin: 10px .2em 0 0;
		
		text-decoration: none;
		width:200px;
		
		-webkit-border-radius: 3px;
		-moz-border-radius: 3px;
		border-radius: 3px;
	}
	
	

	#form1 li a {
		margin:2px;
		padding:4px;
		width:190px;
		
		
		
		-webkit-border-radius: 4px;
		-moz-border-radius: 4px;
		border-radius: 4px;
	}
	#form1 li a,
	#form1 li a:link,
	#form1 li a:visited {
		
		text-decoration: none;
		width:100%;
	}
	#form1 li a:hover {
		border: 1px solid #59b4d4;
		width:200px;
		text-decoration: none;
	}
	#form1 li a:active {
		border: 1px solid #ffaf0f;
		
	}
	#form1 li a:focus {
		border: 1px solid #59b4d4;
		width:200px;
		background-color:white;
	}
	#form1 li a:highlight {
		border: 1px solid #cccccc;
		background-color:white;
	}
	
	
	
	body,td,th {
	font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
}
</style>


        
<script type="text/javascript" charset="utf-8">

jq = jQuery.noConflict();
jq('#ajaxmessage1').show();


function newstatus(status) {
	
	jq('#ajaxmessage1').show();
	
	var src
	src = '/parts_images/' + (status) + '.jpg'
	jq.ajax({
		url:src,
		type:'HEAD',
		error: function()
		{
			src = '/parts_images/' + (status) + '.png'
			jq.ajax({
			url:src,
			type:'HEAD',
			error: function()
			{
	
				//file not exists
				//alert('Not Found ' + src);
				jq("#ImagenSeccion").attr("src", "#");
	
			},
			success: function()
			{
			  jq("#ImagenSeccion").attr("src", src);
			}
			});
		
		},
		success: function()
		{
		  jq("#ImagenSeccion").attr("src", src);
		}
	});
	
	

	jq("#titulo").html(status);
	jq("#addtitulo").val(status);
	
	jq('#upload').prop('title', 'Upload imágen de ' + status);
	jq('#upload').prop('href', 'routecatuploader.asp?id=' + status);	
	
	
	jq("#addnota").val("");
	jq("#addpartno").val("");
	
	titulo = encodeURIComponent(jq("#addtitulo").val());
	vdata = "titulo=" + titulo
	
	jq.ajax({
			type:"POST",
			url: "/addpart.asp",
			context: document.body,
			data: "titulo=" + titulo,
			beforeSend: function() {jq('#ajaxmessage1').show();},
			success: function(outputhtml){
				jq("#ajaxdiv").html(outputhtml);
				jq('#ajaxmessage1').fadeOut("slow" );		
				
				
				
		
								
		   },
		   error: function(xhr,textStatus, errorThrown){
				jq('#ajaxdiv').html(textStatus);
				jq("#ajaxdiv").show('');
		   }            
		});	
		
	
	
	
	
	
		
}
function ajaxdelpart(uid){
	if (uid > 0) {
		jq.ajax({
			type:"POST",
			url: "/addpart.asp",
			context: document.body,
			data: "action=DEL" + "&uid=" + uid,
			beforeSend: function() {jq('#ajaxmessage1').show();},
			success: function(outputhtml){
				jq("#ajaxdiv").html(outputhtml);
				
				window.location.href = "catmaint.asp";

				jq('#ajaxmessage1').fadeOut("slow" );					
		   },
		   error: function(xhr,textStatus, errorThrown){
				jq('#ajaxdiv').html(textStatus);
				jq("#ajaxdiv").show('');
		   }               
		});	
	}
}
function ajaxaddpart(){
	var vaddpartno,vtitulo,vnota
	vaddpartno = encodeURIComponent(document.getElementById("addpartno").value);
	vtitulo = encodeURIComponent(document.getElementById("addtitulo").value);
	vnota = encodeURIComponent(document.getElementById("addnota").value);
	
	if (vaddpartno > "" && vtitulo > "" && vnota > "") {
		jq.ajax({
			type:"POST",
			url: "/addpart.asp",
			context: document.body,
			data: "action=ADD" + "&partno=" + vaddpartno + "&titulo=" + vtitulo + "&nota=" + vnota,
			beforeSend: function() {jq('#ajaxmessage1').show();},
			success: function(outputhtml){
				jq("#ajaxdiv").html(outputhtml);
				jq('#ajaxmessage1').fadeOut("slow" );						
		   },
		   error: function(xhr,textStatus, errorThrown){
				jq('#ajaxdiv').html(textStatus);
				jq("#ajaxdiv").show('');
		   }               
		});	
	} else { alert("Título y/o pieza es requerido"); }
	
	
}


	
jq(document).ready(function () {
	
	
	//imagePreview();	
	
	jq('#ajaxmessage1').show();
		
	jq('#print').button();
	jq('#back').button();
	jq('#catalog1').button();
	jq('#catalog2').button();
	
	jq("#ajaxdiv").html("");

	jq("#addpart").on("click",ajaxaddpart);
	jq('#addpart').button();
	
	// GGG 02/16/2016 image_exist not populated	
	//$(function() {
//		var availabletags = <= availabletags() %>;
//		$( "#addpartno" ).autocomplete({
//			source: availabletags,
//			minLength: 5
//		});
//	});
	var status = "<%= Session("MM_Titulo")%>"
	//status = encodeURIComponent(status);
	newstatus(status);
	jq('#upload').prop('title', 'Upload imágen de ' + status);	
	jq('#upload').prop('href', 'routecatuploader.asp?id=' + status);	
		
	// entry screen	
	jq('#menu li').click(function() {
		jq('#menu li a').removeClass('ui-state-active');
		jq('a', this).addClass('ui-state-active');
	});
	
		
	jq('#ajaxmessage1').fadeOut("slow" );
	
	jq("#container2").show();
	
	jq("#ajaxdiv").show();
	
	jq("#panel").show();
	
	
	
	
});
</script>
</head>

<body id="dt_example" >
<div class="container"    >
	
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
   
   
   	 <div id="ajaxmessage1" name="ajaxmessage1" class="ajaxmessage1" style="float:right;margin-right:50px">
        <img src="images/ajax-loader.gif" alt="image loader" width="16" height="16">
     </div>
        
  <div id="breadcrumb" style="width:90%">
      <div class="breadcrumb" >
        <a title="Regresar a lista de marcas" href="/catalog_maint_cat.asp" >
        <%= GetNombreCategoria(Session("MM_CatID"))%></a>
        &nbsp;&nbsp;/&nbsp;&nbsp;
        <a id="backtomodel" name="backtomodel" title="Regresar a lista de modelos" href="/catalog_maint2.asp" >
        <%= GetNombreSubCategoria(vcatid,vsubcatid)%></a>
        </div>    
      
     
    </div>
        
    
    
 
    
   
    
    <form id="form1" name="form1" method="post"  >
        <div class="table-responsive">
            <table class="table table-striped" >
              <tr>
                <td width="206px" valign="top" align="right">
                  <ul id="menu"   >
                    <%
			while not (Recordset2.eof) 
                If LEN(Recordset2("titulo")) > 0 Then 
			%>
            
            	<% if Session("MM_Titulo") = (Recordset2.Fields.Item("titulo").Value) then %>

                    <li class="ui-state-default" style="font-size:12px" ><a title="<%=(Recordset2("titulo"))%>" id="<%=(Recordset2.Fields.Item("titulo").Value)%>" name="<%=(Recordset2.Fields.Item("titulo").Value)%>"  onClick="javascript:newstatus('<%=(Recordset2.Fields.Item("titulo").Value)%>')" class="ui-state-active" ><%=(Recordset2.Fields.Item("titulo").Value)%></a></li>
                
                <% else %>

                    <li class="ui-state-default" style="font-size:12px" ><a title="<%=(Recordset2("titulo"))%>" id="<%=(Recordset2.Fields.Item("titulo").Value)%>" name="<%=(Recordset2.Fields.Item("titulo").Value)%>"  onClick="javascript:newstatus('<%=(Recordset2.Fields.Item("titulo").Value)%>')"  ><%=(Recordset2.Fields.Item("titulo").Value)%></a></li>
                    
				<% end if %>
                	
				<% 
					End If
				   Recordset2.MoveNext()
				 Wend
				%>
                  </ul>
                </td>
                <td width="538" valign="top">
                
                
                <%
				
				
				
				Dim objFSO
            	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			
            Dim PartNumJPG
            PartNumJPG = Server.MapPath("./parts_images/" & (vpartno) & ".png")
				
				
				%>
                
				<% If objFSO.FileExists( PartNumJPG ) Then %>
                
                	<div class="ui-widget-content"><img   id="ImagenSeccion"  name="ImagenSeccion" src="/parts_images/<%= vpartno %>.png" width="499" height="550" alt=""></div>
                
                <% else %>
				
					<div class="ui-widget-content"><img   id="ImagenSeccion"  name="ImagenSeccion" src="#" width="499" height="550" alt=""></div>
				
				<% end if %>
                
                
                <% Set objFSO = Nothing %>
            
                
                </td>
                <td width="263" valign="top">
                    
                  <% if UCase(Session("MM_UserName")) = "Z099" then %>
                    
                         <p style="margin-top:10px">    
                        
                        <a style="float:right"  id="upload" name="upload" href="routecatuploader.asp?id=<%=Session("MM_Titulo")%>">
                        <img src="/images/upload.png" width="25"  height="25" alt="catupload"  >
                        </a>
                        
                        <input  id="addtitulo" name="addtitulo"  type="text"   value="<%=Session("MM_Titulo")%>" class="ui-corner-all" style="padding:3px;width:85%" >
                        
                      </p>
                      
                      <table border="0" cellpadding="0" cellspacing="0" width="100%" >
                        <tr>
                          <td height="30">&nbsp;No. de Pieza</td>
                          <td> <input  id="addpartno" name="addpartno" class="ui-widget" type="text" size="12" maxlength="50" value="" ></td>
                        </tr>
                        <tr>
                          <td height="30">&nbsp;Id de Pieza (1)</td>
                          <td><input  id="addnota" name="addnota" class="ui-widget" type="text" size="12" maxlength="50" value="" ></td>
                        </tr>
                        <tr>
                          <td height="50" colspan="2">
                            <input type="button" id="addpart" name="addpart" value="Añadir pieza ..." style="margin-bottom:10px;width:90%;margin-left:10px" >
                          </td>
                        </tr>
                      </table>
                      
                 
                    
                  <br> 
                    
                    
                    
                  <% Else %>
                  <p>
                    <input  id="addtitulo" name="addtitulo"  type="text" size="18" maxlength="50" value="<%=Session("MM_Titulo")%>" class="ui-widget-header"  style="margin-left:2px;text-align:center"  >
                  </p>        
                    
                  <% End IF %>
                    
                    
                  <div id="ajaxdiv" class="ajaxdiv ui-widget-content ui-corner-all" style="display:none"  ></div>
                    
                  <%
            if session("MM_Update_Message") = "OK" then
				response.write( "<p>" + lang("item_added") + "</p>" )
			End If
            %>
                </td>
                <td width="176">&nbsp;</td>
              </tr>
            </table>
      </div>

           
         
       
  
  
  
         
  </form>
    
    
    
</div>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
session("MM_Update_Message") = ""
%>