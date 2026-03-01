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
Set oRS2 = Server.CreateObject ("ADODB.Command")
oRS2.ActiveConnection = MM_overseaspr_STRING
If Session("lang") = "S" Then
	oRS2.CommandText = "SELECT field_1, field_2, english_desc FROM dbo.familicat WHERE field_2 in ( select distinct field_6 from partmst1_distinct) ORDER BY field_1 ASC" 
Else
	oRS2.CommandText = "SELECT field_1, field_2, english_desc FROM dbo.familicat WHERE field_2 in ( select distinct field_6 from partmst1_distinct) ORDER BY english_desc ASC" 
End if
oRS2.Prepared = true
Set dd_category = oRS2.Execute
%>
<!doctype html>
<head>
<meta charset="utf-8"> 
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" >
<link rel="icon" href="images/favicon.ico" type="image/x-icon" >
<title>Overseas Import Corporation</title>

<!-- Bootstrap -->
<link rel="stylesheet" type="text/css" href="bootstrap-3.3.6-dist/css/bootstrap.min.css">

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script type="text/javascript" charset="utf-8" src="bootstrap-3.3.6-dist/jquery.min.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->

<link href="overseas.css" rel="stylesheet" type="text/css" >


<style type="text/css">
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
.wordwrap {
    word-break: break-word;
    display: inline-block;
}
body,td,th {
font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
font-size: 14px;
color: #000;
}

</style>


<script type="text/javascript" charset="utf-8">

function DisplayParts(){	
	var src="AjaxDisplayParts.asp";	
	$.ajax({
		type:"GET",
		url:src,
		context: document.body,
		data: "action=VIEW",
		beforeSend: function() {
			$("#loader").fadeIn("slow");
			$("#loader").removeClass("display_no").addClass("display_yes");
			},
		
		success: function(outputhtml){
			$("#DisplayParts").html(outputhtml);
			$("#DisplayParts").className = "display_yes" ;
			$("#loader").fadeOut("fast");
			
		},
		error: function(xhr,textStatus, errorThrown){
			alert(textStatus);
		}               
	});				

};

function AddToCart(str,amount){	
	var src="add_to_shopping_cart.asp?p=" + str + "&a=" + amount + "&sid="+Math.random();	
	$.ajax({
		type:"GET",
		url:src,
		context: document.body,
		data: "action=ADD",
		beforeSend: function() {
			$("#loader").fadeIn("slow");
			$("#loader").removeClass("display_no").addClass("display_yes");
		},
		
		success: function(outputhtml){
			$("#shopping_cart").html(outputhtml);
			$("#shopping_cart").className = "display_yes" ;
			$("#loader").fadeOut("fast");
			
			// Ajax to prevent adding same part
			DisplayParts();	
			
		},
		error: function(xhr,textStatus, errorThrown){
			alert(textStatus);
		}               
	});				

};


function AddToCart2(str,amount){	
	var src="add_to_shopping_cart.asp?p=" + str + "&a=" + amount + "&sid="+Math.random();	
	$.ajax({
		type:"GET",
		url:src,
		context: document.body,
		data: "action=ADD",
		beforeSend: function() {
			$("#loader").fadeIn("slow");
			$("#loader").removeClass("display_no").addClass("display_yes");
		},
		
		success: function(outputhtml){
			$("#shopping_cart").html(outputhtml);
			$("#shopping_cart").className = "display_yes" ;
			$("#loader").fadeOut("fast");
			
			// Ajax to prevent adding same part
			//DisplayParts();	
			
		},
		error: function(xhr,textStatus, errorThrown){
			alert(textStatus);
		}               
	});				

};

function DelFromCart(str) { 

	var src="del_p_shopping_cart.asp?p=" + str + "&t=1&sid="+Math.random();	
	$.ajax({
		type:"GET",
		url:src,
		context: document.body,
		data: "action=DEL",
		beforeSend: function() {
			$("#loader").fadeIn("slow");
			$("#loader").removeClass("display_no").addClass("display_yes");
			},
		
		success: function(outputhtml){
			$("#shopping_cart").html(outputhtml);
			
			$("#loader").removeClass("display_yes").addClass("display_no");		
			$("#loader").fadeOut("fast");
			
			// Ajax to prevent adding same part
			DisplayParts();	
			
		},
		error: function(xhr,textStatus, errorThrown){
			alert(textStatus);
		}               
	});				

};		
</script>
															      
<script type="text/javascript" charset="utf-8">
function ValidateCart(part,amt) {
	// cant_ordenar_
	var hidefield = 'add_cart_' + amt.substr(13);
	
	$("#" + hidefield).hide();
	amount = document.getElementById(amt).value;
	AddToCart2(part,amount);
};
</script>

<script type="text/javascript" charset="utf-8">
function searchSel() {
  var input=document.getElementById('fks').value.toUpperCase();
  var output=document.getElementById('fcs').options;

  for(var i=0;i<output.length;i++) {
	if(output[i].value.indexOf(input)==0){
	  output[i].selected=true;
	  };
	if(document.forms[0].fks.value==''){
	  output[0].selected=true;
	  };
  }
};
</script>

<script type="application/javascript" language="javascript" >
$(document).ready(function() {
	
	$(function(){
		var options = {};
		$("#loader").fadeOut("slow");
	});
  		
});
</script>

</head>
<body>
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
                     <li class="active"><a href="part_search.asp" title="<%= Lang("buscar") %>" target="_self"><%= Lang("buscar") %></a></li>
                
                     <li><a id="completar" href="cart.asp"  title="<%= Lang("completar_orden")%>" target="_self" ><%= Lang("completar_orden") %></a></li>
                     <li><a href="catalog_maint_CAT.asp" title="<%= Lang("catalogo") %> Catalog" target="new"><%= Lang("catalogo") %></a></li>
                    
                    
                    <% If LEN(Session("MM_Multi_Username")) = 0 or Session("MM_Multi_Username") = "CARLE BETANCOURT" then %>
                    
                    	<li><a href="account_statement_iframe.asp" title="<%= Lang("estado_de_cuenta_actual") %>" target="_self"><%= Lang("estado_de_cuenta_actual") %></a></li>
                        
                    <% End IF %>
            
            
                    
                    <% if ucase(Session("MM_UserName")) = "Z099" then %>
                        
                        <li><a href="uploadDataDaily.asp" title="Import parts" target="_self">Import parts</a></li>
                        <li><a href="monitor_beta.asp" title="Monitor" target="_self">Monitor  <span class="badge"><%= GetCurrentSales(Session("MM_UserName")) %> </span></a></li>
                        <li><a href="usuarios.asp" title="Usuarios" target="_self">Usuarios</a></li>
                        
                        <li id="show" class="active" style="height:50px"><%= CheckTime()%></li>
                        
                    <% End If %>
                </ul>
            </div>
        </div>
    </nav>
    

    <a class="navbar-brand" rel="home" href="#" title="Overseas Import Corporation">
        			<img style="max-width:256px; margin-top: -7px;" src="/images/oiclogo2.gif"></a>
                    
	<form id="search_part2" name="search_part2"  action="part_search.asp" method="post" >
	   
   	<div class="table-responsive pull-right col-xs-12 col-sm-2 col-md-2 col-lg-2" style="overflow-y:hidden;overflow-x:hidden;margin-right:3px;padding:0px" >   
        
        <img class="pull-right" style="margin-right:90px" id="loader" name="loader" src="images/ajax-loader.gif" width="16" height="16" alt="loader" >
              
		<% If lcase(Request.ServerVariables("SCRIPT_NAME")) = "/overseaspr/part_search.asp" or lcase(Request.ServerVariables("SCRIPT_NAME")) = "/part_search.asp" Then %>
            <br>
            <div class="grayback">
                <%= Lang("preparado") %>
            </div>
            <br>
            <div id="shopping_cart">
                <% CartDisplay("N") %>
            </div>
        <% End if %>
        
        <br>
        
        
               
	</div>
  
    <div class="panel panel-default col-xs-12 col-sm-8 col-md-8 col-lg-9" >
   		<a href="<%= MM_Logout %>" class="btn btn-danger pull-right" style="margin-left:6px" ><%= Lang("salir") %></a>
        
        <% if ucase(Session("MM_Username")) = "B001" AND ucase(Session("MM_Multi_Username")) = "CARLE BETANCOURT" then %>
        <a href="multipass.asp" class="btn btn-danger pull-right" style="margin-left:6px" >Cambiar contraseña</a>
        <% End If %>
        
   		<!--<a id="completar" href="../cart.asp" class="btn btn-default btn-sm pull-right"><= Lang("completar_orden") %></a>-->
        <div class="panel-heading">
            <%= GetUserName(Session("MM_Username"))%>
            &nbsp;<%= lang("fecha")%>&nbsp;<%=Date()%> 
            <% if LEN(Session("MM_Multi_Username")) > 0 then %>
            <p><%= Session("MM_Multi_Username") %></p>
            <% End If %>               
        </div>
        <div class="panel-body">
            <div class="row" >
            	<div class="col-xs-12 col-sm-12 col-md-2 col-lg-2" >    
            		<%= Lang("entre_5_piezas") %> 
        		</div>
                <div class="col-xs-12 col-sm-12 col-md-2 col-lg-2" style="padding:2px">    
            		<input name="p1" type="text" class="form-control" id="p1" size="9" maxlength="15" style="padding:2px" > 
        		</div>
                <div class="col-xs-12 col-sm-12 col-md-2 col-lg-2" style="padding:2px">    
            		<input name="p2" type="text" class="form-control" id="p2" size="9" maxlength="15" style="padding:2px" > 
        		</div>
                <div class="col-xs-12 col-sm-12 col-md-2 col-lg-2" style="padding:2px">    
            		<input name="p3" type="text" class="form-control" id="p3" size="9" maxlength="15" style="padding:2px" > 
        		</div>
                <div class="col-xs-12 col-sm-12 col-md-2 col-lg-2" style="padding:2px">    
            		<input name="p4" type="text" class="form-control" id="p4" size="9" maxlength="15" style="padding:2px" > 
        		</div>
                <div class="col-xs-12 col-sm-12 col-md-2 col-lg-2" style="padding:2px">    
            		<input name="p5" type="text" class="form-control" id="p5" size="9" maxlength="15" style="padding:2px" > 
        		</div>
            </div>	
        </div> 
	</div>
    
    <div class="panel panel-default col-xs-12 col-sm-8 col-md-8 col-lg-9" >    
        
        <div class="panel-body">
            <div class="row" >
            	<div class="col-xs-12 col-sm-12 col-md-3 col-lg-3 small"> 
                    
                    <%= Lang("family_keyword") %>
                    <br>
                    <input class="form-control" type="text"  id="fks"  onKeyUp="searchSel()" >
                                   
                </div>
             	<div class="col-xs-12 col-sm-12 col-md-3 col-lg-3 small">
					<%= Lang("family_category") %><br>
                    <select class="form-control" name="fcs"  id="fcs" >
                    <option selected value="">&nbsp;</option> 
                    <% While (NOT dd_category.EOF) %>
                    <% If Session("lang") = "S" Then %>
                        <option value="<%=trim(dd_category.Fields.Item("field_1").Value)%>" ><%=trim(dd_category.Fields.Item("field_1").Value)%></option>
                    <% Else %>
                        <option value="<%=trim(dd_category.Fields.Item("english_desc").Value)%>" ><%=trim(dd_category.Fields.Item("english_desc").Value)%></option>
                    <% End if %>
                    <% dd_category.MoveNext()
                    Wend
                    If (dd_category.CursorType > 0) Then
                    dd_category.MoveFirst
                    Else
                    dd_category.Requery
                    End If %>
                    </select>
             	</div>
             	<div class="col-xs-12 col-sm-12 col-md-3 col-lg-3 small">
					<%= Lang("modelo") %>
                    <br><input class="form-control" name="model" type="text"  id="fs2" >
             	</div>
                <div class="col-xs-12 col-sm-12 col-md-1 col-lg-1 small">
					
					<% 'if ucase(Session("MM_UserName")) = "Z099" then %>
                    	<%= Lang("Specials") %>
                    	<br><input name="special_items" type="checkbox" class="form-control" id="special_items" >
                	<% 'end if %>
                
                </div>
    			<div class="col-xs-12 col-sm-12 col-md-2 col-lg-2 small">
    				<br><input class="btn btn-primary" name="submit" type="submit" id="submit" value="<%=Lang("buscar")%>" >
    			</div>
            </div>
        </div> 			
	</div>
 
 
 
      
    <div class="row">         
        <div class="col-xs-12 col-lg-9 pull-left" >
            <div class="table-responsive">
            
                <table class="table"  >
                <tr>
                <td style="border-top-color:transparent" >        
                
				<%
                model = Trim(Ucase(Request("model")))
                family_keyword_list = Trim(FamilyCode(Trim(Request("fcs"))))
                special_items = Trim(Ucase(Request("special_items")))
                liquidation_items = Trim(Ucase(Request("liquidation_items")))
                If special_items = "ON" then
                	sql_select_price = "SELECT field_1, field_2, field_4, field_5, familia_descripcion, english_version, 'N/A' as replacement , family_description , fam_make_item FROM dbo.partmst1_distinct WHERE field_1 in ( select dbo.prespecials.specials from dbo.prespecials ) "
                elseIf liquidation_items = "ON" then
                	sql_select_price = "SELECT field_1, field_2, field_4, field_5, familia_descripcion, english_version, 'N/A' as replacement , family_description , fam_make_item FROM dbo.partmst1_distinct WHERE field_1 in ( select field_1 from dbo.prod_liqui ) "
                else
					If Session("user_level") = 3 Then
					sql_select_price = "SELECT field_1, field_2, field_3, field_5, familia_descripcion, english_version, replacement, family_description , fam_make_item FROM dbo.partmst1_distinct WHERE  "
					Else
					sql_select_price = "SELECT field_1, field_2, field_4, field_5, familia_descripcion, english_version, replacement , family_description , fam_make_item FROM dbo.partmst1_distinct WHERE "
					End if
                End if
				
                If len(model) = 0 or special_items = "ON" or liquidation_items = "ON" Then
					sql_parts 	= ""
					sql_group 	= ""
					sql_family	= ""
					p1	= trim(Request("p1"))
					p2	= trim(Request("p2"))
					p3	= trim(Request("p3"))
					p4	= trim(Request("p4"))
					p5	= trim(Request("p5"))
					order_count = 0
					' CONCATENATE PART FOR WHERE AND GROUP BY CLAUSE, ONLY IF HAD VALUE
					If len(p1) > 0 Then 
					
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '" + p1 + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '" + p1 + "%' "
						End if
						sql_group = " WHEN field_1 like '" +  p1 + "%' THEN " & CStr(order_count)
						order_count = order_count + 1
						
						Set oRS_similares = Server.CreateObject("ADODB.Recordset")
						oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p1 & "' or REPLACE(replacement,' ','') = '" & trim(p1) & "' ;", MM_overseaspr_STRING
						
						If Not oRS_similares.EOF Then
							Do While Not oRS_similares.EOF
								If len(sql_parts) = 0 Then
									sql_parts = sql_parts + " field_1 like '%" + trim(oRS_similares.Fields.Item("partid").value) + "%' "
								Else
									sql_parts = sql_parts + " OR field_1 like '%" + trim(oRS_similares.Fields.Item("partid").value) + "%' "
								End if
								sql_group = sql_group + " WHEN field_1 like '%" +  trim(oRS_similares.Fields.Item("partid").value) + "%' THEN " & CStr(order_count)
								order_count = order_count + 1
								oRS_similares.MoveNext
							Loop
						
						Else
							order_count = order_count + 1
						End if
						oRS_similares.Close                            
					End if
					
					If len(p2) > 0 Then 
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '%" + p2 + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '%" + p2 + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '%" +  p2 + "%' THEN " & CStr(order_count)
						
						Set oRS_similares = Server.CreateObject("ADODB.Recordset")
						oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p2 & "' or REPLACE(replacement,' ','') = '" & trim(p2) & "' ;", MM_overseaspr_STRING
						
						If Not oRS_similares.EOF Then
							Do While Not oRS_similares.EOF
								If len(sql_parts) = 0 Then
									sql_parts = sql_parts + " field_1 like '%" + oRS_similares.Fields.Item("partid") + "%' "
								Else
									sql_parts = sql_parts + " OR field_1 like '%" + oRS_similares.Fields.Item("partid") + "%' "
								End if
								sql_group = sql_group + " WHEN field_1 like '%" +  oRS_similares.Fields.Item("partid") + "%' THEN " & CStr(order_count)
								order_count = order_count + 1
								oRS_similares.MoveNext
							Loop
						End if
						oRS_similares.Close			
					End if
					
					If len(p3) > 0 Then 
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '%" + p3 + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '%" + p3 + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '%" +  p3 + "%' THEN " & CStr(order_count)
						
						Set oRS_similares = Server.CreateObject("ADODB.Recordset")
						oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p3 & "' or REPLACE(replacement,' ','') = '" & trim(p3) & "' ;", MM_overseaspr_STRING
						
						If Not oRS_similares.EOF Then
							Do While Not oRS_similares.EOF
								If len(sql_parts) = 0 Then
									sql_parts = sql_parts + " field_1 like '%" + oRS_similares.Fields.Item("partid") + "%' "
								Else
									sql_parts = sql_parts + " OR field_1 like '%" + oRS_similares.Fields.Item("partid") + "%' "
								End if
								sql_group = sql_group + " WHEN field_1 like '%" +  oRS_similares.Fields.Item("partid") + "%' THEN " & CStr(order_count)
								order_count = order_count + 1
								oRS_similares.MoveNext
							Loop
						End if
						oRS_similares.Close			
					End if
					
					If len(p4) > 0 Then 
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '%" + p4 + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '%" + p4 + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '%" +  p4 + "%' THEN " & CStr(order_count)
						
						Set oRS_similares = Server.CreateObject("ADODB.Recordset")
						oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p4 & "' or REPLACE(replacement,' ','') = '" & trim(p4) & "' ;", MM_overseaspr_STRING
						
						If Not oRS_similares.EOF Then
							Do While Not oRS_similares.EOF
								If len(sql_parts) = 0 Then
									sql_parts = sql_parts + " field_1 like '%" + oRS_similares.Fields.Item("partid") + "%' "
								Else
									sql_parts = sql_parts + " OR field_1 like '%" + oRS_similares.Fields.Item("partid") + "%' "
								End if
								sql_group = sql_group + " WHEN field_1 like '%" +  oRS_similares.Fields.Item("partid") + "%' THEN " & CStr(order_count)
								
								order_count = order_count + 1
								oRS_similares.MoveNext
							Loop
						End if
						oRS_similares.Close
					End if
					
					If len(p5) > 0 Then 
						If len(sql_parts) = 0 Then
						sql_parts = sql_parts + " field_1 like '%" + p5 + "%' "
						Else
						sql_parts = sql_parts + " OR field_1 like '%" + p5 + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '%" +  p5 + "%' THEN " & CStr(order_count)
						
						Set oRS_similares = Server.CreateObject("ADODB.Recordset")				
						oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p5 & "' or REPLACE(replacement,' ','') = '" & trim(p5) & "' ;", MM_overseaspr_STRING
						
						If Not oRS_similares.EOF Then
							Do While Not oRS_similares.EOF
								If len(sql_parts) = 0 Then
									sql_parts = sql_parts + " field_1 like '%" + oRS_similares.Fields.Item("partid") + "%' "
								Else
									sql_parts = sql_parts + " OR field_1 like '%" + oRS_similares.Fields.Item("partid") + "%' "
								End if
								sql_group = sql_group + " WHEN field_1 like '%" +  oRS_similares.Fields.Item("partid") + "%' THEN " & CStr(order_count)
								
								order_count = order_count + 1
								oRS_similares.MoveNext
							Loop
						End if
						oRS_similares.Close			
					End if
					
	
					
							   
					If Request("Page") = "" Then
						If len(sql_group) > 0 Then
							sql_group = " ORDER BY fam_make_item "
							Session("SQLSearch") = sql_select_price + sql_parts + sql_group + " ;"
							' GGG DisplayParts(15)	
						Else
							If special_items = "ON" or liquidation_items = "ON" then
								'response.Write sql_select_price + " ;"				
								Session("SQLSearch") = sql_select_price + " ORDER BY fam_make_item ;"
								' GGG DisplayParts(15)	
							else 
								If len(family_keyword_list) > 0 Then
									sql_family = sql_family + " field_6 = '" + family_keyword_list + "' "
									Session("SQLSearch") = sql_select_price + sql_parts + sql_family + " ORDER BY fam_make_item ;"
									' GGG DisplayParts(15)	
								ELSE
								
									Session("SQLSearch") = ""
									'response.write("<div class='col-xs-7 col-xs-offset-1'><img src='logooic.png' width='568' height='322' alt='TEMP' ></div>")	
								
								End if	
							End if
						End if
					Else
						' GGG DisplayParts(15)
					End if
					
				' MODEL SEARCH
				Else
					
					sql_select_price = sql_select_price + " (" + MultipleModelKeywords("field_2",model) + ") "
					If len(family_keyword_list) > 0 Then
						sql_family = sql_family + " AND field_6 = '" + family_keyword_list + "' "
						Session("SQLSearch") = sql_select_price + sql_family + " ;"
					else
						Session("SQLSearch") = sql_select_price + " ORDER BY fam_make_item ;"
					end if	
					' GGG DisplayParts(15)
                
                End if
                
                
                %>
                <div id="DisplayParts" ><% DisplayParts(15)%></div> 
                
              
                </td>
                </tr>
                </table>
            
            <!--table responsive-->
            </div>
		</div>      
    	<!--class=row-->  
	</div>
    <input id="Page" type="hidden" value="<%= request("Page") %>" >
	</form>            
  	<script type="text/javascript" src="../bootstrap-3.3.6-dist/js/bootstrap.min.js"></script>
    <!--div container-->
</div>

</body>
</html>
<%
 
dd_category.Close()
Set dd_category = Nothing



%>