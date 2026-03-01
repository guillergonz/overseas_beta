<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" --> 

<%

Response.Expires = -1 

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


<!DOCTYPE HTML>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">

<title>Overseas Import Corporation - Order Detail</title>

<link rel="icon" href="overseaspr/images/favicon.ico" type="image/x-icon" /> 
<link rel="shortcut icon" href="overseaspr/images/favicon.ico" type="image/x-icon" />


<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
order_detail_numRows = order_detail_numRows + Repeat1__numRows
%>
<%
Dim order_detail__MMColParam
order_detail__MMColParam = "1"
If (Request.QueryString("o") <> "") Then 
  order_detail__MMColParam = Request.QueryString("o")
End If
%>

<!--Verificar usuario es dueño de la orden ... -->
<%
Dim count_records

Set mycmd = Server.CreateObject ("ADODB.Command")
mycmd.ActiveConnection = MM_overseaspr_STRING
mycmd.CommandText = "SELECT count(order_number) as count_rec FROM dbo.clients_orders WHERE order_user = '" + Session("MM_UserName") + "' and order_number = '" + order_detail__MMColParam + "' ;" 
mycmd.Prepared = true
Set RS_mycmd = mycmd.Execute

if RS_mycmd.EOF then
	count_records = 0
else
	count_records = (RS_mycmd.Fields.Item("count_rec").Value)
end if

RS_mycmd.Close
Set RS_mycmd = Nothing

%>

<%
Dim order_detail
Dim order_detail_cmd
Dim order_detail_numRows

Set order_detail_cmd = Server.CreateObject ("ADODB.Command")
order_detail_cmd.ActiveConnection = MM_overseaspr_STRING

	order_detail_cmd.CommandText = "SELECT order_number, order_part, order_user, order_qty, order_status, order_date, order_type, item_price, deliv_type, comments FROM dbo.clients_orders WHERE order_number = ? ORDER BY order_id ASC" 

	order_detail_cmd.Prepared = true
	order_detail_cmd.Parameters.Append order_detail_cmd.CreateParameter("param1", 200, 1, 20, order_detail__MMColParam) ' adVarChar

Set order_detail = order_detail_cmd.Execute
order_detail_numRows = 0
%>

<link href="overseas.css" rel="stylesheet" type="text/css" />
<style type="text/css">
body,td,th {
	font-family: "Segoe UI", "Segoe UI Light", "Segoe UI Semibold";
	font-size: 10px;
}
</style>

	<link rel="stylesheet" href="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/themes/base/jquery.ui.all.css">
	<script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/jquery-1.6.2.js"></script>
	<script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/ui/jquery.ui.core.js"></script>
	<script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/ui/jquery.ui.widget.js"></script>
	<script src="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/ui/jquery.ui.button.js"></script>
	<link rel="stylesheet" href="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/demos/demos.css">

<style type="text/css">
	h1 {
	font-size: 18px;
}
h2 {
	font-size: 16px;
}
h3 {
	font-size: 14px;
}
h4 {
	font-size: 12px;
}
h5 {
	font-size: 10px;
}
h6 {
	font-size: 9px;
}
a  {
	font-size:9px;
}
body {
	background-repeat: repeat;
	margin: 0px;
}
.ui-button {
	display: inline-block;
	position: relative;
	padding: .1em;
	/* [disabled]margin-right: 0.1em; */
	text-decoration: none !important;
	cursor: pointer;
	text-align: center;
	zoom: 1;
	overflow: visible;
} /* the overflow property removes extra width in IE */

.ui-button-icon-only { width: 2.2em; }  to make room for the icon, a width needs to be set here 
button.ui-button-icon-only { width: 2.4em; }  button elements seem to need a little more width 
.ui-button-icons-only { width: 3.4em; } 
button.ui-button-icons-only { width: 3.7em; } 


/*button text element */
.ui-button .ui-button-text {
	display: block;
	line-height: 1.4;
	font-size: 11px;
	font-family: "Segoe UI", "Segoe UI Light", "Segoe UI Semibold", "Lucida Console";
}
.ui-button-text-only .ui-button-text { padding: 0; margin:0; }
.ui-button-icon-only .ui-button-text, .ui-button-icons-only .ui-button-text { padding: .0em; text-indent: -9999999px; }
.ui-button-text-icon-primary .ui-button-text, .ui-button-text-icons .ui-button-text { padding: 0, .3em, 0, .3em ; }
.ui-button-text-icon-secondary .ui-button-text, .ui-button-text-icons .ui-button-text { padding: 0, .3em, 0, .3em ; }
.ui-button-text-icons .ui-button-text { padding: 0, .3em, 0, .3em ; }
/* no icon support for input elements, provide padding by default */
input.ui-button { padding: 0em, .3em, 0, .3em ; }

/*button icon element(s) */
.ui-button-icon-only .ui-icon, .ui-button-text-icon-primary .ui-icon, .ui-button-text-icon-secondary .ui-icon, .ui-button-text-icons .ui-icon, .ui-button-icons-only .ui-icon { position: absolute; top: 50%; margin-top: -8px; }
.ui-button-icon-only .ui-icon { left: 50%; margin-left: -8px; }
.ui-button-text-icon-primary .ui-button-icon-primary, .ui-button-text-icons .ui-button-icon-primary, .ui-button-icons-only .ui-button-icon-primary { left: .5em; }
.ui-button-text-icon-secondary .ui-button-icon-secondary, .ui-button-text-icons .ui-button-icon-secondary, .ui-button-icons-only .ui-button-icon-secondary { right: .5em; }
.ui-button-text-icons .ui-button-icon-secondary, .ui-button-icons-only .ui-button-icon-secondary { right: .5em; }

/*button sets*/
.ui-buttonset { margin-right: 7px; }
.ui-buttonset .ui-button { margin-left: 0; margin-right: -.3em; }

/* workarounds */
button.ui-button::-moz-focus-inner { border: 0; padding: 0; } /* reset extra padding in Firefox */

</style>


<script>
	$(function() {
		$( "input:submit, a, button, #Printerfriendly" ).button();
		//$( "a, #Printerfriendly" ).click(function() { return false; });
	});
</script>




</head>

<body>


<%
Dim dd_category
Dim dd_category_cmd
Dim dd_category_numRows

Set dd_category_cmd = Server.CreateObject ("ADODB.Command")
dd_category_cmd.ActiveConnection = MM_overseaspr_STRING
' if English use column english_desc ggg
'LRO - done
If Session("lang") = "S" Then
	dd_category_cmd.CommandText = "SELECT field_1, field_2, english_desc FROM dbo.familicat ORDER BY field_1 ASC" 
Else
	dd_category_cmd.CommandText = "SELECT field_1, field_2, english_desc FROM dbo.familicat ORDER BY english_desc ASC" 
End if

dd_category_cmd.Prepared = true

Set dd_category = dd_category_cmd.Execute
dd_category_numRows = 0

%>

	<div style="margin:4px">
	<a href="/part_search.asp" target="_self"><img src="images/oiclogo2.gif" width="256" height="31" border="0" align="absmiddle" /></a>
    </div>
	


    
<table width="1158px" border="0" align="center" cellpadding="0" cellspacing="0" class="borderaround_white">
    <tr>
      <td height="1050px" align="center" valign="top"><br />
        <table align="center" width="1189px" height="1040" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td colspan="2"><table width="1130px" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="4%" height="18">&nbsp;</td>
                <td width="16%"><%= GetUserName(Session("MM_Username"))%> 
                            </td>
                <td width="13%">
               
                 </td>
                <td width="13%"> <p>&nbsp;FECHA&nbsp;:&nbsp;&nbsp;<%=Date()%></p>  
                </td>
                <td width="16%"><a id="completar" name="completar" href="cart.asp" style="width:100px" ><%= Lang("completar_orden") %></a></td>
                
                <td width="15%" align="center" >
                 <a id="buscar" name="buscar" href="part_search.asp" style="width:100px"><%= Lang("buscar") %></a>
                
                </td>
                <td width="23%" align="center" >&nbsp;&nbsp;&nbsp;&nbsp;<a id="logout" href="<%= MM_Logout %>" style="width:100px">&nbsp;<%= Lang("salir") %>&nbsp;</a></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td valign="top"> </td>
            <td align="left" valign="top">&nbsp;</td>
          </tr>
          <tr>
            <td width="180" height="2614" valign="top">
              <% If Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/cart.asp" or Request.ServerVariables("SCRIPT_NAME") = "/cart.asp" or len(Request("o")) > 0 Then %>              <br>
              <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablas">
                <tr>
                <td align="center" ><div class="grayback" ><%= Lang("preparado") %></div>
                </td>
                </tr>
                
                <tr>
                <td align="center"> <div id="shopping_cart"  >
                  <% CartDisplay("N") %>
                </div>
                </td>
                </tr>
                
                </table>
              
              <% End if %>
              <br>
              <% If Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/cart.asp" or Request.ServerVariables("SCRIPT_NAME") = "/cart.asp" or len(Request("o")) > 0 Then %>
              
              <table width = "100%" border="0" cellpadding="0" cellspacing="0" class="left_menu" >
                <tr>
                  <td  align="center" height='10'><div class="grayback"><%= Lang("ordenes") %></div></td>
                </tr>
                <tr>
                  <td>
                  <div id="orders">
                    <% OrdersDisplay() %>
                  </div>
                  
                  <div class="display_no" id="loader"><img src="images/ajax-loader.gif" width="16" height="16"></div></td>
              </table>
              
              <br>
              <hr>
              <br>
              
              <table width="200px" border="0" cellpadding="0" class="left_menu" cellspacing="0"  >
                <tr>
                  <td align="center" width="100%"><div class="grayback"><%= Lang("ACCOUNTSTATEMENT") %></div></td>
                </tr>
                <tr>
                  <td align="center" width="100%">
                  <a style="width:98%" href="account_statement_iframe.asp" title="Estado de Cuenta" target="_self"><%= Lang("estado_de_cuenta_actual") %></a>
                  </td>
                </tr>
                 <tr><td>&nbsp;</td></tr>
                <tr>
                 <td align="center" >
      <div class="grayback" style="width:100%">OVERSEAS IMPORT</div>
      </td></tr>
                <tr>
                  <td align="center">
                  <a style="width:98%" href="http://www1.overseaspr.com/soportes/index.html" title="Catálogo de Soportes" target="_self" ><%= Lang("catalogo") %></a>
                  <br></td>
                </tr>
                
              </table>
              <br>
              
              
              <% End if %>
              
              
            <td width="1006" align="left" valign="top">
              <form id="search_part" action="part_search.asp" method="post">
                <script>
				function searchSel() {
				  var input=document.getElementById('fks').value.toUpperCase();
				  var output=document.getElementById('fcs').options;

				  for(var i=0;i<output.length;i++) {
					if(output[i].value.indexOf(input)==0){
					  output[i].selected=true;
					  }
					if(document.forms[0].fks.value==''){
					  output[0].selected=true;
					  }
				  }
				}
				</script>
                <table width="1015" height="2153" border="0" cellpadding="0" cellspacing="0" class="tablasFORMATO" >
                              <tr>
              <td height="46" colspan="2" align="center" valign="top">
                <table width="966" border="0" cellpadding="0" cellspacing="0" style="font-size:9px" >
                  <tr height="20" style="font-size:9px">
                  	<td height="46">
                  	  &nbsp;<%= Lang("family_keyword") %>&nbsp;
                  	  <input type="text" class="ui-widget-content" id="fks"  onKeyUp="searchSel()" size="16"></td>
                    <td width="240" align="center">&nbsp;<%= Lang("family_category") %>&nbsp;<select name="fcs" class="ui-widget-content" id="fcs" >
                      <option selected value=""> </option> 
                      <%While (NOT dd_category.EOF) %>
                      <%If Session("lang") = "S" Then %>
                      <option value="<%=(dd_category.Fields.Item("field_1").Value)%>"><%=(dd_category.Fields.Item("field_1").Value)%></option>
                      <% Else %>
                      <option value="<%=(dd_category.Fields.Item("english_desc").Value)%>"><%=(dd_category.Fields.Item("english_desc").Value)%></option>
                      <% End if %>
                      <% dd_category.MoveNext()
					Wend
					If (dd_category.CursorType > 0) Then
					  dd_category.MoveFirst
					Else
					  dd_category.Requery
					End If %>
                    </select></td>
                    
                    <td width="165" align="center">
                      &nbsp;<%= Lang("modelo") %>&nbsp;<input name="model" type="text" class="ui-widget-content" id="fs2" size="16"></td>
                    
                    <td width="90" align="center">
                      <p>&nbsp;</p></td>
                    
                    <td align="center" valign="middle">
                      
                      
                     <input style="font-size:12px; font-weight:bold;width:140px" class="ui-widget" name="submit" type="submit" id="submit" value="<%=Lang("buscar")%>">                    </td>
                    </tr>
                  </table>
                	
                </td>
              </tr>

                    
                  <tr>
                    <td width="89" height="1735"  align="center" valign="top"><table  border="0" cellpadding="0" cellspacing="1">
                        <tr>
                          <td colspan="2" align="center"><%= Lang("entre_10_piezas") %></td>
                        </tr>
                        <tr>
                          <td width="21" height="25" align="left" valign="middle">&nbsp;1</td>
                          <td width="70" align="left" valign="middle"><input name="p1" type="text" class="search_box" id="p1" size="9" maxlength="15" /></td>
                        </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;2</td>
                          <td height="25" align="left" valign="middle"><input name="p2" type="text" class="search_box" id="p2" size="9" maxlength="15" /></td>
                        </tr>
                        <tr>
                          <td height="25" align="left">&nbsp;3</td>
                          <td height="25" align="left"><input name="p3" type="text" class="search_box" id="p3" size="9" maxlength="15" /></td>
                        </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;4</td>
                          <td height="25" align="left" valign="middle"><input name="p4" type="text" class="search_box" id="p4" size="9" maxlength="15" /></td>
                        </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;5</td>
                          <td height="25" align="left" valign="middle"><input name="p5" type="text" class="search_box" id="p5" size="9" maxlength="15" /></td>
                        </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;6</td>
                          <td height="25" align="left" valign="middle"><input name="p6" type="text" class="search_box" id="p6" size="9" maxlength="15" /></td>
                        </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;7                      </td>
                          <td height="25" align="left" valign="middle"><input name="p7" type="text" class="search_box" id="p7" size="9" maxlength="15" /></td>
                        </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;8                      </td>
                          <td height="25" align="left" valign="middle"><input name="p8" type="text" class="search_box" id="p8" size="9" maxlength="15" /></td>
                        </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;9</td>
                          <td height="25" align="left" valign="middle"><input name="p9" type="text" class="search_box" id="p9" size="9" maxlength="15" /></td>
                        </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">10                        </td>
                          <td height="25" align="left" valign="middle"><input name="p10" type="text" class="search_box" id="p10" size="9" maxlength="15" /></td>
                        </tr>
                    </table>
                    <br><br><br>
                    
                    </td>
                    <td width="915" align="center" valign="top"><div class="display_yes" style="margin:10px;width:98%;" >
                       
                            
                            <% if count_records > 0 OR Session("MM_UserName") = "Z099" then %>
                            
                           
                        <table  border="0" align = "left" cellpadding="0" cellspacing="0" class="ui-widget-content"  >
                          <tr height="20px">
                            <td width="886" align="center"><table width="720px" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td align="center">
                                  <p><input style="width:98%" name="Printerfriendly" type="button" id="Printerfriendly" value="PRINT / IMPRIMIR " onClick="window.location='printfriendly_salesorder.asp?o=<%=( trim(order_detail.Fields.Item("order_number").Value))%>'" /></p>
                                  
                                </td>
                                
                              </tr>
                              </table>
                          <br>    
                        <table align = "left" border="0" cellspacing="0" cellpadding="0"  >

                              <tr>
                                <td width="780" align="center">&nbsp;
                                  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" class="tablas" style="font-size:10px" >
                                    <tr>
                                      <td width = "19%" align="right"><%= Lang("detalles_orden") %>&nbsp;</td>
                                      <td width = "39%" align="left"><%=(order_detail.Fields.Item("order_number").Value)%>&nbsp;<%= (order_detail.Fields.Item("order_user").Value) %></td>
                                      <td width="18%" align="right"><strong><%= Lang("fecha_y_hora") %>&nbsp;</strong></td>
                                      <td width="24%" colspan="2" align="left"><%= trim(order_detail.Fields.Item("order_date").Value)%></td>
                                      
                                    </tr>
                                    <tr>
                                    <td colspan="5" align="center"></td>
                                    </tr>
                                    <tr>
                                      <td align="right"><strong>&nbsp;<%= Lang("estatus") %>&nbsp;</strong></td>
                                      <td align="left"><label for="order_status"></label>
                                        <select name="order_status" disabled="disabled" class="display_yes" id="order_status">
                                          <option value="O" <%If (Not isNull((order_detail.Fields.Item("order_status").Value))) Then If ("O" = CStr((order_detail.Fields.Item("order_status").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("status_pending") %></option>
                                          <option value="P" <%If (Not isNull((order_detail.Fields.Item("order_status").Value))) Then If ("P" = CStr((order_detail.Fields.Item("order_status").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("status_processed") %></option>
                                      </select></td>
                                      <td align="right"><strong><%= Lang("metodo_envio") %></strong>&nbsp;</td>
                                      <td colspan="2" align="left">
                                        <select name="delivery_type" class="display_yes" id="delivery_type" disabled="disabled">
                                          <option value="D" <%If (Not isNull((order_detail.Fields.Item("deliv_type").Value))) Then If ("D" = CStr((order_detail.Fields.Item("deliv_type").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("entrega") %></option>
                                          <option value="P" <%If (Not isNull((order_detail.Fields.Item("deliv_type").Value))) Then If ("P" = CStr((order_detail.Fields.Item("deliv_type").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("recoger") %></option>
                                      </select></td>
                                    </tr>
								  </table>
                                    <br>
                                    <table width="700" align = "left" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td colspan="5" align="center"></td>
                                      </tr>
                                    <tr>
                                      <td width="20%" align="center">&nbsp;</td>
                                      <td width="30%" align="center">&nbsp;</td>
                                      <td width="10%" align="center">&nbsp;</td>                                
                                      <td width="20%" align="center">&nbsp;</td>
                                      <td width="20%" align="center">&nbsp;</td>
                                      </tr>
                                    <tr>
                                      <td align="center"><%= Lang("num_pieza") %></td>
                                      <td align="left"><%= Lang("descripcion") %></td>
                                      <td align="center"><%= Lang("ordenado") %></td>
                                      <td align="center"><%= Lang("precio") %></td>
                                      <td align="center">SUB-TOTAL</td>
                                      </tr>
                                    <tr>
                                      <td align="center">&nbsp;</td>
                                      <td align="center">&nbsp;</td>
                                      <td align="center">&nbsp;</td>
                                      <td align="center">&nbsp;</td>
                                      <td align="center">&nbsp;</td>
                                      </tr>
                                   
                                   
                                        <%
	color1 = "#CCCCCC"
	color2 = "#FFFFFF"
	current_color = color1
	rc = 1
	total = 0
	ldescription = ""
%>
                                        <% 
While ((Repeat1__numRows <> 0) AND (NOT order_detail.EOF)) 

	' Record set to get description in english or spanish for part item ...
	Set oRS2 = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.partmst1_distinct.field_2, dbo.partmst1_distinct.english_version FROM dbo.partmst1_distinct WHERE dbo.partmst1_distinct.field_1 = '" + CStr(order_detail.Fields.Item("order_part").Value)  + "' ;"
	oRS2.Open strSQL, MM_overseaspr_STRING
	
	// New code for description/instructions

	if trim(order_detail.Fields.Item("order_part").Value) = "ZZZNOF" then
	
		ldescription = trim(order_detail.Fields.Item("comments").Value)
	
	else
		
		ldescription = ""
		If Session("lang") = "E" Then
			if IsNull(oRS2.Fields.Item("english_version")) then
				ldescription = trim(oRS2.Fields.Item("field_2"))
			else	
				if trim(oRS2.Fields.Item("english_version")) > "" then
					ldescription = CStr(oRS2.Fields.Item("english_version"))
				else
					ldescription = CStr(oRS2.Fields.Item("field_2"))
				end if		
			end if	
		else
			ldescription = CStr(oRS2.Fields.Item("field_2"))
		End if
		
	End if
		
	'  close oRS2 ...
	oRS2.Close
	Set oRS2 = Nothing	
	'  mod ends here .

%>
                                     
                                    <tr>
                                      <td height="20" align="center" bgcolor='<%=  current_color %>'><%=(order_detail.Fields.Item("order_part").Value)%></td>
                                      <td align="left" bgcolor='<%=  current_color %>'><%= ldescription %></td>
                                      <td align="center" bgcolor='<%=  current_color %>'><%=(order_detail.Fields.Item("order_qty").Value)%></td>
                                      <td align="center" bgcolor='<%=  current_color %>'><%=(CurrencyConvert(order_detail.Fields.Item("item_price").Value))%></td>
                                      <td align="center" bgcolor='<%=  current_color %>'><%= CurrencyConvert(order_detail.Fields.Item("order_qty").Value * order_detail.Fields.Item("item_price").Value) %></td>
                                      </tr>
                                    <% 		
If current_color = color2 then current_color = color1 else current_color = color2
rc = rc + 1
total = total + (order_detail.Fields.Item("item_price").Value * order_detail.Fields.Item("order_qty").Value)
%>
                                    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  order_detail.MoveNext()
Wend
%>
                                    <tr>
                                      <td align="center">&nbsp;</td>
                                      <td align="center"><%= CurrencyConvert(total) %></td>
                                      <td align="center">TAX</td>
                                      <td align="center" class="ui-state-highlight">
                                      
                                      <% 
									  If Session("MM_CityTax")="Y" Then 
									   	vCityTax  = round( (total * .01) ,2)
	 									'vStateTax = round( (total * .06) ,2)
										vStateTax = round( (total * .105) ,2)
										response.write(  CurrencyConvert(total) + " / " + CurrencyConvert(vCityTax) + " / " + CurrencyConvert(vStateTax))
                                      End IF
                                      %>
                                      
                                      </td>
                                      </tr>
                                    <tr>
                                      <td align="center">&nbsp;</td>
                                      <td align="center">&nbsp;</td>
                                      <td align="center" >TOTAL</td>
                                      <td align="center" class="ui-state-highlight">
									  
									  <% 
									  If Session("MM_CityTax")="Y" Then 
										  response.write( CurrencyConvert(total + vCityTax + vStateTax)) 
                                      Else
   										  response.write( CurrencyConvert(total))  
                                      End IF
                                      %>
                                      
                                      </td>
                                      <td align="left" class="table_header">&nbsp;</td>
                                      </tr>
                                    <tr>
                                      <td align="center">&nbsp;</td>
                                      <td align="center">&nbsp;</td>
                                      <td align="center" class="table_header">&nbsp;</td>
                                      <td align="center" class="table_header">&nbsp;</td>
                                      </tr>
                                      <tr><td colspan="5">&nbsp;</td></tr>

                                      <tr><td colspan="5">&nbsp;</td></tr>

                                      <tr><td colspan="5">&nbsp;</td></tr>

                                    </table>
                                  
                                  
                                  <% end if %>
                                  
                                  <br><br><br>
                                  
                                </td>
                              </tr>
                            </table>
                                                        <br><br><br>
                            </td>
                          </tr>
                        </table>
                                                    					
                      </div>
                     
                    
                      </td>
                  </tr>
                  <tr>
                    <td  colspan="2"  align="center" valign="top">&nbsp;</td>
                  </tr>
               	
                </table>
			
              </form>
				<br>
                
            </td>
          </tr>
          

          <tr><td height="359">&nbsp;</td></tr>
          
          
        </table>
        <br><br>
        
      </td>
      </tr>
      
  </table>


</body>
</html>
<%
dd_category.Close()
Set dd_category = Nothing
%>
<%
order_detail.Close()
Set order_detail = Nothing
%>
