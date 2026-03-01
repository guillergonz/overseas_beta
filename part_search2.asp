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

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Search ...</title>

<link href="overseas.css" rel="stylesheet" type="text/css" />
<style type="text/css">
body,td,th {
	font-size: 9px;
}
body {
	background-image: url(/images/back_pagina.gif);
	margin-left: 20px;
	margin-top: 20px;
	margin-right: 20px;
	margin-bottom: 20px;
	background-repeat: repeat;
}
</style>
</head>

<body>

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

<script>
var xmlHttp

function AddToCart(str,amount)
{ 
xmlHttp=GetXmlHttpObject()
if (xmlHttp==null)
{
alert ("Browser does not support HTTP Request")
return
} 
//$("#txtHint").className = "display_no" ;
document.getElementById("shopping_cart").className = "display_no" ;
document.getElementById("loader").className = "display_yes" ;

var url="add_to_shopping_cart.asp"
url=url+"?p=" + str + "&a=" + amount
url=url+"&sid="+Math.random()
xmlHttp.onreadystatechange=stateChanged 
xmlHttp.open("GET",url,true)
xmlHttp.send(null)
}

function stateChanged() 
{ 
if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
{ 
document.getElementById("shopping_cart").innerHTML=xmlHttp.responseText 
document.getElementById("shopping_cart").className = "display_yes" ;
document.getElementById("loader").className = "display_no" ;
} 
} 

function GetXmlHttpObject()
{ 
var objXMLHttp=null
if (window.XMLHttpRequest)
{
objXMLHttp=new XMLHttpRequest()
}
else if (window.ActiveXObject)
{
objXMLHttp=new ActiveXObject("Microsoft.XMLHTTP")
}
return objXMLHttp
}
</script>

<script>
function ValidateCart(part,amt) {
	amount = document.getElementById(amt).value
	AddToCart(part,amount);
}
</script>


<script>
var xmlHttp2

function DelFromCart(str)
{ 
xmlHttp2=GetXmlHttpObject()
if (xmlHttp2==null)
{
alert ("Browser does not support HTTP Request")
return
} 
//$("#txtHint").className = "display_no" ;
document.getElementById("shopping_cart").className = "display_no" ;
document.getElementById("loader").className = "display_yes" ;

var url2="del_p_shopping_cart.asp"
url2=url2+"?p=" + str
url2=url2+"&sid="+Math.random()
xmlHttp2.onreadystatechange=stateChanged2 
xmlHttp2.open("GET",url2,true)
xmlHttp2.send(null)
}

function stateChanged2() 
{ 
if (xmlHttp2.readyState==4 || xmlHttp2.readyState=="complete")
{ 
document.getElementById("shopping_cart").innerHTML=xmlHttp2.responseText 
document.getElementById("shopping_cart").className = "display_yes" ;
document.getElementById("loader").className = "display_no" ;
} 
} 

function GetXmlHttpObject2()
{ 
var objXMLHttp2=null
if (window.XMLHttpRequest)
{
objXMLHttp2=new XMLHttpRequest()
}
else if (window.ActiveXObject)
{
objXMLHttp2=new ActiveXObject("Microsoft.XMLHTTP")
}
return objXMLHttp2
}
</script>


  <table width="1000px" height="480px" border="0" align="center" cellpadding="0" cellspacing="0" class="borderaround">
    <tr>
      <td width="1052" height="450" align="center" valign="top"><br />
        <table align="center" width="989" height="433" border="0" cellpadding="0" cellspacing="1">
          <tr>
            <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#FFFFFF">
                <td width="0%" height="48">&nbsp;</td>
                <td width="26%"><img src="images/oiclogo2.gif" width="256" height="31" align="absmiddle"></td>
                <td width="9%"> 
                  <% if Session("MM_UserName") = "Z099" then %><%= GetCurrentSales(Session("MM_UserName")) %><% End if %>
                </td>
                <td width="24%">
                  </div><div id="linea"><%= GetUserName(Session("MM_Username"))%><br> UID&nbsp;:&nbsp;&nbsp;<%= Session("MM_Username")%>&nbsp;&nbsp;&nbsp;<%= Lang("fecha") %>&nbsp;:&nbsp;&nbsp;<%=Date()%></div></td>
                <td width="41%" align="center" class="tablas_font12"><a href="cart.asp"><%= Lang("completar_orden") %></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%= MM_Logout %>"><%= Lang("salir") %></a>
                  </div>
                </td>
              </tr>
             </table></td>
          </tr>
          <tr>
            <td height="15" valign="top"> </td>
            <td align="left" valign="top">&nbsp;</td>
          </tr>
          <tr>
            <td width="82" valign="top">
              <% If Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/part_search.asp" or Request.ServerVariables("SCRIPT_NAME") = "/part_search.asp" Then %>
              <table width="200px" border="0" cellpadding="0" cellspacing="0" >
                <tr><td align="center" ><div class="left_menu"><%= Lang("preparado") %></div></td></tr>
                <tr><td align="center">
                  <div class="display_yes" id="shopping_cart"><% CartDisplay("N") %></div>
              <div class="display_no" id="loader"><img src="images/ajax-loader.gif" width="20" height="20"></div></td></tr></table><% End if %>
              

			  
			  <% If Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/cart.asp" or Request.ServerVariables("SCRIPT_NAME") = "/cart.asp" or len(Request("o")) > 0 Then %>
              <table width = "200px" border="0" cellpadding="0" cellspacing="0" class="left_menu" >
                <tr class="left_menu">
                  <td width = "200px" align="center" height='10'><%= Lang("ordenes") %></td>
                </tr>
                <tr>
                  <td><div class="display_yes" width="200px" id="orders">
                    <% OrdersDisplay() %>
                    </div>
                    <div class="display_no" id="loader2"><img src="images/ajax-loader.gif" width="16" height="16"></div></td>
              </table>
              <% End if %>
              <table width="200px" border="0" cellpadding="0" cellspacing="0" >
                

<tr><td align="center" height='20' class="left_menu"><%= Lang("estado_de_cuenta_actual") %></td></tr>
<tr><td align="center">&nbsp;</td></tr>
<tr><td height='20' align="center"><div id="orders"><a href="account_statement_iframe.asp" title="Estado de Cuenta" target="_self"><%= Lang("estado_de_cuenta_actual") %></a><br><br></div></td></tr>
<tr><td align="center" height='20' class="left_menu">OVERSEAS IMPORT</td></tr>
<tr><td align="center">&nbsp;</td></tr>
<tr><td align="center"><div id="orders"><a href="http://69.89.36.77/soportes/index.html" title="Catálogo de Soportes" target="_self"><%= Lang("catalogo") %></a></div>
<br><br></td></tr></table>
              
<td width="900" align="left" valign="top">
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
                <table width="884" border="0" cellpadding="0" cellspacing="0" class="borderaround_white">
  <tr>
                    <td colspan="2" align="center" valign="top">                      <table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td colspan="4">&nbsp;</td></tr>
<tr><td colspan="4">&nbsp;</td></tr>
                        <tr height="25px">
                          <td width="21%" height="31" align="center">
  <%= Lang("family_keyword") %>                        <input type="text" class="search_box_NORADIUS" id="fks" onKeyUp="searchSel()"></td>
                          <td width="26%" align="center">
<%= Lang("family_category") %>                          <select name="fcs" class="search_box_NORADIUS" id="fcs">
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
                          <td width="21%" align="center">
                          <%= Lang("modelo") %><input name="model" type="text" class="search_box_NORADIUS" id="fs2"></td>
<td width="32%" align="left" height="31"><h5><input type="checkbox" name="special_items" id="special_items">&nbsp;&nbsp;<%= Lang("Specials") %>&nbsp;&nbsp;<input name="submit" type="submit" id="submit" value="<%= Lang("buscar") %>">
</h5></td></tr></table></td>
</tr><tr>
<td width="153" align="center" valign="top"><table width=150px" border="0" cellpadding="0" cellspacing="1">
<tr><td colspan="2" align="center"><%= Lang("entre_10_piezas") %></td>
</tr>
<tr><td height="25" align="left" valign="middle">&nbsp;1</td><td align="left" valign="middle"><input name="p1" type="text" class="search_box" id="p1" size="15" /></td>
</tr><tr><td height="25" align="left" valign="middle">&nbsp;2</td><td height="25" align="left" valign="middle"><input name="p2" type="text" class="search_box" id="p2" size="15" /></td>
                      </tr>
                        <tr>
                          <td height="25" align="left">&nbsp;3</td>
                          <td height="25" align="left"><input name="p3" type="text" class="search_box" id="p3" size="15" /></td>
                      </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;4</td>
                          <td height="25" align="left" valign="middle"><input name="p4" type="text" class="search_box" id="p4" size="15" /></td>
                      </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;5</td>
                          <td height="25" align="left" valign="middle"><input name="p5" type="text" class="search_box" id="p5" size="15" /></td>
                      </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;6</td>
                          <td height="25" align="left" valign="middle"><input name="p6" type="text" class="search_box" id="p6" size="15" /></td>
                      </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;7                      </td>
                          <td height="25" align="left" valign="middle"><input name="p7" type="text" class="search_box" id="p7" size="15" /></td>
                      </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;8                      </td>
                          <td height="25" align="left" valign="middle"><input name="p8" type="text" class="search_box" id="p8" size="15" /></td>
                      </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;9</td>
                          <td height="25" align="left" valign="middle"><input name="p9" type="text" class="search_box" id="p9" size="15" /></td>
                      </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">10                        </td>
                          <td height="25" align="left" valign="middle"><input name="p10" type="text" class="search_box" id="p10" size="15" /></td>
                      </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;</td>
                          <td height="25" align="left" valign="middle">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="25" align="left" valign="middle">&nbsp;</td>
                          <td height="25" align="left" valign="middle">&nbsp;</td>
                        </tr>
                  </table></td>
                    <td width="731" align="center" valign="top">
                      <!-- TemplateBeginEditable name="main_page" -->
                      <table border="0" cellpadding="0" cellspacing="0" >
                        <tr>
                          <td width="731" align="left">

<%

model = Trim(Ucase(Request("model")))
family_keyword_list = Trim(FamilyCode(Trim(Request("fcs"))))
special_items = Trim(Ucase(Request("special_items")))



If special_items = "ON" then

		sql_select_price = "SELECT field_1, field_2, field_4, field_5, familia_descripcion, english_version, 'N/A' as replacement , family_description , fam_make_item FROM dbo.partmst1_distinct WHERE field_1 in ( select dbo.prespecials.specials from dbo.prespecials ) "

else
	
	If Session("user_level") = 3 Then
		
		sql_select_price = "SELECT field_1, field_2, field_3, field_5, familia_descripcion, english_version, similar.replacement, family_description , fam_make_item FROM {oj dbo.partmst1_distinct LEFT OUTER JOIN dbo.similar ON dbo.partmst1_distinct.field_1 = dbo.similar.partid} WHERE "
		
	Else
	
		sql_select_price = "SELECT field_1, field_2, field_4, field_5, familia_descripcion, english_version, similar.replacement , family_description , fam_make_item FROM {oj dbo.partmst1_distinct LEFT OUTER JOIN dbo.similar ON dbo.partmst1_distinct.field_1 = dbo.similar.partid} WHERE "
		
	End if

End if

	If len(model) = 0 or special_items = "ON" Then

		sql_parts 	= ""
		sql_group 	= ""
		sql_family	= ""
		
		p1	= trim(Request("p1"))
		p2	= trim(Request("p2"))
		p3	= trim(Request("p3"))
		p4	= trim(Request("p4"))
		p5	= trim(Request("p5"))
		p6	= trim(Request("p6"))
		p7	= trim(Request("p7"))
		p8	= trim(Request("p8"))
		p9 	= trim(Request("p9"))
		p10	= trim(Request("p10"))
		
	'	model				= Trim(Request("model"))
		
		'If len(family_keyword_list) = 0 Then
		
		Set oRS = Server.CreateObject("ADODB.Recordset")
		Set oRS_similares = Server.CreateObject("ADODB.Recordset")
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
				
				'SIMILAR STEPS
				oRS_similares.Open "SELECT replacement FROM similar WHERE partid = '" & p1 & "' ;", MM_overseaspr_STRING
				
				If Not oRS_similares.EOF Then
					Do While Not oRS_similares.EOF
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '" +  oRS_similares.Fields.Item("replacement") + "%' THEN " & CStr(order_count)
						order_count = order_count + 1
						oRS_similares.MoveNext
					Loop
			'		oRS_similares.Close
				Else
					order_count = order_count + 1
				End if
				oRS_similares.Close
											
			End if
			
			If len(p2) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '" + p2 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '" + p2 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '" +  p2 + "%' THEN " & CStr(order_count)
			
				'SIMILAR STEPS
				oRS_similares.Open "SELECT replacement FROM similar WHERE partid = '" & p2 & "' ;", MM_overseaspr_STRING
				If Not oRS_similares.EOF Then
					Do While Not oRS_similares.EOF
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '" +  oRS_similares.Fields.Item("replacement") + "%' THEN " & CStr(order_count)
						order_count = order_count + 1
						oRS_similares.MoveNext
					Loop
		'			oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p3) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '" + p3 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '" + p3 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '" +  p3 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
				oRS_similares.Open "SELECT replacement FROM similar WHERE partid = '" & p3 & "' ;", MM_overseaspr_STRING
				If Not oRS_similares.EOF Then
					Do While Not oRS_similares.EOF
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '" +  oRS_similares.Fields.Item("replacement") + "%' THEN " & CStr(order_count)
						order_count = order_count + 1
						oRS_similares.MoveNext
					Loop
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p4) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '" + p4 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '" + p4 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '" +  p4 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
				oRS_similares.Open "SELECT replacement FROM similar WHERE partid = '" & p4 & "' ;", MM_overseaspr_STRING
				If Not oRS_similares.EOF Then
					Do While Not oRS_similares.EOF
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '" +  oRS_similares.Fields.Item("replacement") + "%' THEN " & CStr(order_count)
						
						order_count = order_count + 1
						oRS_similares.MoveNext
					Loop
	'				oRS_similares.Close
				End if
				oRS_similares.Close
			End if
			
			If len(p5) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '" + p5 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '" + p5 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '" +  p5 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
				oRS_similares.Open "SELECT replacement FROM similar WHERE partid = '" & p5 & "' ;", MM_overseaspr_STRING
				If Not oRS_similares.EOF Then
					Do While Not oRS_similares.EOF
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '" +  oRS_similares.Fields.Item("replacement") + "%' THEN " & CStr(order_count)
						
						order_count = order_count + 1
						oRS_similares.MoveNext
					Loop
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p6) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '" + p6 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '" + p6 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '" +  p6 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
				oRS_similares.Open "SELECT replacement FROM similar WHERE partid = '" & p6 & "' ;", MM_overseaspr_STRING
				If Not oRS_similares.EOF Then
					Do While Not oRS_similares.EOF
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '" +  oRS_similares.Fields.Item("replacement") + "%' THEN " & CStr(order_count)
						
						order_count = order_count + 1
						oRS_similares.MoveNext
					Loop
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p7) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '" + p7 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '" + p7 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '" +  p7 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
				oRS_similares.Open "SELECT replacement FROM similar WHERE partid = '" & p7 & "' ;", MM_overseaspr_STRING
				If Not oRS_similares.EOF Then
					Do While Not oRS_similares.EOF
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '" +  oRS_similares.Fields.Item("replacement") + "%' THEN " & CStr(order_count)
						
						order_count = order_count + 1
						oRS_similares.MoveNext
					Loop
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p8) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '" + p8 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '" + p8 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '" +  p8 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
				oRS_similares.Open "SELECT replacement FROM similar WHERE partid = '" & p8 & "' ;", MM_overseaspr_STRING
				If Not oRS_similares.EOF Then
					Do While Not oRS_similares.EOF
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '" +  oRS_similares.Fields.Item("replacement") + "%' THEN " & CStr(order_count)
						
						order_count = order_count + 1
						oRS_similares.MoveNext
					Loop
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p9) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '" + p9 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '" + p9 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '" +  p9 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
				oRS_similares.Open "SELECT replacement FROM similar WHERE partid = '" & p9 & "' ;", MM_overseaspr_STRING
				If Not oRS_similares.EOF Then
					Do While Not oRS_similares.EOF
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '" +  oRS_similares.Fields.Item("replacement") + "%' THEN " & CStr(order_count)
						
						order_count = order_count + 1
						oRS_similares.MoveNext
					Loop
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p10) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '" + p10 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '" + p10 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '" +  p10 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
				oRS_similares.Open "SELECT replacement FROM similar WHERE partid = '" & p10 & "' ;", MM_overseaspr_STRING
				If Not oRS_similares.EOF Then
					Do While Not oRS_similares.EOF
						If len(sql_parts) = 0 Then
							sql_parts = sql_parts + " field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						Else
							sql_parts = sql_parts + " OR field_1 like '" + oRS_similares.Fields.Item("replacement") + "%' "
						End if
						sql_group = sql_group + " WHEN field_1 like '" +  oRS_similares.Fields.Item("replacement") + "%' THEN " & CStr(order_count)
						
						order_count = order_count + 1
						oRS_similares.MoveNext
					Loop
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			'CONCATENATE sql_group if HAVE VALUE at END using (THIS IS FOR THE SPECIFIC ORDER UN SQL RESULT)
			'ORDER BY with SPECIFIC ORDER
			'Order by (Case 
			'	When field_1 = "MD343605" then 0
			'	When field_1 = "MD172449" then 1 
			'	Else field_1 End) ;
		
			
		If Request("Page") = "" Then
		
			'response.write special_items
								
			If len(sql_group) > 0 Then
				sql_group = " ORDER BY (CASE " + sql_group + " ELSE fam_make_item END)"
				Session("SQLSearch") = sql_select_price + sql_parts + sql_group + " ;"
				
				'Session("SQLSearch") = "SELECT field_1, field_2, field_3, field_5, fam_make_item FROM partmst1_distinct WHERE (field_3 > 0) AND field_1 = 'MD343605' ORDER BY (CASE WHEN field_1 = 'MD343605' THEN 0 ELSE field_1 END) ; "
				
				
				DisplayParts(15)	
				'response.Write sql_select_price + sql_parts + sql_group + " ;"
			Else
				If special_items = "ON" then
					'response.Write sql_select_price + " ;"				
					Session("SQLSearch") = sql_select_price + " ORDER BY field_1 ;"
					DisplayParts(15)	
				else 
					If len(family_keyword_list) > 0 Then
						sql_family = sql_family + " field_6 = '" + family_keyword_list + "' "
						Session("SQLSearch") = sql_select_price + sql_parts + sql_family + " ORDER BY fam_make_item ;"
						'response.Write sql_select_price + sql_parts + sql_family + " ;"
						DisplayParts(15)	
					End if	
				End if
			End if
		Else
			'response.Write "session"
			DisplayParts(15)
		End if
			
	' MODEL SEARCH
	Else
	
	'	Response.Write "entro"	
				
			sql_select_price = sql_select_price + " (" + MultipleModelKeywords("field_2",model) + ") "
		' ggg
		If len(family_keyword_list) > 0 Then
			sql_family = sql_family + " AND field_6 = '" + family_keyword_list + "' "
			Session("SQLSearch") = sql_select_price + sql_family + " ;"
	'		Response.Write Session("SQLSearch")			
		else
		' ggg	
			Session("SQLSearch") = sql_select_price + " ORDER BY fam_make_item ;"
		end if	
		'Response.Write Session("SQLSearch")
		DisplayParts(15)
	
	End if

%>
<br>
</div>

</td>
                        </tr>
                      </table>
                    <!-- TemplateEndEditable --></td>
                  </tr>
                </table>
              </form>
            </td>
          </tr>
          <tr>
            <td height="12" colspan="2"><div class="footer"></div></td>
          </tr>
        </table>
      </div></td>
    </tr>
    <br><br><br><br>
</table>
  <br><br><br>
</body>
</html>
<%
dd_category.Close()
Set dd_category = Nothing
%>
