
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
%>

<!DOCTYPE HTML>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation</title>

<link rel="icon" href="images/favicon.ico" type="image/x-icon" /> 
<link rel="shortcut icon" href="images/favicon.ico" type="image/x-icon" />

<link href="overseas.css" rel="stylesheet" type="text/css" />
<style type="text/css">

body {
	margin-left: 10px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
	background-color: #FFF;
	background-repeat: repeat;
}

.input focus{
	border: #FC0;
}
</style>

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
	dd_category_cmd.CommandText = "SELECT field_1, field_2, english_desc FROM dbo.familicat WHERE field_2 in ( select distinct field_6 from partmst1_distinct) ORDER BY field_1 ASC" 
Else
	dd_category_cmd.CommandText = "SELECT field_1, field_2, english_desc FROM dbo.familicat WHERE field_2 in ( select distinct field_6 from partmst1_distinct) ORDER BY english_desc ASC" 
End if

dd_category_cmd.Prepared = true

Set dd_category = dd_category_cmd.Execute
dd_category_numRows = 0

%>


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
	font-size: 10px;
	font-family: "Segoe UI", "Segoe UI Light", "Segoe UI Semibold", "Lucida Console";
	font-weight: bold;
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
body,td,th {
	font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
	font-size: 10px;
}
    </style>


<script>
	$(function() {
		$( "input:submit, a, button, .demo" ).button();
		//$( "a, .demo" ).click(function() { return false; });
	});
</script>
    
    
</head>

<body>

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
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){ 
		document.getElementById("shopping_cart").innerHTML=xmlHttp.responseText 
		document.getElementById("shopping_cart").className = "display_yes" ;
		document.getElementById("loader").className = "display_no" ;
		$( "input:submit, a, button, .demo" ).button();
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
<table border="0" cellpadding="0" cellspacing="0" class="borderaround_white">
  <tr valign="top" >
    <td width="10">&nbsp;</td>
    <td width="256" align="center" valign="top"><a href="/part_search.asp" target="_self"><img src="images/oiclogo2.gif" width="256" height="31" border="0" align="absmiddle" /></a></td>
    <td width="922" align="center" valign="top">&nbsp;</td>
    <td width="45" align="center" valign="top">&nbsp;</td>
  </tr>
  <tr valign="top" >
    <td>&nbsp;</td>
    
    <td colspan="3" align="center" valign="top"><table width="1200" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="53" colspan="2">

          
          <div>
            
            <table border="0" cellspacing="0" cellpadding="2" height="31" >
              <tr>
                <td width="5">&nbsp;</td>
                <td width="256">&nbsp;</td>
                <td width="134" align="center">
                  <p>
                    <%= GetUserName(Session("MM_Username"))%>
                    &nbsp;
                   <!--<= Session("MM_Discount") %>-->
                    </p>
                  </td>
                <td width="230" align="center">
                  <p>
                    <% if Session("MM_UserName") = "Z099" then %>
                    &nbsp;<%= GetCurrentSales(Session("MM_UserName")) %>&nbsp;<a href="monitor_beta.asp">&nbsp;MONITOR&nbsp;</a>
                    </a>
                    &nbsp;
                    <a href="catalog_beta.asp">&nbsp;UPLOAD&nbsp;</a><br>
                    
                    <% End if %>
                    </p>
                  </td>
                
                <td width="146">
                  <p>
                    &nbsp;UID&nbsp;:&nbsp;&nbsp;<%= Session("MM_Username")%>
                    </p>  
                  </td>
                
                <td width="103">
                  <p>&nbsp;FECHA&nbsp;:&nbsp;&nbsp;<%=Date()%></p>  
                </td>
                
                <td width="148" align="center">
                  <a id="completar" name="completar" href="cart.asp">&nbsp;<%= Lang("completar_orden") %>&nbsp;</a>
                </td>
                
                <td width="116" align="center">
                  <a href="<%= MM_Logout %>" >&nbsp;<%= Lang("salir") %>&nbsp;</a>
                </td>
                
                <td width="26">
                <div id="loader"><img src="/overseaspr/images/ajax-loader.gif" width="16" height="16" /></div>
                </td>
                
                </tr>
              </table>
            </div>
          
          </td>
      </tr>
      <tr>
        <td width="246" valign="top">
          
          <% If Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/part_search.asp" or Request.ServerVariables("SCRIPT_NAME") = "/part_search.asp" Then %>
          
          <table border="0" cellpadding="0" cellspacing="0" width="100%" class="tablas" style="background-color:transparent" >
            
            <tr><td>&nbsp;</td></tr>
            
            <tr height="20">
              <td align="center" class="grayback">
                <p><%= Lang("preparado") %></p>
                </td>
              </tr>
            
              <tr>
               <td align="center"  >
                <div id="shopping_cart">
                  <% CartDisplay("N") %>
                </div>
               </td>
              </tr>
              
            </table>
          <% End if %>
          <br>
          <% If Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/cart.asp" or Request.ServerVariables("SCRIPT_NAME") = "/cart.asp" or len(Request("o")) > 0 Then %>
          <table width = "100%" border="0" cellpadding="0" cellspacing="0" class="tablas" style="background-color:transparent" >
            <tr>
              <td align="center" height='10'><%= Lang("ordenes") %></td>
              </tr>
            <tr>
              <td><div class="display_yes" width="100%" id="orders">
                <% OrdersDisplay() %>
                </div>
                <div class="display_no" id="loader2"><img src="/overseaspr/ajax-loader.gif" width="16" height="16" align="absmiddle" /></div></td>
              </tr>
            </table>
          <% End if %>
          <br>
          <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablas" style="background-color:transparent">
            <tr><td>&nbsp;</td></tr>	
            <tr>
              <td align="center" class="grayback">
                <p><%= Lang("estado_de_cuenta_actual") %></p></td>
              </tr>
            
            <tr>
              <td height="20" align="center">
                <a style="width:98%" href="account_statement_iframe.asp" title="Estado de Cuenta" target="_self"><%= Lang("estado_de_cuenta_actual") %></a>            
                </td>
              </tr>
            
            <tr><td>&nbsp;</td></tr>
            
            <tr>
              <td align="center" height="20" class="grayback"><p>OVERSEAS IMPORT</p></td>
              </tr>
            
            <tr>
              <td align="center" >
                <a style="width:98%" href="http://www1.overseaspr.com/soportes/index.html" title="Catálogo de Soportes" target="_self"><%= Lang("catalogo") %></a>
                </td>
              </tr>
            
            <tr><td>&nbsp;</td></tr>
            
            </table>
          
          
          </td>
        
        <td width="963" align="left" valign="top"><form id="search_part2" action="part_search.asp" method="post">
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
          <table width="960" border="0" cellpadding="0" cellspacing="2" class="tablasFORMATO">
            <tr>
              <td height="84" colspan="2" align="center" valign="top">
              	<br>
                <table width="966" border="0" cellpadding="0" cellspacing="0" style="font-size:9px" >
                  <tr height="20" style="font-size:9px" >
                  	<td width="9">&nbsp;</td>
                    <td width="197" height="46" align="center">
                      &nbsp;<%= Lang("family_keyword") %><br>
                      <input type="text" class="ui-widget-content" id="fks"  onKeyUp="searchSel()" size="16"></td>
                    <td width="240" align="center">&nbsp;<%= Lang("family_category") %><br><select name="fcs" class="ui-widget-content" id="fcs" >
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
                      &nbsp;<%= Lang("modelo") %><br><input name="model" type="text" class="ui-widget-content" id="fs2" size="16"></td>
                    
                    <td width="90" align="center">
                      <p><%= Lang("Specials") %><br>
                      <input name="special_items"  type="checkbox" class="ui-widget" id="special_items" />&nbsp;</p></td>
                    
                    <td width="80" align="center">
                      <p><%= Lang("Liquidations") %><br>
                      <input type="checkbox" name="liquidation_items" id="liquidation_items" />
                        </p></td>
                    
                    <td width="185" align="center" valign="middle" >
                      <div style="background-color:#FFC;width:150px"  >
                        <input style="font-size:12px;width:140px" class="ui-widget" name="submit" type="submit" id="submit" value="<%=Lang("buscar")%>">
                        </div>
                        
                      </td>
                  </tr>
                  </table>
                  <hr style="width:95%">
                  		
                </td>
              </tr>
            <tr>
              <td width="88" height="545" align="center" valign="top">
                <br>
                <table width="88" border="0" cellpadding="1" cellspacing="1" align="left" class="tablas" >
                  <tr valign="top">
                    <td colspan="2" align="center"><p><%= Lang("entre_10_piezas") %></p></td>
                    </tr>
                  <tr>
                    <td width="15" height="25" align="center" valign="middle">&nbsp;1</td>
                    <td width="66" align="left" valign="middle"><input name="p1" type="text" class="search_box" id="p1" size="9" maxlength="15" /></td>
                    </tr>
                  <tr>
                    <td height="25" align="center" valign="middle">&nbsp;2</td>
                    <td height="25" align="left" valign="middle"><input name="p2" type="text" class="search_box" id="p2" size="9" maxlength="15" /></td>
                    </tr>
                  <tr>
                    <td height="25" align="center" valign="middle">&nbsp;3</td>
                    <td height="25" align="left"><input name="p3" type="text" class="search_box" id="p3" size="9" maxlength="15" /></td>
                    </tr>
                  <tr>
                    <td height="25" align="center" valign="middle">&nbsp;4</td>
                    <td height="25" align="left" valign="middle"><input name="p4" type="text" class="search_box" id="p4" size="9" maxlength="15" /></td>
                    </tr>
                  <tr>
                    <td height="25" align="center" valign="middle">&nbsp;5</td>
                    <td height="25" align="left" valign="middle"><input name="p5" type="text" class="search_box" id="p5" size="9" maxlength="15" /></td>
                    </tr>
                  <tr>
                    <td height="25" align="center" valign="middle">&nbsp;6</td>
                    <td height="25" align="left" valign="middle"><input name="p6" type="text" class="search_box" id="p6" size="9" maxlength="15" /></td>
                    </tr>
                  <tr>
                    <td height="25" align="center" valign="middle">&nbsp;7 </td>
                    <td height="25" align="left" valign="middle"><input name="p7" type="text" class="search_box" id="p7" size="9" maxlength="15" /></td>
                    </tr>
                  <tr>
                    <td height="25" align="center" valign="middle">&nbsp;8 </td>
                    <td height="25" align="left" valign="middle"><input name="p8" type="text" class="search_box" id="p8" size="9" maxlength="15" /></td>
                    </tr>
                  <tr>
                    <td height="25" align="center" valign="middle">&nbsp;9</td>
                    <td height="25" align="left" valign="middle"><input name="p9" type="text" class="search_box" id="p9" size="9" maxlength="15" /></td>
                    </tr>
                  <tr>
                    <td height="25" align="center" valign="middle">10 </td>
                    <td height="25" align="left" valign="middle"><input name="p10" type="text" class="search_box" id="p10" size="9" maxlength="15" /></td>
                    </tr>
                  <tr>
                    <td height="25" align="center" valign="middle">&nbsp;</td>
                    <td height="25" align="left" valign="middle">&nbsp;</td>
                    </tr>
                  <tr>
                    <td height="25" align="center" valign="middle">&nbsp;</td>
                    <td height="25" align="left" valign="middle">&nbsp;</td>
                    </tr>
                  </table></td>
              <td width="876" align="center" valign="top">
              <!-- TemplateBeginEditable name="main_page" -->
                <table width="100%" border="0" cellpadding="0" cellspacing="0" >
                  <tr>
                    <td align="center" >
                    
                    <div class="tablas">
                      <%

model = Trim(Ucase(Request("model")))
family_keyword_list = Trim(FamilyCode(Trim(Request("fcs"))))
special_items = Trim(Ucase(Request("special_items")))
liquidation_items = Trim(Ucase(Request("liquidation_items")))

If special_items = "ON" then

	'Response.Write( "Specials: " + special_items )
	sql_select_price = "SELECT field_1, field_2, field_4, field_5, familia_descripcion, english_version, 'N/A' as replacement , family_description , fam_make_item FROM dbo.partmst1_distinct WHERE field_1 in ( select dbo.prespecials.specials from dbo.prespecials ) "

elseIf liquidation_items = "ON" then

	'Response.Write( "Liquidation: " + liquidation_items ) 
	sql_select_price = "SELECT field_1, field_2, field_4, field_5, familia_descripcion, english_version, 'N/A' as replacement , family_description , fam_make_item FROM dbo.partmst1_distinct WHERE field_1 in ( select field_1 from dbo.prod_liqui ) "

else
	
	If Session("user_level") = 3 Then
		
' join with similar
'		sql_select_price = "SELECT field_1, field_2, field_3, field_5, familia_descripcion, english_version, similar.replacement, family_description , fam_make_item FROM {oj dbo.partmst1_distinct LEFT OUTER JOIN dbo.similar ON dbo.partmst1_distinct.field_1 = dbo.similar.partid} WHERE "

		sql_select_price = "SELECT field_1, field_2, field_3, field_5, familia_descripcion, english_version, replacement, family_description , fam_make_item FROM dbo.partmst1_distinct WHERE  "
		
	Else
	
'		sql_select_price = "SELECT field_1, field_2, field_4, field_5, familia_descripcion, english_version, similar.replacement , family_description , fam_make_item FROM {oj dbo.partmst1_distinct LEFT OUTER JOIN dbo.similar ON dbo.partmst1_distinct.field_1 = dbo.similar.partid} WHERE "

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
				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p1 & "' or REPLACE(replacement,' ','') = '" & trim(p1) & "' ;", MM_overseaspr_STRING

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
			'		oRS_similares.Close
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
			
				'SIMILAR STEPS
	'			oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p2 & "' ;", MM_overseaspr_STRING

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
				
				'SIMILAR STEPS
				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p3 & "' or REPLACE(replacement,' ','') = '" & trim(p3) & "' ;", MM_overseaspr_STRING
	
'				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p3 & "' ;", MM_overseaspr_STRING
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
	'				oRS_similares.Close
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
				
				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p4 & "' or REPLACE(replacement,' ','') = '" & trim(p4) & "' ;", MM_overseaspr_STRING

				'SIMILAR STEPS
'				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p4 & "' ;", MM_overseaspr_STRING
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
	'				oRS_similares.Close
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
				
				'SIMILAR STEPS
'				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p5 & "' ;", MM_overseaspr_STRING
				
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
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p6) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '%" + p6 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '%" + p6 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '%" +  p6 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
'				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p6 & "' ;", MM_overseaspr_STRING

				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p6 & "' or REPLACE(replacement,' ','') = '" & trim(p6) & "' ;", MM_overseaspr_STRING

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
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p7) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '%" + p7 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '%" + p7 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '%" +  p7 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
'				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p7 & "' ;", MM_overseaspr_STRING

				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p7 & "' or REPLACE(replacement,' ','') = '" & trim(p7) & "' ;", MM_overseaspr_STRING

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
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p8) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '%" + p8 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '%" + p8 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '%" +  p8 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
'				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p8 & "' ;", MM_overseaspr_STRING

				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p8 & "' or REPLACE(replacement,' ','') = '" & trim(p8) & "' ;", MM_overseaspr_STRING

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
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p9) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '%" + p9 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '%" + p9 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '%" +  p9 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
'				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p9 & "' ;", MM_overseaspr_STRING
				
				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p9 & "' or REPLACE(replacement,' ','') = '" & trim(p9) & "' ;", MM_overseaspr_STRING

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
	'				oRS_similares.Close
				End if
				oRS_similares.Close			
			End if
			
			If len(p10) > 0 Then 
				If len(sql_parts) = 0 Then
					sql_parts = sql_parts + " field_1 like '%" + p10 + "%' "
				Else
					sql_parts = sql_parts + " OR field_1 like '%" + p10 + "%' "
				End if
				sql_group = sql_group + " WHEN field_1 like '%" +  p10 + "%' THEN " & CStr(order_count)
				
				'SIMILAR STEPS
'				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p10 & "' ;", MM_overseaspr_STRING
				
				oRS_similares.Open "SELECT partid FROM similar WHERE replacement = '" & p10 & "' or REPLACE(replacement,' ','') = '" & trim(p10) & "' ;", MM_overseaspr_STRING

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
				
				DisplayParts(15)	
				' response.Write "1" + sql_select_price + sql_parts + sql_group + " ;"
				
			Else
				If special_items = "ON" or liquidation_items = "ON" then
					'response.Write sql_select_price + " ;"				
					Session("SQLSearch") = sql_select_price + " ORDER BY fam_make_item ;"
					DisplayParts(15)	
				else 
					If len(family_keyword_list) > 0 Then
						sql_family = sql_family + " field_6 = '" + family_keyword_list + "' "
						Session("SQLSearch") = sql_select_price + sql_parts + sql_family + " ORDER BY fam_make_item ;"
					' 	response.Write "2 Price" + sql_select_price + " Parts " + sql_parts + " family " + sql_family + " ;"
					' 	response.Write "3" + Session("SQLSearch")
						
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
                    </div>                       <script language="javascript">

document.getElementById("loader").className = "display_no" ;

</script>
                      <br /></td>
                    </tr>
                  </table>
                <!-- TemplateEndEditable --></td>
              </tr>
            </table>
          </form></td>
      </tr>
      <tr>
        <td height="12" colspan="2"><div class="footer"></div></td>
      </tr>
      
    </table></td>
  </tr>
</table>
<br>
</body>
</html>
<%
dd_category.Close()
Set dd_category = Nothing
%>
