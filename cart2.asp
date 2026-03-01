<!-- InstanceBegin template="/Templates/master_template_overseas.dwt.asp" codeOutsideHTMLIsLocked="false" -->
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
<!-- InstanceBeginEditable name="doctitle" -->
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
cart_list_numRows = cart_list_numRows + Repeat1__numRows
%>
<%
Dim cart_list__MMColParam
cart_list__MMColParam = "1"
If (Session("MM_Username") <> "") Then 
  cart_list__MMColParam = Session("MM_Username")
End If
%>
<%
Dim cart_list
Dim cart_list_cmd
Dim cart_list_numRows

Set cart_list_cmd = Server.CreateObject ("ADODB.Command")
cart_list_cmd.ActiveConnection = MM_overseaspr_STRING
cart_list_cmd.CommandText = "SELECT shop_auto_id, shop_product_id, shop_quantity, shop_part_price FROM dbo.clients_cart WHERE shop_client_user = ?" 
cart_list_cmd.Prepared = true
cart_list_cmd.Parameters.Append cart_list_cmd.CreateParameter("param1", 200, 1, 20, cart_list__MMColParam) ' adVarChar

Set cart_list = cart_list_cmd.Execute
cart_list_numRows = 0
%>
<title>Overseas Import Corporation - Shopping Cart</title>
<!-- InstanceEndEditable -->
<link href="overseas.css" rel="stylesheet" type="text/css" />
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
body,td,th {
	font-size: 10px;
}
</style>
<!-- InstanceEndEditable -->
<style type="text/css">
body,td,th {
	font-size: 9px;
}
body {
	background-color: #FFF;
	background-repeat: repeat;
	margin-left: 10px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
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
<table border="0" align="center" cellpadding="0" cellspacing="0" class="borderaround_white">
<tr valign="top" width="100%">
  <td align="center" valign="top"><table align="center" width="1183" height="350" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td colspan="2"><table width="1170" height="454" cellpadding="0" cellspacing="0" class="borderaround_white">
        <tr>
          <td width="25%" align="center" ><table height="400px" border="0" align="center" cellpadding="0" cellspacing="0" class="borderaround_white">
            <tr>
              <td height="450" align="center" valign="top"><table align="center" width="1183" height="350" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td colspan="2">
                  <table width="1000" cellspacing="0" cellpadding="0" class="borderaround">
                    <tr>
                      <td width="2%" height="48">&nbsp;</td>
                      <td width="21%" style="background-color:#FFF"><img src="images/oiclogo2.gif" width="256" height="31" align="absmiddle"></td>
                      <td width="21%"><% if Session("MM_UserName") = "Z099" then %>
                        <%= GetCurrentSales(Session("MM_UserName")) %>
                        <% End if %></td>
                      <td width="31%"><div id="linea"><%= GetUserName(Session("MM_Username"))%><br>
                        UID&nbsp;:&nbsp;&nbsp;<%= Session("MM_Username")%>&nbsp;&nbsp;&nbsp;<%= Lang("fecha") %>:&nbsp;:&nbsp;&nbsp;<%=Date()%></div></td>
                      <td width="25%" align="center" class="tablas_font12"><a href="cart.asp" class="Overseas_Title"><%= Lang("completar_orden") %></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%= MM_Logout %>" class="Overseas_Title"><%= Lang("salir") %></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td width="236" valign="top"></td>
                  <td align="left" valign="top">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top"><% If Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/part_search.asp" or Request.ServerVariables("SCRIPT_NAME") = "/part_search.asp" Then %>
                    <table width="230px" border="0" cellpadding="0" cellspacing="0" >
                      <tr>
                        <td align="center" ><div class="grayback"><%= Lang("preparado") %></div></td>
                      </tr>
                      <tr>
                        <td align="center"><div class="display_yes" id="shopping_cart">
                          <% CartDisplay("N") %>
                        </div>
                          <div class="display_no" id="loader"><img src="ajax-loader.gif" width="20" height="20"></div></td>
                      </tr>
                    </table>
                    <% End if %>
                    <% If Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/cart.asp" or Request.ServerVariables("SCRIPT_NAME") = "/cart.asp" or len(Request("o")) > 0 Then %>
                    <table width = "230px" border="0" cellpadding="0" cellspacing="0" class="left_menu" >
                      <tr>
                        <td width = "230px" align="center" height='10'><div class="grayback"><%= Lang("ordenes") %></div></td>
                      </tr>
                      <tr>
                        <td><div class="display_yes" width="230px" id="orders">
                          <% OrdersDisplay() %>
                        </div>
                          <div class="display_no" id="loader2"><img src="ajax-loader.gif" width="16" height="16"></div></td>
                      </table>
                    <% End if %>
                    <table width="230px" border="0" cellpadding="0" class="display_yes" cellspacing="0" >
                      <tr>
                        <td width = "230px" align="center" height='10'><div class="grayback"><%= Lang("ACCOUNTSTATEMENT") %></div></td>
                      </tr>
                      <tr>
                        <td align="center">&nbsp;</td>
                      </tr>
                      <tr>
                        <td height='20' align="center"><div class="display_yes" id="orders"><a href="account_statement_iframe.asp" title="Estado de Cuenta" target="_self"><%= Lang("estado_de_cuenta_actual") %></a><br><br>
                          <br>
                        </div></td>
                      </tr>
                      <tr>
                        <br><br><td align="center" height='20' ><div class="grayback">OVERSEAS IMPORT</div></td>
                      </tr>
                      <tr>
                        <td align="center">&nbsp;</td>
                      </tr>
                      <tr>
                        <td align="center" class="display_yes"><div class="display_yes" id="orders"><a href="http://69.89.36.77/soportes/index.html" title="Catálogo de Soportes" target="_self" class="display_yes"><%= Lang("catalogo") %></a><br></div>
                          <br>
                          <br></td>
                      </tr>
                    </table>
                  <td width="947" align="left" valign="top"><form id="search_part" action="part_search.asp" method="post">
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
                    <table width="940" border="0" cellpadding="0" cellspacing="1" class="tablasFORMATO">
                      <tr>
                        <td colspan="2" align="center" valign="top"><table class="tablas" width="100%" border="1" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="21%" align="center" class="tablas_font12"><%= Lang("family_keyword") %></td>
                            <td width="26%" align="center" class="tablas_font12"><%= Lang("family_category") %></td>
                            <td width="21%" align="center" class="tablas_font12"><%= Lang("modelo") %></td>
                            <td width="32%" align="center" class="tablas_font12">&nbsp;</td>
                          </tr>
                          <tr>
                            <td align="center"><input type="text" class="search_box_NORADIUS" id="fks" onKeyUp="searchSel()"></td>
                            <td align="center"><select name="fcs" class="search_box_NORADIUS" id="fcs">
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
                            <td align="center"><input name="model" type="text" class="search_box_NORADIUS" id="fs2"></td>
                            <td align="left"><div class="grayback">
                              <input type="checkbox" name="special_items" id="special_items">
                              <label for="special_items"><%= Lang("Specials") %></label>
                              <input name="submit" type="submit" id="submit" value="<%= Lang("buscar") %>">
                            </div></td>
                          </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td width="150" align="center" valign="top"><table width=134 border="0" cellpadding="1" cellspacing="1">
                          <tr>
                            <td colspan="2" align="center"><%= Lang("entre_10_piezas") %></td>
                          </tr>
                          <tr>
                            <td width="25" height="25" align="left" valign="middle">&nbsp;1</td>
                            <td width="106" align="left" valign="middle"><input name="p1" type="text" class="search_box" id="p1" size="15" /></td>
                          </tr>
                          <tr>
                            <td height="25" align="left" valign="middle">&nbsp;2</td>
                            <td height="25" align="left" valign="middle"><input name="p2" type="text" class="search_box" id="p2" size="15" /></td>
                          </tr>
                          <tr>
                            <td height="25" align="left">&nbsp;3</td>
                            <td height="25" align="left"><input name="p3" type="text" class="search_box" id="p3" size="15" /></td>
                          </tr>
                          <tr>
                            <td height="25" align="left" valign="middle">&nbsp;4</td>
                            <td height="25" align="left" valign="middle"><input name="p4" type="text" class="search_box" id="p4" size="15" /></td>
                          </tr>
                          <tr align="left">
                            <td height="25" align="left" valign="middle">&nbsp;5</td>
                            <td height="25" align="left" valign="middle"><input name="p5" type="text" class="search_box" id="p5" size="15" /></td>
                          </tr>
                          <tr>
                            <td height="25" align="left" valign="middle">&nbsp;6</td>
                            <td height="25" align="left" valign="middle"><input name="p6" type="text" class="search_box" id="p6" size="15" /></td>
                          </tr>
                          <tr>
                            <td height="25" align="left" valign="middle">&nbsp;7 </td>
                            <td height="25" align="left" valign="middle"><input name="p7" type="text" class="search_box" id="p7" size="15" /></td>
                          </tr>
                          <tr>
                            <td height="25" align="left" valign="middle">&nbsp;8 </td>
                            <td height="25" align="left" valign="middle"><input name="p8" type="text" class="search_box" id="p8" size="15" /></td>
                          </tr>
                          <tr>
                            <td height="25" align="left" valign="middle">&nbsp;9</td>
                            <td height="25" align="left" valign="middle"><input name="p9" type="text" class="search_box" id="p9" size="15" /></td>
                          </tr>
                          <tr>
                            <td height="25" align="left" valign="middle">10 </td>
                            <td height="25" align="left" valign="middle"><input name="p10" type="text" class="search_box" id="p10" size="15" /></td>
                          </tr>
                        </table></td>
                        <td width="814" align="center" valign="top"><!-- InstanceBeginEditable name="main_page" -->
				  <table width="775" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td align="center"><% If Not cart_list.EOF Or Not cart_list.BOF Then %>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tablas">
    <tr>
      <td width="25%">&nbsp;</td>
      <td width="25%" colspan="2" align="center"><h3>Lista de piezas en el cart</h3></td>
      <td width="25%">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td width="25%">&nbsp;</td>
      <td width="25%">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td align="center"><strong>ID. DE PIEZA</strong></td>
      <td align="center"><strong>PRECIO</strong></td>
      <td align="center"><strong>CANTIDAD</strong></td>
      <td align="center">&nbsp;</td>
    </tr>
    <%
	language = Session("Lang")
	color1 = "#CCCCCC"
	color2 = "#FFFFFF"
	current_color = color1
	rc = 1
%>
    <% 
While ((Repeat1__numRows <> 0) AND (NOT cart_list.EOF)) 
%>
      <tr>
        <td height="25" align="center" bgcolor='<%=  current_color %>'><%=(cart_list.Fields.Item("shop_product_id").Value)%></td>
        <td align="center" bgcolor='<%=  current_color %>'><%=(cart_list.Fields.Item("shop_quantity").Value)%></td>
        <td align="center" bgcolor='<%=  current_color %>'><%=(cart_list.Fields.Item("shop_part_price").Value)%></td>
        <td align="center" bgcolor='<%=  current_color %>'><a href="del_p_shopping_cart.asp?p=<%=(trim(cart_list.Fields.Item("shop_auto_id").Value))%>&amp;t=1"><img src="images/del.gif" width="20" height="20" border="0" /></a></td>
      </tr>
      <% 		
If current_color = color2 then current_color = color1 else current_color = color2
rc = rc + 1
%>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  cart_list.MoveNext()
Wend
%>
    <tr>
      <td align="center">&nbsp;</td>
      <td align="center">&nbsp;</td>
      <td align="center">&nbsp;</td>
      <td align="center">&nbsp;</td>
    </tr>
    <tr>
      <td align="center">&nbsp;</td>
      <td colspan="2" align="center" class="Overseas_Title" ><p>
        <% If PartsOnCart() Then %>
        <span class="Overseas_Title"><%= Lang("metodo_envio") %></span>
        <select name="delivery_type" class="tablas">
          <option value="D"><%= Lang("entrega") %></option>
          <option value="P"><%= Lang("recoger") %></option>
        </select>
      </p>
        <p><br />
          <input name="Someter Orden" type="submit" value="Someter Orden" />
          <% End if %>
        </p></td>
      <td align="center">&nbsp;</td>
    </tr>
  </table>
  
  <% End If ' end Not cart_list.EOF Or NOT cart_list.BOF %>
                        <% If cart_list.EOF And cart_list.BOF Then %>
                        <h3>NO HAY ARTICULOS EN EL CART </h3>
                        <% End If ' end cart_list.EOF And cart_list.BOF %></td>
                      </tr>
                      <tr>
                        <td align="center">&nbsp;</td>
                      </tr>
                    </table>
                  <!-- InstanceEndEditable --></td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td height="12" colspan="2"><div class="footer"></div></td>
                </tr>
              </table></td>
            </tr>
          </table></td>
        </tr>
      </table></td>
    </tr>
  </table></td>
</tr>
</table>
<tr>
</body>
<!-- InstanceEnd --></html>
<%
dd_category.Close()
Set dd_category = Nothing
%>
<%
cart_list.Close()
Set cart_list = Nothing
%>
