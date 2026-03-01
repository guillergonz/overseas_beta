<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<!DOCTYPE HTML>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation - Order Detail</title>

<link rel="icon" href="images/favicon.ico" type="image/x-icon" /> 
<link rel="icon" href="images/favicon.ico" type="image/x-icon" />

<link href="overseas.css" rel="stylesheet" type="text/css" />

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

    body,td,th {
	font-family: "Segoe UI", "Segoe UI Light", "Segoe UI Semibold", "Lucida Console";
	font-size: 14px;
	color: #333;
}

body {
	background-image: url();
	margin-left: 10px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
}
</style>


<script>
	$(function() {
		$( "input:submit, a, button, #back, #print " ).button();
		//$( "a, #Printerfriendly" ).click(function() { return false; });
	});
</script>

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
<hr>
<table class = "borderaround_printpreview" align="center" width="800" height="470" border="0" cellpadding="0" cellspacing="0">
 <tr>
   <td height="20" align="left" valign="top"><img name="overseas_logo" src="images/oiclogo2.gif" width="256" height="31" alt="Overseas Logo"></td>
 </tr>
 <tr>
  <td width="1052" height="380" align="left" valign="top">Sales Orders - Printer friendly version <br><table width="100%" border="0" cellspacing="10" cellpadding="2">
  	    <tr>
  	      <td width="100" align="center" border = "0" >
		  <SCRIPT LANGUAGE="JavaScript"> 
			if (window.print) {
			document.write('<form><input id="print" align="center"  type=button name=print value="Print" onClick="window.print()"></form>');
			}
			</script>
          </td>
  	      <td align="left" border = "0">
            <FORM>
            <INPUT id="back" align = "center" TYPE="button" VALUE="Back" onClick="history.go(-1);return true;"></FORM>
          </td>
	      </tr>
	    </table>

    <div class="tablas">
      <table width="800" border="0" cellpadding="0" cellspacing="0" class="tablas">
        <tr>
          <td width = "24%" align="right"><%= Lang("detalles_orden") %>&nbsp;</td>
          <td width="28%" align="left"><strong><%=(order_detail.Fields.Item("order_number").Value)%>&nbsp;<%= (order_detail.Fields.Item("order_user").Value) %></strong></td>
          <td width="20%" align="right"><%= Lang("fecha_y_hora") %>:&nbsp;</td>
          <td width="28%" align="left"><strong><%=(order_detail.Fields.Item("order_date").Value)%></strong></td>
          </tr>
        <tr>
          <td align="right">&nbsp;</td>
          <td align="left">&nbsp;</td>
          <td align="right">&nbsp;</td>
          <td align="center">&nbsp;</td>
          </tr>
        <tr>
          <td align="right"><div class="display_yes"><%= Lang("estatus") %>&nbsp;&nbsp;</div></td>
          <td align="left"><label for="order_status"></label>
            
            <select name="order_status" disabled="disabled" class="display_yes" id="order_status">
            <option value="O" <%If (Not isNull((order_detail.Fields.Item("order_status").Value))) Then If ("O" = CStr((order_detail.Fields.Item("order_status").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("status_pending") %></option>
            <option value="P" <%If (Not isNull((order_detail.Fields.Item("order_status").Value))) Then If ("P" = CStr((order_detail.Fields.Item("order_status").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("status_processed") %></option>
            </select>
            
            </td>
          <td align="right"><%= Lang("metodo_envio") %>&nbsp;</td>
          <td align="left"><label for="delivery_type"></label>
            <select name="delivery_type" class="display_yes" id="delivery_type" disabled>
              <option value="D" <%If (Not isNull((order_detail.Fields.Item("deliv_type").Value))) Then If ("D" = CStr((order_detail.Fields.Item("deliv_type").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("entrega") %></option>
              <option value="P" <%If (Not isNull((order_detail.Fields.Item("deliv_type").Value))) Then If ("P" = CStr((order_detail.Fields.Item("deliv_type").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("recoger") %></option>
            </select></td>
          </tr>
        <tr>
          <td align="center">&nbsp;</td>
          <td align="center">&nbsp;</td>
          <td align="center">&nbsp;</td>
          <td align="center">&nbsp;</td>
          </tr>
          </table>
          
      <table width="800" height ="100%" border="0" cellpadding="0" cellspacing="0" class="tablas">
          
        <tr>
          <td colspan="5" align="center"><div class="search_box"><%= Lang("detalles_piezas") %>
            </div></td>
          </tr>
        <tr>
          <td width="13%" align="center">&nbsp;</td>
          <td width="43%" align="center">&nbsp;</td>
          <td width="18%" align="center">&nbsp;</td>                                
          <td width="14%" align="center">&nbsp;</td>
          <td width="12%" align="center">&nbsp;</td>
          </tr>
        <tr>
          <td align="center"><%= Lang("num_pieza") %></td>
          <td align="center"><%= Lang("descripcion") %></td>
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
	ldescription = ""
	
	if trim(order_detail.Fields.Item("order_part").Value) = "ZZZNOF" then
	
		ldescription = order_detail.Fields.Item("comments").Value
	
	else

		If Session("lang") = "E" Then
			if oRS2.Fields.Item("english_version") > "" then
				ldescription = CStr(oRS2.Fields.Item("english_version"))
			else
				ldescription = CStr(oRS2.Fields.Item("field_2"))
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
          <td align="center" bgcolor='<%=  current_color %>'><%= ldescription %></td>
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
          <td align="center">&nbsp;</td>
          <td align="center">&nbsp;</td>
          <td align="center">&nbsp;</td>
          </tr>
        <tr>
          <td align="center">&nbsp;</td>
          <td align="center">&nbsp;</td>
          <td align="center" class="table_header">TOTAL</td>
          <td align="center" class="table_header"><%= CurrencyConvert(total) %></td>
          </tr>
        <tr>
          <td align="center">&nbsp;</td>
          <td align="center">&nbsp;</td>
          <td align="center" class="table_header">&nbsp;</td>
          <td align="center" class="table_header">&nbsp;</td>
          </tr>
      </table>
    </div>
      </div></td>
  </tr>
</table>
<hr>


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
