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
	background-image: url(AnimatedFrom/images/bg.gif);
	margin:0;
	background-color: #FFF;
	background-repeat: repeat;
}
.input focus{
	border: #FC0;
}
body,td,th {
font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
font-size: 14px;
color: #000;
}
</style>
</head>
<body>






<script type="text/javascript" charset="utf-8">

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
			
		},
		error: function(xhr,textStatus, errorThrown){
			alert(textStatus);
		}               
	});				

};
</script>

<script type="text/javascript" charset="utf-8">
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
		},
		error: function(xhr,textStatus, errorThrown){
			alert(textStatus);
		}               
	});				

};		
</script>
															      
<script type="text/javascript" charset="utf-8">
function ValidateCart(part,amt) {
	amount = document.getElementById(amt).value;
	AddToCart(part,amount);
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




<div class="container">
   
  	<div class="row">
    
    	<div class="col-xs-6 col-lg-3">
        <a href="part_search.asp" target="_self"><img src="images/oiclogo2.gif" width="256" height="31" alt="logo"  ></a>
        </div>
          
	</div>    
  	<br>
 
    <% if count_records > 0 OR Session("MM_UserName") = "Z099" then %>
      
    <div class="row">         
        <div class="col-xs-8 pull-left" >
        
        	<p><input class="btn btn-primary" style="width:98%" name="Printerfriendly" type="button" id="Printerfriendly" value="PRINT / IMPRIMIR " onClick="window.print();" /></p>
            
            <div class="table-responsive">

               <table class="table table-striped"  >              
                <tr>
                <td width = "19%" align="right"><%= Lang("detalles_orden") %>&nbsp;</td>
                <td width = "39%" align="left"><%=(order_detail.Fields.Item("order_number").Value)%>&nbsp;<%= (order_detail.Fields.Item("order_user").Value) %></td>
                <td width="18%" align="right"><strong><%= Lang("fecha_y_hora") %>&nbsp;</strong></td>
                <td width="24%" colspan="2" align="left"><%= trim(order_detail.Fields.Item("order_date").Value)%></td>
                </tr>

                <tr>
                <td align="right"><strong>&nbsp;<%= Lang("estatus") %>&nbsp;</strong></td>
                <td align="left">
                <select style="width:120px"  name="order_status" disabled="disabled" class="form-control" id="order_status">
                <option value="O" <%If (Not isNull((order_detail.Fields.Item("order_status").Value))) Then If ("O" = CStr((order_detail.Fields.Item("order_status").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("status_pending") %></option>
                <option value="P" <%If (Not isNull((order_detail.Fields.Item("order_status").Value))) Then If ("P" = CStr((order_detail.Fields.Item("order_status").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("status_processed") %></option>
                </select>
                </td>
                <td align="right">
                <strong><%= Lang("metodo_envio") %></strong>
                </td>
                <td colspan="2" >
                <select style="width:120px" name="delivery_type" class="form-control" id="delivery_type" disabled="disabled">
                <option value="D" <%If (Not isNull((order_detail.Fields.Item("deliv_type").Value))) Then If ("D" = CStr((order_detail.Fields.Item("deliv_type").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("entrega") %></option>
                <option value="P" <%If (Not isNull((order_detail.Fields.Item("deliv_type").Value))) Then If ("P" = CStr((order_detail.Fields.Item("deliv_type").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%>><%= Lang("recoger") %></option>
                </select>
                </td>
                </tr>
               
                <tr>
                <td align="center"><%= Lang("num_pieza") %></td>
                <td align="left"><%= Lang("descripcion") %></td>
                <td align="center"><%= Lang("ordenado") %></td>
                <td align="center"><%= Lang("precio") %></td>
                <td align="center">SUB-TOTAL</td>
                </tr>
<%
	rc = 1
	total = 0
	ldescription = ""
	
	While ((Repeat1__numRows <> 0) AND (NOT order_detail.EOF)) 

	' Record set to get description in english or spanish for part item ...
	Set oRS2 = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.partmst1_distinct.field_2, dbo.partmst1_distinct.english_version FROM dbo.partmst1_distinct WHERE dbo.partmst1_distinct.field_1 = '" + CStr(order_detail.Fields.Item("order_part").Value)  + "' ;"
	oRS2.Open strSQL, MM_overseaspr_STRING
	
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
	oRS2.Close
	Set oRS2 = Nothing	
%>
                                     
                <tr>
                <td height="20" align="center" ><%=(order_detail.Fields.Item("order_part").Value)%></td>
                <td><%= ldescription %></td>
                <td align="center" ><%=(order_detail.Fields.Item("order_qty").Value)%></td>
                <td align="center" ><%=(CurrencyConvert(order_detail.Fields.Item("item_price").Value))%></td>
                <td align="center" ><%= CurrencyConvert(order_detail.Fields.Item("order_qty").Value * order_detail.Fields.Item("item_price").Value) %></td>
                </tr>
<% 		
rc = rc + 1
total = total + (order_detail.Fields.Item("item_price").Value * order_detail.Fields.Item("order_qty").Value)

  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  order_detail.MoveNext()
Wend
%>
                <tr>
                <td colspan="2">&nbsp;</td>
                <td align="center">TOTAL TAX<br>CITY TAX<br>STATE TAX</td>
                <td align="center" >
                <% 
                If Session("MM_CityTax")="Y" Then 
                vCityTax  = round( (total * .01) ,2)
                'vStateTax = round( (total * .06) ,2)
                vStateTax = round( (total * .105) ,2)
                response.write(  CurrencyConvert(total) + "<br>" + CurrencyConvert(vCityTax) + "<br>" + CurrencyConvert(vStateTax))
                End IF
                %>
                </td>
                </tr>

                <tr>
                <td align="center">&nbsp;</td>
                <td align="center">&nbsp;</td>
                <td align="center">TOTAL</td>
                <td align="center">									  
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

          		</table>
                                  
                                  
               
            
            <!--table responsive-->
            </div>
		</div>      
    	<!--class=row-->  
	</div>
    <% End If %>
    
	</form>            
  	<script type="text/javascript" src="../bootstrap-3.3.6-dist/js/bootstrap.min.js"></script>
    <!--div container-->
</div>

</body>
</html>
