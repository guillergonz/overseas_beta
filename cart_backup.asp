<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "/index.asp"
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

<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation :: Shopping Cart</title>

<link rel="icon" href="images/favicon.ico" type="image/x-icon" /> 
<link rel="icon" href="images/favicon.ico" type="image/x-icon" />


<link href="overseas.css" rel="stylesheet" type="text/css" />

<style type="text/css">
body,td,th {
	font-size: 12px;
	font-family: "Lucida Sans Unicode", "Lucida Grande", sans-ser;
	color: #000;
}
body {
	background-color: #FFF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-repeat: repeat;
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
	/* [disabled]font-size:9px; */
}

.ui-button {
	display: inline-block;
	position: relative;
	padding: .1em;
	/* [disabled]margin-right: 0.1em; */
	text-decoration: none;
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
}
.ui-button-text-only .ui-button-text { padding: 0; margin:0; line-height:normal; font-size:12px; }
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
	font-size: 12px;
	color: #000;
}
</style>

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
//document.getElementById("shopping_cart").className = "display_no" ;

$("#loader").fadeIn("slow");
$("#loader").removeClass("display_no").addClass("display_yes")
//document.getElementById("loader").className = "display_yes" ;

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
		$("#loader").fadeOut("fast");
		
		$( "input:submit, a, button, .demo, .submit" ).button();
		
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

function DelFromCart(str) { 
	xmlHttp2=GetXmlHttpObject()
	if (xmlHttp2==null) {
	alert ("Browser does not support HTTP Request")
	return
	} 
	//$("#txtHint").className = "display_no" ;
	//document.getElementById("shopping_cart").className = "display_no" ;
	//document.getElementById("loader").className = "display_yes" ;
	
	$("#loader").fadeIn("slow");
	$("#loader").removeClass("display_no").addClass("display_yes")
	
	
	var url2="del_p_shopping_cart.asp"
	url2=url2+"?p=" + str + "&t=1"  
	url2=url2+"&sid="+Math.random()
	xmlHttp2.onreadystatechange=stateChanged2 
	xmlHttp2.open("GET",url2,true)
	xmlHttp2.send(null)
}

function stateChanged2() { 
	if (xmlHttp2.readyState==4 || xmlHttp2.readyState=="complete"){ 
		document.getElementById("shopping_cart").innerHTML=xmlHttp2.responseText 
//		document.getElementById("shopping_cart").className = "display_yes" ;
		$( "input:submit, a, button, .demo, .submit" ).button();
		
		$("#loader").removeClass("display_yes").addClass("display_no")		
		$("#loader").fadeOut("fast");
	} 
} 
function GetXmlHttpObject2(){ 
	var objXMLHttp2=null
	if (window.XMLHttpRequest)	{
		objXMLHttp2=new XMLHttpRequest()
	}
	else if (window.ActiveXObject){
		objXMLHttp2=new ActiveXObject("Microsoft.XMLHTTP")
	}
	return objXMLHttp2
}
</script>


<script>
$(document).ready(function() {
		$( "input:submit, a, button, .demo, #DelFromcart, #UpdCart" ).button();
		//$( "a, .DelFromCart, .monitor" ).click(function() { return false; });

	$( "#grayback" ).fadeTo( "slow" , 0.2, function() {
		 $( "#grayback" ).fadeTo( "slow" , 1);
    	// Animation complete.
  	});
});
</script>

</head>

<body>

<div id="grayback" class="grayback" style="height:32px;margin-left:400px;width:180px">
  <p >NEW CATALOG !
      <a style="margin-left:10px;margin-right:10px;width:160px;float:left;font-size:14px;color:navy;padding:0px" href="catalog_maint_CAT.asp" title="Catálogo de Soportes" target="_self"><%= Lang("catalogo") %>
    </a>
  </p>
</div>

<table border="0" cellpadding="0" cellspacing="0" class="borderaround_white" >
	<tr valign="top" >
      <td>&nbsp;</td>
      
    <td width="1209" align="center" valign="top"><table align="center"  border="0" cellpadding="0" cellspacing="0" >
    <tr>
      <td colspan="2">
        
                 
          
          <table width="1079" height="26" border="0" cellpadding="2" cellspacing="0" >
            <tr>
              
              <td width="5" height="26">&nbsp;</td>
              <td width="256">
              
               <div id="logooicggg" style="margin-left:auto;margin-right:auto;width:100%; text-align:center; margin-top:3px;margin-bottom:3px" >	
                  <a href="part_search.asp" target="_self"><img src="images/oiclogo2.gif" width="256" height="31" border="0" align="absmiddle" /></a>
					<br>                  
                  <!-- 9/24/2015  User ID:&nbsp;<= GetUserName(Session("MM_Username"))%>-->
                  
                </div>
                </td>
              <td width="275" align="center">
                
                  <span style="text-align:center;margin:3px">Welcome&nbsp;<%= GetUserName(Session("MM_Username"))%> </span>
                  <% if LEN(Session("MM_Multi_Username")) > 0 then %>
                  <br><label style="color:blue"><%= Session("MM_Multi_Username") %></label>
                  <% End If %>
                  
                </td>
              
              <td width="47">
                <p>&nbsp;</p>  
                </td>
              
              <td width="102">
                <p>&nbsp;FECHA&nbsp;:&nbsp;&nbsp;<%=Date()%></p>  
                </td>
              
              <td  width="163" align="center">
				</td>
              <td  width="183" align="center">  
              <!--&nbsp;&nbsp;<a href="<%= MM_Logout %>" style="width:100px" ><%= Lang("salir") %></a>&nbsp;&nbsp;-->
              
               <a href="<%= MM_Logout %>" style="font-size:14px;height:30px;width:150px;vertical-align:middle" ><%= Lang("salir") %></a>
               
              </td>
              
              <td width="16">
              
                <div id="loader" class="display_no"><img src="/ajax-loader.gif" width="16" height="16" /></div>
                </td>
              
            </tr>
            </table>
                  
        
      </td>
      </tr>
          <tr>
            <td width="150" valign="top">
			<br><br>
			<% 
			  
			  
			  
			  If Request.ServerVariables("SCRIPT_NAME") = "/cart.asp" or Request.ServerVariables("SCRIPT_NAME") = "/part_search.asp" Then %>
              <table width="150px" border="0" cellpadding="0" cellspacing="0" style="background-color:transparent" >

              <tr height="20">
                <td align="center" class="grayback" >
                <p><%= Lang("preparado") %></p></td>
                </tr>
                
                <tr>
                <td align="center">
	                
                    <div class="display_yes" id="shopping_cart">
                      <% CartDisplay("N") %>
                    </div>
    	          	
                    
                </td>
                </tr>
              </table>
                
			  <% End if %>
 			  
              <br>
              	
              <table width="150px" border="0" cellpadding="0" class="tablas" cellspacing="0" style="background-color:transparent">
                <tr>
                  <td width = "100%" align="center" height='10'>
                  
                  <% If LEN(Session("MM_Multi_Username")) = 0 or Session("MM_Multi_Username") = "CARLE BETANCOURT" then %>
                  	<div class="grayback"><%= Lang("ACCOUNTSTATEMENT") %></div>
                  <% End IF %>
                  </td>
                </tr>
                
               
                <tr>
                  <td height='20' align="center">
                  
                  <% If LEN(Session("MM_Multi_Username")) = 0 or Session("MM_Multi_Username") = "CARLE BETANCOURT" then %>
                  
                  <div class="display_yes" id="orders">
                  <a style="width:148px" href="account_statement_iframe.asp" title="Estado de Cuenta" target="_self"><%= Lang("estado_de_cuenta_actual") %></a>
                  </div>
                  
				  <% End IF %>
                  
                  </td>
                </tr>
                
                <tr><td>&nbsp;</td></tr>
                
                <tr>
                  <td align="center" ><!--<div class="grayback">OVERSEAS IMPORT</div>--></td></tr>
                
                <tr>
                <td align="center" class="display_yes">
                <!--<div class="display_yes" id="orders">
                <a style="width:148px" href="http://www1.overseaspr.com/soportes/index.html" title="Catálogo de Soportes" target="_self" class="display_yes">
				<= Lang("catalogo") %>
                </a><br>
                </div>--></td></tr>
                
              </table>
              
		 <br>
     <% If Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/cart.asp" or Request.ServerVariables("SCRIPT_NAME") = "/cart.asp" or len(Request("o")) > 0 Then %>
              <table width = "150px" border="0" cellpadding="0" cellspacing="0"  >
                <tr>
                  <td width = "150px" align="center" height='10'>
                  <%
				   if LEN(Session("MM_Multi_Username")) = 0 or (LEN(Session("MM_Multi_Username")) > 0 AND Session("MM_Multi_Username") <> "COUNTER") then %>
                  
                  <!--style="color:blue;background-color:transparent"-->
                  
                  <div class="grayback" id="order_list"  >
				  <%= Lang("ordenes") %>
                  </div>
                  
                  <% end if %>
                  </td>
                  
                </tr>
                <tr>
                   <td align="center">
                  
          <% if LEN(Session("MM_Multi_Username")) = 0 or (LEN(Session("MM_Multi_Username")) > 0 AND Session("MM_Multi_Username") <> "COUNTER") then %>
                   
          <div class="display_yes" width="200px" id="orders"   >
             <% OrdersDisplay() %>
          </div>
                   
                   <% end if %>
                   
                    <div class="display_no" id="loader2"><img src="images/ajax-loader.gif" width="16" height="16"></div>
                    
                    </td>
              </table>
              <% End if %>
              
              
              
            <td width="1043" align="left" valign="top">
              <form id="search_part" action="part_search.asp" method="post" >
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
                <table border="0" cellpadding="0" cellspacing="0" class="tablasFORMATO">
                <tr>
                  <td colspan="2" align="center" valign="top">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" align="center" valign="top">
                    <table width="946" border="0" cellpadding="0" cellspacing="0" >
                    <tr height="20">
                  	<td width="10">&nbsp;</td>
                    <td width="220" height="41" align="center">
                      <%= Lang("family_keyword") %>&nbsp;
                      <input type="text" class="ui-widget-content" id="fks"  onKeyUp="searchSel()" size="16"></td>
                    <td width="306" align="center">&nbsp;<%= Lang("family_category") %>&nbsp;
                    <select name="fcs" class="ui-widget-content" id="fcs" >
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
                    
                    <td width="166" align="center">
                      <%= Lang("modelo") %>&nbsp;<input name="model" type="text" class="ui-widget-content" id="fs2" size="16"></td>
                    
                    <td width="244" align="center">
                      
                      <input name="submit" type="submit" style="font-size:12px; font-weight:bold;width:140px" class="ui-widget" id="submit" value="<%=Lang("buscar")%>">
                      
                    </td>
                  </tr>
                  </table>
                  
                    
                    
                  </td>
                  </tr>
                  <tr>
                    <td width="82" align="center" valign="top">&nbsp;&nbsp;&nbsp;&nbsp;<table width=90 border="0" cellpadding="0" cellspacing="1">
                    <tr><td colspan="2" align="center"><%= Lang("entre_10_piezas") %></td>
                      </tr>
                        <tr>
                          <td width="17" height="25" align="left" valign="middle">&nbsp;1</td>
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
                    </table></td>
                    <td width="924" align="center" valign="top" style="background-color:transparent"  >
                    
                    <iframe src="cart_iframe.asp"  frameborder="0" width="924px" height="4980px"  style="overflow-x:hidden; overflow-y:hidden;background-color:transparent" allowtransparency="yes" >
                  </iframe>
                  
                  </td>
                  
                  <td width="12">&nbsp;</td>
                  
                  </tr>
                </table>
              </form>
            </td>
          </tr>
          <tr>
            <td height="12" colspan="2"><div class="footer"></div></td>
          </tr>
        </table>
      </td>
    </tr>
</table>
  
</body>
</html>
<%
dd_category.Close()
Set dd_category = Nothing
%>
