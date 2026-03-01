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
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation - Order Detail</title>

<link rel="icon" href="/images/favicon.ico" type="image/x-icon"> 
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">

<link href="overseas.css" rel="stylesheet" type="text/css">


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
	font-family: "Segoe UI", "Segoe UI Light", "Segoe UI Semibold", "Lucida Console";
	font-size: 12px;
	color: #000;
}

</style>


<script type="text/javascript">
function send_onclick() {
	 
	var bolSubmit
	var deltype
	bolSubmit = true;
	
	// if ( order_submit.delivery_type.value == "" ) { 
	deltype = $('#delivery_type').val()
	if ( deltype == '' ) {
		alert("You must enter a delivery method ( Delivery or Pickup ) ");
		return false
	}
	{
		document.getElementById("order_submit").submit();
		return true
	}
   
}
</script>


<script>
	$(function() {
		$( "input:submit, a, button, #dev_type " ).button();
		//$( "a, #dev_type" ).click(function() { return false; });
	});
</script>

<script>
var xmlHttp3
						
function UpdCart(str){ 
	xmlHttp3=GetXmlHttpObject3()
	if (xmlHttp3==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	
//	$("#shopping_cart").className = "display_no" ;
	$("#loader").className = "display_yes" ;
	
	var oshop_quantity	= "shop_quantity" + (str);
	var xshop_quantity	= document.getElementById(oshop_quantity).value; 


//	qty2 = document.getElementById("shop_quantity").value
	
	//document.getElementById("shopping_cart").className = "display_no" ;
	//document.getElementById("loader").className = "display_yes" ;

	var url3="upd_quantity_shopping_cart.asp"
	url3=url3+"?p=" + str
	//url2=url2+"&sid="+Math.random()
	url3=url3+"&t=2"
	url3=url3+"&q=" + xshop_quantity
	xmlHttp3.onreadystatechange=stateChanged3 
	xmlHttp3.open("GET",url3,true)
	xmlHttp3.send(null)
	
	parent.window.location = "cart.asp"
}
	
function stateChanged3() { 
	if (xmlHttp3.readyState==4 || xmlHttp3.readyState=="complete"){ 
	
		document.getElementById("shopping_cart").innerHTML=xmlHttp3.responseText 

//		$("#shopping_cart").className = "display_yes" ;
		$("#loader").className = "display_no" ;

		$( "input:submit, a, button, .demo, #DelFromcart, #UpdCart" ).button();

		//document.getElementById("shopping_cart").className = "display_yes" ;
		//document.getElementById("loader").className = "display_no" ;
	} 
} 

function GetXmlHttpObject3(){ 
	var objXMLHttp3=null
	if (window.XMLHttpRequest){
		objXMLHttp3=new XMLHttpRequest()
	}
	else if (window.ActiveXObject){
		objXMLHttp3=new ActiveXObject("Microsoft.XMLHTTP")
	}

	return objXMLHttp3
}


var xmlHttp2

function DelFromCart(str)
{ 
xmlHttp2=GetXmlHttpObject2()
if (xmlHttp2==null){
	alert ("Browser does not support HTTP Request")
	return
} 
//$("#txtHint").className = "display_no" ;
//document.getElementById("shopping_cart").className = "display_no" ;
document.getElementById("loader").className = "display_yes" ;

var url2="del_p_shopping_cart.asp"
url2=url2+"?p=" + str
//url2=url2+"&sid="+Math.random()
url2=url2+"&t=2"
xmlHttp2.onreadystatechange=stateChanged2 
xmlHttp2.open("GET",url2,true)
xmlHttp2.send(null)

parent.window.location = "cart.asp"

}

function stateChanged2(){ 
	if (xmlHttp2.readyState==4 || xmlHttp2.readyState=="complete")
{ 
	document.getElementById("shopping_cart").innerHTML=xmlHttp2.responseText 
	//document.getElementById("shopping_cart").className = "display_yes" ;
	document.getElementById("loader").className = "display_no" ;
	
	$( "input:submit, a, button, .demo, #DelFromcart, #UpdCart" ).button();
	
	} 
} 

function GetXmlHttpObject2(){ 
	var objXMLHttp2=null
	if (window.XMLHttpRequest){
		objXMLHttp2=new XMLHttpRequest()
	}
	else if (window.ActiveXObject)
	{
	objXMLHttp2=new ActiveXObject("Microsoft.XMLHTTP")
	}
	return objXMLHttp2
}
</script>


<script type='text/javascript'>
 function stripIt(x){x.value = x.value.replace(/['".,]/g,'');};
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



 
 
<% If PartsOnCart() Then %>

<div class="borderaround_white" style="margin:25px">
	<form id="order_submit" method="post" action="/cart_proceed.asp" >
    
    
            <br>
            <div class="display_yes" id="shopping_cart" style="margin:10px" >
                <% CartDisplay("Y") %>
            </div>

			<hr>

            <h2 class="ui-state-highlight"><%= Lang("verificar_cart") %></h2>

            <h3><%= Lang("orden_procesada") %></h3>

		<table width="100%" border="0" cellpadding="0" cellspacing="0" >
            <tr>
            <td align="center" >
			<div class="display_no" id="loader"><img src="images/ajax-loader.gif" width="16"></div>
            <div class="ui-widget-header" style="vertical-align:bottom;background-image:none">
                
                <table border="0" cellpadding="0" cellspacing="0" >
                <tr>    
                <td>
				<%= Lang("metodo_envio") %>
                </td>
                <td>
                <select id="delivery_type" name="delivery_type" class="ui-state-highlight" align="center"  style="font-size:14px;width:100px;margin-left:20px" >
               <option selected="selected" value="">
               <option value="D"><%= Lang("entrega") %></option>
               <option value="P"><%= Lang("recoger") %></option>
               </select>
                </td>
                <td>
               &nbsp;&nbsp;<%= Lang("Instrucciones") %>&nbsp;&nbsp;</td>
               <td>
               <textarea name="instructions" cols="25" rows="4" class="ui-state-highlight" id="instructions" onBlur="stripIt(this);" maxlength="100" ></textarea>
               </td>
               <td>	        
               <input name="submitform" type="submit" style="font-size:12px; font-weight:bold;width:140px;margin-left:20px" class="ui-widget" id="submitform" value="<%=Lang("someter_orden")%>"  onClick="return send_onclick()" >
               </td>
               </<br>
     		   </tr>
               </table>
                	          
            </div>
            <br>
            <hr>
            <br>
                    
            </td>
			</tr>
   		</table>
	</form>
</div>

<% End if %>

</body>
</html>
