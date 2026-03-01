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
<!doctype html>
<html>
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
.input focus{
	border: #FC0;
}
body,td,th {
font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
font-size: 14px;
color: #000;
}
.red {
  color: #d14;
}
textarea { 
    resize: none; 
}
</style>

<script type="text/javascript" charset="utf-8">
	$(document).ready(function () {

		// disable enter submitting form'
		$(document).keypress(function (e) {
			if (e.which == 13) {
				return false;
			}
		});

		
		
	});
</script>


<script type="text/javascript" charset="utf-8">

function send_onclick() {
	var deltype
	deltype = $('#delivery_type').val();
	
	if (deltype == 'Q') {
	
		alert("You must enter a delivery method ( Delivery or Pickup ) ");
		return false
		
	}
	else {
	
		var d = new Date(); // for now
		//d.getHours(); // => 9
		//d.getMinutes(); // =>  30
		//d.getSeconds(); // => 51
		if (d.getHours() == 9 && d.getMinutes() >= 51 && deltype == 'S') {
			deltype = 'D';
		}
		document.getElementById("order_submit").submit();
		return true
		
	}
};
		
function UpdCart(str){ 	

	
	var oshop_quantity	= "shop_quantity" + (str);
	var xshop_quantity	= document.getElementById(oshop_quantity).value; 
	
	var src="upd_quantity_shopping_cart.asp?p=" + str + "&t=2&q=" + xshop_quantity;
	
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
function DelFromCart(str) { 

	var src="del_p_shopping_cart.asp?p=" + str + "&t=2&sid="+Math.random();	
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
function stripIt(x){
	
	// con esto acabamos el problema
	x.value = encodeURIComponent(x.value)
	
	// / /g global
	x.value = x.value.replace(/%C3%A1/g,'a');
	x.value = x.value.replace(/%C3%A9/g,'e');
	x.value = x.value.replace(/%C3%AD/g,'i');
	x.value = x.value.replace(/%C3%B3/g,'o');
	x.value = x.value.replace(/%C3%BA/g,'u');
	
	x.value = x.value.replace(/%C3%B1/g,'n');
	
	x.value = x.value.replace(/['".,]/g,'');	
	
	x.value = x.value.replace(/%20/g,' ');
	x.value = x.value.replace(/%0A/g,'');
	
};
</script>

</head>
<body>
 
<% If PartsOnCart() Then %>

<form  id="order_submit" method="post" action="/cart_proceed.asp" >
<div class="container">
	
    
    		<!--<label class="alert alert-danger" >Atención: El próximo día de entrega es el MARTES 27 de Diciembre.</label>-->
            
            
           
            <!--style="margin:10px"-->
            <div class="display_yes" id="shopping_cart"  >
            	<h4 class="red"><%= Lang("verificar_cart") %></h4>
                <% CartDisplay("Y") %>
            </div>
			
            <div class="display_no" id="loader"><img src="images/ajax-loader.gif" alt="loader" width="16" height="16"></div>
			

            <!--<label class="alert alert-danger" >Attention: Next delivery day will be TUESDAY December 27 2016.</label>-->

            
            <%
			
			tlimit = FormatDateTime("9:50", vbShortTime)
			tnow = FormatDateTime(Now, vbShortTime)
			
			tlimit2 = FormatDateTime("17:00", vbShortTime)
			
			%>    
            
            <% if False and (Session("lang") = "S") then %>
            <div class="row">
            <img src="images/rutas_same_day.png" alt="rutas" width="770" height="456" longdesc="http://www.overseaspr.com">
            <br><br>
            </div>
            <% end if %>
            
			<hr>
            <div class="row">
                <div class="col-xs-6">
					<label><%= Lang("metodo_envio") %></label>
                	<br>
                    <select name="delivery_type" class="form-control" id="delivery_type"  >
                    <option class="alert-info" selected="selected" value="Q"></option>
                    
                    
                    <option class="alert-info" value="D"><%= Lang("entrega") %></option>
                    
                    
                    <% if False and ((tnow < tlimit or tnow > tlimit2) OR weekday(Now)=7 OR weekday(Now)=1) AND (Session("lang") = "S") then %>
                    
                    <% if False and weekday(Now)=7 OR weekday(Now)=1 then %>
                    <option class="alert-info" value="S"><%= Lang("entregamismodia") %></option>
                    <% else %>
                    <option class="alert-info" value="S"><%= Lang("entregamismodiaSD") %></option>	
					<% end if %>
                    <% end if %>
                    
                    <option class="alert-info" value="P"><%= Lang("recoger") %></option> 
                    </select>
                </div>
				
				<div class="col-xs-4">
					<label><%= Lang("Instrucciones") %>&nbsp;(25)</label>
               		<br>
               		<input type="text"  name="instructions" class="form-control" id="instructions" onBlur="stripIt(this);" maxlength="25" >
				</div>
				<div class="col-xs-2">   
               		<input name="submitform" type="submit" class="btn btn-danger" style="margin-top:24px" id="submitform" value="<%=Lang("someter_orden")%>"  onClick="return send_onclick()" >
               </div>
           </div>
           
		   <hr>
      
		  <div class="row" id="credito" style="background-color:#FFF; margin-top:10px">
		   
			 <div >
				<h3 class="label label-warning">Métodos de pago para COD solamente.</h3>
				<h3 class="label label-warning">Para mas informacion comunicarse al 787 751-4036 Ext 15</h3>
				<img  src="/images/visamasteramex.jpg" width="483" height="104" alt="credito">
				<img  src="/images/athbusiness.png" width="211" height="104" alt="ath">
			 </div> 
			 
		   </div>      




		   
	
</div>
</form>

<% End if %>

</body>
</html>
