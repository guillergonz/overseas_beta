<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
Response.Expires = -1

Set oRS = Server.CreateObject("ADODB.Recordset")
' Check websote on maintenance
strSQL = "SELECT forceinventoryupd FROM dbo.counters"
oRS.Open strSQL, MM_overseaspr_STRING
If Not oRS.EOF AND Not IsNull(oRS("forceinventoryupd")) Then
	onmaintenance = oRS("forceinventoryupd")	
Else
	onmaintenance = "0"
End If		
		
		
		
Session("MM_Multi_Username") = ""

' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = UCASE(CStr(Request.Form("usr")))
MM_valUserpwd = UCASE(CStr(Request.Form("pwd")))

If MM_valUsername <> "" Then

  	if Not MM_valUsername = "Z099" AND 	onmaintenance = "1" then
		Response.Redirect("index.asp?maint=1")
	end if		
		
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "security_proxy.asp"
  MM_redirectLoginFailed = "index.asp?e=1"

'  MM_loginSQL = "SELECT user_name, user_pwd"
  MM_loginSQL = "SELECT user_auto_id, user_pwd"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
'  MM_loginSQL = MM_loginSQL & " FROM dbo.users WHERE user_name = ? AND user_pwd = ?"
  MM_loginSQL = MM_loginSQL & " FROM dbo.users WHERE user_auto_id = ? AND user_pwd = ? AND activestatus = 'A' "
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_overseaspr_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 20, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 10, MM_valUserpwd) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
	
	
	if MM_valUsername = "B001" then
		Response.Redirect("multilogin.asp")
	else
		Response.Redirect(MM_redirectLoginSuccess)
	end if
    
	
	
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<%
Dim UserIPAddress
UserIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If UserIPAddress = "" Then
UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
End If
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
.input focus{
	border: #FC0;
}
.red {
	color:red;
}
body,td,th {
font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
font-size: 14px;
color: #000;
}
.form-signin {
	max-width:550px;
	padding:15px;
	margin:0px;
}

</style>

	
    
<script type="application/javascript">

//if($("#usr").val() == "B001") {
//	window.location.href = "multilogin.asp";
//}
			
$(document).ready(function() {

	$("#usr").val("");
	$("#pwd").val("");
	
	$("#effect").css("display", "inline");

	$("#usr").change(function() {
	
		var usr = $("#usr").val(); 
		if(usr == 'B001' || usr == 'B0011' || usr == 'B0012' || usr == 'B0013' || usr == 'B0014') {
			window.location.href = "multilogin.asp";
		}
		
		$("#Entrar").show();
		$("#accessdenied").hide();
		var usr = $("#usr").val();
		var ip = $("#ip").val();
		//if(usr.toUpperCase() == "Z099" && ip.substring(0,7) != '162.222') {
		//	$("#usr").val('');
		//	$("#Entrar").hide();
		//	$("#accessdenied").show();
		//	return false;
		//}
		
	});
	
	
		
});
</script>
</head>

<body  >

<div class="container"  >
	
    <div class="jumbotron" style="padding-bottom:4px;padding-top:4px;margin-bottom:0px" >
	    <img src="images/oiclogo2.gif" alt="overseas_logo" width="256px" height="31px" align="top">
        <h4>Welcome<br>Looking for wholesale auto parts? &nbsp;We specialize in spare parts for Japanese and Korean cars. We offer a large selection of wholesale car parts through this easy to use website.&nbsp;Find wholesale parts at the lowest prices ...</h4>
    </div>
    
  	<div style="width:800px;float:right;margin:3px">
      <!-- Begin DWUser_EasyRotator -->
      <script type="text/javascript" src="http://c520866.r66.cf2.rackcdn.com/1/js/easy_rotator.min.js"></script>
      <div class="dwuserEasyRotator" style="width: 700px; height: 400px; position:relative; text-align: left;" data-erconfig="{autoplayEnabled:true, lpp:'102-105-108-101-58-47-47-47-67-58-47-85-115-101-114-115-47-71-71-71-47-68-111-99-117-109-101-110-116-115-47-69-97-115-121-82-111-116-97-116-111-114-80-114-101-118-105-101-119-47-112-114-101-118-105-101-119-95-115-119-102-115-47', wv:1}" data-ername="oicslide" data-ertid="{k3f88ghj1677019407351}">
        <div data-ertype="content" style="display: none;">
          <ul data-erlabel="Main Category">
          <li> <img class="main" src="easyrotatorimages/1ST.png" /> <img class="thumb" src="easyrotatorimages/1ST.png" /> </li>
           <li> <img class="main" src="easyrotatorimages/2ND.png" /> <img class="thumb" src="easyrotatorimages/2ND.png" /> </li>
            <li> <img class="main" src="easyrotatorimages/3RD.png" /> <img class="thumb" src="easyrotatorimages/3RD.png" /> </li>
           <li> <img class="main" src="easyrotatorimages/4RTH.png" /> <img class="thumb" src="easyrotatorimages/4RTH.png" /> </li>
           <li> <img class="main" src="easyrotatorimages/5TH.png" /> <img class="thumb" src="easyrotatorimages/5TH.png" /> </li>
           <li> <img class="main" src="easyrotatorimages/6TH.png" /> <img class="thumb" src="easyrotatorimages/6TH.png" /> </li>
           <li> <img class="main" src="easyrotatorimages/7TH.png" /> <img class="thumb" src="easyrotatorimages/7TH.png" /> </li>
           <li> <img class="main" src="easyrotatorimages/8TH.png" /> <img class="thumb" src="easyrotatorimages/8TH.png" /> </li>
            <li> <img class="main" src="easyrotatorimages/9TH.png" /> <img class="thumb" src="easyrotatorimages/9TH.png" /> </li>
            <li> <img class="main" src="easyrotatorimages/10TH.png" /> <img class="thumb" src="easyrotatorimages/10TH.png" /> </li>
          </ul>
        </div>
        <div data-ertype="layout" data-ertemplatename="NONE" style="">
          <div class="erimgMain" style="position:absolute; left:0;right:0;top:0;bottom:0;" data-erconfig="{___numTiles:3, scaleMode:'scaleDown', duration:400, imgType:'main', __loopNextButton:false, __arrowButtonMode:'rollover'}">
            <div class="erimgMain_slides" style="position: absolute; left:0; top:0; bottom:0; right:0;">
              <div class="erimgMain_slide">
                <div class="erimgMain_img" style="position: absolute; left: 0; right: 0; top: 0; bottom: 59px;"></div>
                <div class="erimgMain_title" style="position:absolute; left: 0; right: 0; bottom: 39px; height: 15px; text-align:center; color: #000; font-size:11px; font-family:'Courier New',Courier,monospace;"></div>
              </div>
            </div>
            <div class="erimgMain_arrowLeft" style="position:absolute; left: 50%; bottom:15px; margin-left: -30px;" data-erconfig="{image:'http://easyrotator.s3.amazonaws.com/1/i/rotator/custom_takuna/left_arrow_export.png', image2:'http://easyrotator.s3.amazonaws.com/1/i/rotator/custom_takuna/left_arrow_export.png'}"></div>
            <div class="erimgMain_arrowRight" style="position:absolute; right: 50%; bottom:15px; margin-right: -30px;" data-erconfig="{image:'http://easyrotator.s3.amazonaws.com/1/i/rotator/custom_takuna/right_arrow_export.png', image2:'http://easyrotator.s3.amazonaws.com/1/i/rotator/custom_takuna/right_arrow_export.png'}"></div>
          </div>
          <div class="erdynamicText" data-erfield="none" data-erconfig="{fadeDur:0, hideWhenEmpty:true}" style="position:absolute; left: 50%; bottom: 10px; height: 20px; margin-left: -20px; width: 40px; text-align:center; color: #000; font:11px 'Courier New',Courier,monospace;"> {image.index}/{image.total} </div>
          <div class="erabout erFixCSS3" style="color: #FFF; text-align: left; background: #000; background:rgba(0,0,0,0.93); border: 2px solid #FFF; padding: 20px; font: normal 11px/14px Verdana,_sans; width: 300px; border-radius: 10px; display:none;"> This <a style="color:#FFF;" href="http://www.dwuser.com/easyrotator/" target="_blank">jQuery slider</a> was created with the free <a style="color:#FFF;" href="http://www.dwuser.com/easyrotator/" target="_blank">EasyRotator</a> software from DWUser.com. <br />
            <br />
            Use WordPress? The free <a style="color:#FFF;" href="http://www.dwuser.com/easyrotator/wordpress/" target="_blank">EasyRotator for WordPress</a> plugin lets you create beautiful <a style="color:#FFF;" href="http://www.dwuser.com/easyrotator/wordpress/" target="_blank">WordPress sliders</a> in seconds. <br />
            <br />
            <a style="color:#FFF;" href="#" class="erabout_ok">OK</a> </div>
          <noscript>
            Rotator powered by <a href="http://www.dwuser.com/easyrotator/">EasyRotator</a>, a free and easy jQuery slider builder from DWUser.com.  Please enable JavaScript to view.
          </noscript>
          <script type="text/javascript">/*Avoid IE gzip bug*/(function(b,c,d){try{if(!b[d]){b[d]="temp";var a=c.createElement("script");a.type="text/javascript";a.src="http://easyrotator.s3.amazonaws.com/1/js/nozip/easy_rotator.min.js";c.getElementsByTagName("head")[0].appendChild(a)}}catch(e){alert("EasyRotator fail; contact support.")}})(window,document,"er_$144");</script>
        </div>
      </div>
      <!-- End DWUser_EasyRotator -->
    </div>
  
  <!-- <img class="hidden-xs hidden-sm hidden-md pull-right" style="margin-top:50px;margin-right:20px" src="logooic.png" width="568" height="322" alt="logo3">   -->
    
    <form class="form-signin" action="<%=MM_LoginAction%>" method="POST" name="validacion" id="validacion" style="width:350px;float:left"  >
    
    <h3 class="form-signin-heading">Please sign in</h3>
    
    
    <label for="useridl" class="sr-only">User ID</label>
    <input name="usr" type="text" maxlength="10" class="form-control" id="usr" placeholder="Enter userid" required autofocus >
    <br>

    <label for="pwd" class="sr-only">Password</label>
    <input placeholder="Enter Password" name="pwd"  maxlength="10" type="password" class="form-control" id="pwd" value="" required  >
	<br>
    
	<%
    Time1 = TimeValue("9:00:00 PM")
    Time2 = TimeValue("9:15:00 PM") 
    
	
	
    TimeCheck = TimeValue(Now)
    If (Time1 < TimeCheck AND TimeCheck < Time2) Then
    %>
                
	<label class="red">Data loading<br>Please wait...</label>
                        
    <% Else %>
                
     <label id="accessdenied" style="display:none;" class="red">Acceso denegado. Esta página es para uso exclusivo de clientes de Overseas Corp.<br>Estamos monitoreando su cuenta.</label>
	 
     <input style="width:150px" name="Entrar" type="submit" class="btn btn-primary btn-block" id="Entrar" value="Login" >
	 <br>
	 <input type="hidden" value="<%=Request.ServerVariables("REMOTE_ADDR")%>" >
	 <input type="hidden" id="ip" value="<%=UserIPAddress%>" >
      

	  
    <% End If %>
            
	<% 					
    If Request("e") > 0 Then 
        If Request("e") = 1 Then
            Response.Write "<label class='red'>INVALID PASSWORD!</label>"
        End if
    End if
    
     If Request("maint") > 0 Then 
        If Request("maint") = 1 Then
            Response.Write "<label class='red'>Periodo de Mantenimiento<br>Website on maintenance</label>"
        End if
    End if
    %>
    <br>
    </form>     
	
    <br>
    <footer class="panel-footer col-xs-12">
        <p>Contact us:(T) 787-751-4036&nbsp;(F) 787-765-6735	&nbsp;Correo Electrónico/Email: oic@overseaspr.com<br>
        Physical Address: Urb. El Paraiso Calle Ganges #9 San Juan P.R. 00926</p>
    </footer>
                    
    <footer class="panel-footer col-xs-12">Overseas Import Corporation  All rights reserved TM 2016</footer>
       
         
</div>
</body>
</html>
