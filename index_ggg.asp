
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="Connections/overseaspr.asp" -->

<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = UCASE(CStr(Request.Form("usr")))
MM_valUserpwd = UCASE(CStr(Request.Form("pwd")))

If MM_valUsername <> "" Then
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
  MM_loginSQL = MM_loginSQL & " FROM dbo.users WHERE user_auto_id = ? AND user_pwd = ?"
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
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>



<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_overseaspr_STRING
Recordset2_cmd.CommandText = "SELECT titulo1,mensaje1,titulo2,mensaje2,titulo3,mensaje3 FROM dbo.especiales WHERE uniqueidcol = 1 " 
Recordset2_cmd.Prepared = true
Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>


<!DOCTYPE HTML><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Overseas Import Corporation - Login</title>

<link rel="icon" href="images/favicon.ico" type="image/x-icon" /> 
<link rel="icon" href="images/favicon.ico" type="image/x-icon" />

<!--[if lt IE 9]>
<script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
<![endif]-->

<!--

<link rel="stylesheet" href="theme/jquery.ui.all.css">
	<script src="jquery-1.6.2.js"></script>
	<script src="ui/jquery.ui.core.js"></script>
	<script src="ui/jquery.ui.widget.js"></script>
	<script src="ui/jquery.ui.button.js"></script>
    
-->

<script type="text/javascript" src="jquery-ui/jquery-ui-1.8.16-sunny.custom/js/jquery-1.6.2.min.js"></script>
<script type="text/javascript" src="jquery-ui/jquery-ui-1.8.16-sunny.custom/js/jquery-ui-1.8.16.custom.min.js"></script>
<link type="text/css" href="jquery-ui/jquery-ui-1.8.16-sunny.custom/css/sunny/jquery-ui-1.8.16.custom.css" rel="stylesheet" />
	
<script src="jquery-ui/jquery-ui-1.8.16-sunny.custom/development-bundle/ui/jquery.ui.core.js"></script>
<script src="jquery-ui/jquery-ui-1.8.16-sunny.custom/development-bundle/ui/jquery.ui.widget.js"></script>
<script src="jquery-ui/jquery-ui-1.8.16-sunny.custom/development-bundle/ui/jquery.ui.datepicker.js"></script>
        

<link href="overseas.css" rel="stylesheet" type="text/css">

<style type="text/css">

body {
	background-color: #000;
}
.footer {
	font-size: 10px;
	text-align: center;
	padding: 10px;
}
.wrapper {
	display: block;
	width: 1050px;
	margin-left: auto;
	margin-right: auto;
}

.startsUgly { display: none; }

</style>


<script type="text/javascript" charset="utf-8">
$(document).ready(function() {
	
	$("#Entrar").button();
	
	$("input:submit, input[type=button], a, button" ).button();
	
	$(".startsUgly").show();
	
});
</script>


<noscript>
  <style type="text/css">.startsUgly { display: block; }</style>
</noscript>
    
</head>

<body>
<div class="startsUgly">
  <div class="wrapper" >
    
    <table width="97%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="15">&nbsp;</td>
        <td>&nbsp;</td>
        <td align="center" valign="middle"  >&nbsp;</td>
        <td align="center" valign="middle"  >&nbsp;</td>
      </tr>
      <tr>
        <td width="28%" height="474" rowspan="2">
        <form action="<%=MM_LoginAction%>" method="POST" name="validacion" id="validacion" class="borderaround_login" style="width:95%" >
          <br>
          <center>
            <img src="images/oiclogo2.gif" alt="overseas_logo" width="256" height="31" align="top">
            </center>
          <br>
          <center>
            <label>Looking for wholesale auto part? &nbsp;&nbsp;We specialize in spare parts for Japanese and Korean cars. &nbsp;&nbsp;We offer a large selection of wholesale car parts through this easy to use website.&nbsp;&nbsp;Find wholesale parts at the lowest prices ...</label>
            </center>
          <table cellpadding="0" cellspacing="0" border="0" width="90%">
            <tr>
              <td align="center" valign="middle">&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td width="60%" align="right" valign="middle">Enter a valid user id:&nbsp;</td>
              <td width="40%"><input name="usr" type="text" size="10" maxlength="10" class="login_fields" id="usr" />
                </td>
              </tr>
            <tr>
              <td width="60%" align="right" valign="middle" >
              Enter a valid password:&nbsp;
              </td>
              <td><input name="pwd" size="10"  maxlength="10" type="password" class="login_fields" id="pwd" />
                </td>
              </tr>
            </table>
          
          
          <table cellpadding="0" cellspacing="0" border="0" width="90%">
            <tr>
              <td width="60%" height="41" align="center" valign="middle"><label style="color:yellow">
                <% If Request("e") = 1 Then
           Response.Write "&nbsp;&nbsp;INVALID PASSWORD!&nbsp;&nbsp;"
       End if
    %>
                </label></td>
              <td width="40%">
                <input name="Entrar" type="submit" id="Entrar" value="Entrar / Login" ></td>
              </tr>
            </table>
            <br><br>
          <center>
            <label style="color:yellow">Contact us&nbsp;:&nbsp;(T) 787-751-4036<br>
              &nbsp;(F) 787-765-6735</label>
            <br>
            Correo Electrónico/Email: oic@overseaspr.com <br>
             Address: Urb. El Paraiso
            Calle Ganges #9 <br>
            San Juan P.R. 00926
            </center>
          </p>
          <div class="footer">
            Overseas Import Corporation<br>
            All rights reserved TM 2011
            </div>
          
        </form></td>
        <td width="0%" rowspan="2">&nbsp;</td>
        <td width="3%" align="center" valign="middle"  >&nbsp;</td>
        <td width="69%" rowspan="3" align="center" valign="middle"  >
          
          
          
          
          
          
          
          
          
  









































































<!-- Begin DWUser_EasyRotator -->
        <script type="text/javascript" src="http://c520866.r66.cf2.rackcdn.com/1/js/easy_rotator.min.js"></script>
        <div class="dwuserEasyRotator" style="width: 100%; height: 400px; position:relative; text-align: left; vertical-align:top" data-erconfig="{autoplayEnabled:true, lpp:'102-105-108-101-58-47-47-47-67-58-47-85-115-101-114-115-47-103-103-111-110-122-97-108-101-122-47-68-111-99-117-109-101-110-116-115-47-69-97-115-121-82-111-116-97-116-111-114-80-114-101-118-105-101-119-47-112-114-101-118-105-101-119-95-115-119-102-115-47', wv:1, randomize:true, autoplayDelay:4000}" data-ername="Overseas Import Corporation" data-erTID="{pvjdyzj1gt8384210424631}" data-erAudioConfig="{autoplay:false}"   data-erResponsiveRatio="{16:9}">
          <div data-ertype="content" style="display: none;"><ul data-erlabel="Especiales semanales">
	<li>
		<a class="mainLink" href="parts_images/Especial1.jpg" target="_blank" title="Especial1"><img class="main" src="parts_images/Especial1.jpg" /></a>
		<img class="thumb" src="parts_images/Especial1.jpg" />
		<span class="desc"><%= recordset2("mensaje1") %></span>
	</li>
	<li>
		<a class="mainLink" href="link2" target="_blank" title="Especial2"><img class="main" src="parts_images/Especial2.jpg" /></a>
		<img class="thumb" src="parts_images/Especial2.jpg" />
		<span class="desc"><%= recordset2("mensaje2") %></span>
	</li>
	<li>
		<a class="mainLink" href="link3" target="_blank" title="Especial3"><img class="main" src="parts_images/Especial3.jpg" /></a>
		<img class="thumb" src="parts_images/Especial3.jpg" />
		<span class="desc"><%= recordset2("mensaje3") %></span>
	</li>
</ul>
</div>
          <div data-ertype="layout" data-ertemplatename="NONE" style="">			<div class="erimgMain" style="position: absolute; left:0;right:0;top:0;bottom:0;" data-erConfig="{___numTiles:3, scaleMode:'fillArea', duration:800, imgType:'main', alwaysPreviousButton:true, __loopNextButton:false, __arrowButtonMode:'rollover'}">
				<div class="erimgMain_slides" style="position: absolute; left:0; top:0; bottom:0; right:0;">
					<div class="erimgMain_slide">
						<div class="erimgMain_img" style="position: absolute; left: 0; right: 0; top:0;bottom:0;"></div>
					</div>
				</div>
				<!-- <div class="erimgMain_arrowLeft" style="position:absolute; left: 10px; top: 50%; margin-top: -15px;" data-erConfig="{image:'circleSmall', image2:'circleSmall'}"></div>
				<div class="erimgMain_arrowRight" style="position:absolute; right: 10px; top: 50%; margin-top: -15px;"></div> -->
			</div>
			<div class="" style="position: absolute; left:0; right:0; bottom: 20px; padding: 7px 200px 7px 20px; background: #000; background:rgba(0,0,0,0.9); color: #FFF; font-family: Georgia, 'Times New Roman', Times, _serif; text-align: left;">
				<p class="erdynamicText" data-erfield="title" style="padding: 0; margin: 0 0 3px 0; font-weight: bold; font-size: 22px; color: #FFF;"></p>
				<p class="erdynamicText" data-erfield="desc" style="padding: 0; margin: 0; font: 12px/16px Arial,_sans; color: #FFF;"></p>
			</div>
			<div class="erdots" style="overflow: hidden; margin: 0; font-size: 10px; font-family: 'Lucida Grande', 'Lucida Sans', Arial, _sans; color: #FFF; position: absolute; right:6px; bottom:30px; width:200px;" data-erConfig="{showText:false}" align="center">
				<div class="erdots_wrap" style="wasbackground-color: #CFC; float: right;" align="left"> <!-- modify the float on this element to make left/right/none=center aligned. -->
					<span class="erdots_btn_selected" style="padding-left: 0; width: 21px; height: 20px; display: inline-block; text-align: center; vertical-align: middle; line-height: 20px; margin: 0 2px 0 0; cursor: default; background: url(http://easyrotator.s3.amazonaws.com/1/i/rotator/dots/export/20_14_wite_65.png) top left no-repeat;">
						&nbsp;
					</span>
					<span class="erdots_btn_normal" style="padding-left: 0; width: 21px; height: 20px; display: inline-block; text-align: center; vertical-align: middle; line-height: 20px; margin: 0 2px 0 0; cursor: pointer; background: url(http://easyrotator.s3.amazonaws.com/1/i/rotator/dots/export/20_14_wite_35.png) top left no-repeat;">
						&nbsp;
					</span>
					<span class="erdots_btn_hover" style="padding-left: 0; width: 21px; height: 20px; display: inline-block; text-align: center; vertical-align: middle; line-height: 20px; margin: 0 2px 0 0; cursor: pointer; background: url(http://easyrotator.s3.amazonaws.com/1/i/rotator/dots/export/20_14_wite_65.png) top left no-repeat;">
						&nbsp;
					</span>
				</div>
			</div><div class="erabout erFixCSS3" style="color: #FFF; text-align: left; background: #000; background:rgba(0,0,0,0.93); border: 2px solid #FFF; padding: 20px; font: normal 11px/14px Verdana,_sans; width: 300px; border-radius: 10px; display:none;"> This <a style="color:#FFF;" href="http://www.dwuser.com/easyrotator/" target="_blank">jQuery slider</a> was created with the free <a style="color:#FFF;" href="http://www.dwuser.com/easyrotator/" target="_blank">EasyRotator</a> software from DWUser.com. <br />
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










































































          
          
          
          
          
          
          
          
          
  </td>
      </tr>
      <tr>
        <td align="center" valign="middle"  >&nbsp;</td>
      </tr>
      <tr>
        <td align="center" valign="middle">  
          <param name="movie" value="http://www.youtube.com/v/4TshFWSsrn8?version=3">
          <param name="allowFullScreen" value="true">
          <param name="allowScriptAccess" value="always">
          <embed src="http://www.youtube.com/v/4TshFWSsrn8?version=3" type="application/x-shockwave-flash" allowfullscreen="true" allowscriptaccess="always" width="219" height="202">
        </object></td>
        <td align="center" valign="middle">&nbsp;</td>
        <td width="3%" align="center" valign="middle"  >&nbsp;</td>
      </tr>
    </table>
    <br clear="all">
    
    
  </div>
</div>
    

</body>
</html>

<%
Recordset2.Close()
Set Recordset2 = Nothing
%>