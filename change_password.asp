<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
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
</style>

<script language="jscript" type="application/javascript">        
function validate_form() {

	var pwd1	= document.forms[0].pwd1.value;
	var pwd2	= document.forms[0].pwd2.value
	
	if (isEmpty(pwd1))	{
			alert("Por favor indique su contraseña.");
			document.forms[0].pwd1.focus();
			return false;
	}
	
	if (isEmpty(pwd2))	{
			alert("Por favor confirme su contraseña.");
			document.forms[0].pwd2.focus();
			return false;
	}
	
	if (pwd1 !=	pwd2) {
			alert("La contraseña y la confirmación son diferentes por favor, verifíque nuevamente.");
			document.forms[0].pwd1.focus();
			return false;	
	}
		document.form1.submit();

}
function isEmpty(field){
	if (field.length < 1)
	{
		return true;
	}
	else {
		return false;
	}
}
</script>   
 
</head>
<%

validate_change_user = Session("MM_Username2")
'Response.Write( validate_change_user )
If validate_change_user = "" Then Response.Redirect("index.asp")

%>
<body>
<div class="container">

 	<div class="jumbotron">
	    <img src="images/oiclogo2.gif" alt="overseas_logo" width="256" height="31" align="top">
        <h3><%= Lang("cambiar_contrasena") %></h3>
        <h4><%= Lang("indique_nueva_contrasena") %></h4>
    </div>
     
    <form class="form-horizontal" action="change_password2.asp" method="POST" name="validacion" id="validacion" onsubmit="return validate_form();" >
    
	<div class="row">
    
		<div class="col-lg-6 col-xs-12">	
                         
            <div class="form-group">
                <label for="useridl" class="control-label col-xs-4 col-xs-offset-1"><%= Lang("contrasena") %></label>
                <div class="col-xs-6">
                     <input name="pwd1"  maxlength="10" type="password" class="form-control" id="pwd1" value="" style="width:120px" >
                </div>
            </div>
            
             <div class="form-group">
                <label for="passwordl" class="control-label col-xs-4 col-xs-offset-1"><%= Lang("confirmacion") %></label>
                <div class="col-xs-6">
                    <input name="pwd2"  maxlength="10" type="password" class="form-control" id="pwd2" value="" style="width:120px" >
                </div>
            </div>    
    	
         	<div class="form-group">
                <div class="col-xs-offset-5 col-xs-4">
                    <input style="width:120px" name="Entrar" type="submit" class="btn btn-primary" id="Entrar" value="Login" >
                </div>
            </div>
            
        </div>
                
	</div>
     <footer class="panel-footer col-xs-12">Overseas Import Corporation  All rights reserved TM 2016</footer>
    </form>    
         
                     
</div>
</body>
</html>
