<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Overseas Import Corporation - Shopping Cart</title>
<link href="overseas.css" rel="stylesheet" type="text/css" />
</head>

<script>
var xmlHttp2

function DelFromCart(str)
{ 
xmlHttp2=GetXmlHttpObject()
if (xmlHttp2==null)
{
alert ("Browser does not support HTTP Request")
return
} 
//$("#txtHint").className = "display_no" ;
document.getElementById("shopping_cart").className = "display_no" ;
document.getElementById("loader").className = "display_yes" ;

var url2="del_to_shopping_cart.asp"
url2=url2+"?p=" + str
url2=url2+"&sid="+Math.random()
xmlHttp2.onreadystatechange=stateChanged2 
xmlHttp2.open("GET",url2,true)
xmlHttp2.send(null)
}

function stateChanged2() 
{ 
if (xmlHttp2.readyState==4 || xmlHttp2.readyState=="complete")
{ 
document.getElementById("shopping_cart").innerHTML=xmlHttp.responseText 
document.getElementById("shopping_cart").className = "display_yes" ;
document.getElementById("loader").className = "display_no" ;
} 
} 

function GetXmlHttpObject2()
{ 
var objXMLHttp2=null
if (window.XMLHttpRequest)
{
objXMLHttp2=new XMLHttpRequest()
}
else if (window.ActiveXObject)
{
objXMLHttp2=new ActiveXObject("Microsoft.XMLHTTP")
}
return objXMLHttp2
}
</script>

<script>
function ValidateCart(part,amt) {
	amount = document.getElementById(amt).value
	AddToCart(part,amount);
}
</script>                

<body>
<table width="200" border="0" cellpadding="0" cellspacing="0" class="tablas">
  <tr>
    <td width="47%" height="20" align="center" bgcolor="#CCCCCC"><strong>PIEZA</strong></td>
    <td width="16%" align="center" bgcolor="#CCCCCC"><strong>CANT.</strong></td>
    <td width="26%" align="center" bgcolor="#CCCCCC"><strong>PRECIO</strong></td>
    <td width="11%" align="center" bgcolor="#CCCCCC">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
  </tr>
</table>
</body>
</html>
