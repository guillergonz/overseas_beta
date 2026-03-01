<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
<script src="PrototypeFishEye/includes/prototype.js" type="text/javascript"></script>
<script src="PrototypeFishEye/includes/effects.js" type="text/javascript"></script>
<script src="PrototypeFishEye/includes/EventDispatcher.js" type="text/javascript"></script>
<script src="PrototypeFishEye/includes/FishEye.js" type="text/javascript"></script>
<style type="text/css">
/* BeginOAWidget_Instance_2146022: #PrototypeFishEye */

		#PrototypeFishEye.fishEye{
			height:45px;
			border:1px solid #999999;
			padding:10px;
			text-align:center;
			background-color:#cccccc;
			margin-top:30px;
			margin-bottom:0px;
		}
		#PrototypeFishEye.fishEye img{
			width:50px;
			height:50px;
			margin-top:0px;
			margin-bottom:0px;
		}
	
/* EndOAWidget_Instance_2146022 */
</style>
<script type="text/xml">
<!--
<oa:widgets>
  <oa:widget wid="2146022" binding="#PrototypeFishEye" />
</oa:widgets>
-->
</script>
</head>

<body>
<div id="PrototypeFishEye" class="fishEye"> <img src="/soportes/PrototypeFishEye/images/My-PC.png" href="#" title="Products" alt="Products" /> <img src="/soportes/PrototypeFishEye/images/My-Documents.png" href="#" alt="Projects" title="Projects"/> <img src="/soportes/PrototypeFishEye/images/My-Music.png" href="#" title="Blog" alt="Blog"/> <img src="/soportes/PrototypeFishEye/images/Search.png"  href="#" title="Blog Search" alt="Blog Search"/> <img src="/soportes/PrototypeFishEye/images/My-Network.png"  href="#" title="Links I Like" alt="Links I Like"/> </div>
<script type="text/javascript">
// BeginOAWidget_Instance_2146022: #PrototypeFishEye

	var PrototypeFishEye = function(e){
		var ele = Event.element(e);
		window.location = ele.getAttribute("href");
	}
	var PrototypeFishEyelinkRelay = new FishEyeToolBar("PrototypeFishEye");
	PrototypeFishEye.addEventListener("itemClick", PrototypeFishEyelinkRelay);
	
// EndOAWidget_Instance_2146022
</script>
</body>
</html>
