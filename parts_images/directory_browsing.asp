
<%@LANGUAGE="VBSCRIPT"%>
<%
   Option Explicit
   On Error Resume Next

	'Response.Write "LOGON_USER: " & Request.ServerVariables("LOGON_USER") & "<br>"
'	Response.Write "REMOTE_USER: " & Request.ServerVariables("REMOTE_USER") & "<br>"
'	Response.Write "AUTH_USER: " & Request.ServerVariables("AUTH_USER") & "<br>"
'	Response.Write "<br>"
'	'Show all server variables
'	For Each Item In Request.ServerVariables
'	Response.Write Item & " = " & Request.ServerVariables(Item) & "<br>"
'	Next

   ' this section is optional - it just denies anonymous access
'   If Request.ServerVariables("LOGON_USER")="" Then
'      Response.Status = "401 Access Denied GGG"
'   End If

   ' declare variables
   Dim objFSO, objFolder
   Dim objCollection, objItem

   Dim strPhysicalPath, strTitle, strServerName
   Dim strPath, strTemp
   Dim strName, strFile, strExt, strAttr
   Dim intSizeB, intSizeK, intAttr, dtmDate

   ' declare constants
   Const vbReadOnly = 1
   Const vbHidden = 2
   Const vbSystem = 4
   Const vbVolume = 8
   Const vbDirectory = 16
   Const vbArchive = 32
   Const vbAlias = 64
   Const vbCompressed = 128

   ' don't cache the page
   Response.AddHeader "Pragma", "No-Cache"
   Response.CacheControl = "Private"

   ' get the current folder URL path
   strTemp = Mid(Request.ServerVariables("URL"),2)
   strPath = ""

   Do While Instr(strTemp,"/")
      strPath = strPath & Left(strTemp,Instr(strTemp,"/"))
      strTemp = Mid(strTemp,Instr(strTemp,"/")+1)      
   Loop

   strPath = "/" & strPath

   ' build the page title
   strServerName = UCase(Request.ServerVariables("SERVER_NAME"))
   strTitle = "Contents of the " & strPath & " folder"

   ' create the file system objects
   strPhysicalPath = Server.MapPath(strPath)
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set objFolder = objFSO.GetFolder(strPhysicalPath)
%>
<html>
<head>
<title>Overseas PR</title>
<meta name="GENERATOR" content="CEPR">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<link rel="icon" type="image/ico" href="/images/favicon.ico" title="favicon">

<style type="text/css">
body,td,th {
	font-size: 14px;
	font-family: "Segoe UI", "Trajan Pro", "Times New Roman";
}
body {
	background-image: url(AnimatedFrom/images/bg.gif);
	background-repeat: repeat;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #FFF;
}
.goback{
	float: right;
	clear: both;
	font-family: "Segoe UI", "Trajan Pro", "Times New Roman";
	font-size: 11px;
	padding: 5px;
	margin: 0px;
	font-style: normal;
	line-height: normal;
	text-decoration: none;
	font-weight: normal;
	display: inline;
}
.wrapper2{
	width: 968px;
	-webkit-border-radius: 10px;
	-moz-border-radius: 10px;
	border-radius: 10px;
	min-height: 550px;
	margin-top: 0px;
	margin-right: auto;
	margin-bottom: 0px;
	margin-left: auto;
	white-space: normal;
	display: block;
	padding: 0px;
	text-decoration: none;
	border: thin solid #7E9299;
	clear: both;
}

.ceprlogo {
	background-color: #FFFFFF;
	margin: 0px;
	text-align: center;
	padding: 0px;
	text-decoration: none;
	vertical-align: middle;
}
.jqmenu {
	clear: both;
	float: left;
	position: absolute;
	background-color: #FFF;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	margin-left: 3px;
	min-width: 140px;
}
</style>
<link rel="stylesheet" href="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/themes/base/jquery.ui.all.css">
<script src="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/jquery-1.8.2.js"></script>
<script src="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/ui/jquery.ui.core.js"></script>
<script src="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/ui/jquery.ui.widget.js"></script>
<script src="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/ui/jquery.ui.button.js"></script>
<link rel="stylesheet" href="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/demos/demos.css">

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<style type="text/css">
body,td,th {
	font-family: "Segoe UI", Verdana, "Times New Roman";
	font-size: 14px;
	color: rgba(0,0,0,1);
}
</style>
</head>
<body>

<!-- =strServerName%>-->
<!--<p align="center">Server:&nbsp;<=strServerName%> - <=strTitle%>
<p align="center" class="ui-widget-header"><=strTitle%></p>
-->



<div align="center"><center>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="ui-widget-content" >
<tr>
   <th width="11" >&nbsp;</th>
   <th width="108" align="center">Name</th>
   <th width="67" align="center">Bytes</th>
   <th width="47" align="center">KB</th>
   <!--<th width="103" align="left">Attributes</th> -->
   <th width="46" align="center">Ext</th>
   <th width="181" align="center">Type</th>
   <th width="108" align="center">Date</th>
   <th width="107" align="center">Time</th>
</tr>

<%
   ''''''''''''''''''''''''''''''''''''''''
   ' output the folder list
   ''''''''''''''''''''''''''''''''''''''''

   Set objCollection = objFolder.SubFolders

   For Each objItem in objCollection
      strName = objItem.Name
      strAttr = MakeAttr(objItem.Attributes)      
      dtmDate = CDate(objItem.DateLastModified)
%>
<tr>
   <th width="11" >&nbsp;</th>
   <td align="left"><b><a href="<%=strName%>"><%=strName%></a></b></td>
   <td align="right">N/A</td>
   <td align="right">N/A</td>
   <!--<td align="left"><tt><%=strAttr%></tt></td> -->
   <td align="left"><DIR></td>
   <td align="left"><b>Directory</b></td>
   <td align="left"><%=FormatDateTime(dtmDate,vbShortDate)%></td>
   <td align="left"><%=FormatDateTime(dtmDate,vbLongTime)%></td>
</tr>
<% Next %>

<%
   ''''''''''''''''''''''''''''''''''''''''
   ' output the file list
   ''''''''''''''''''''''''''''''''''''''''

   Set objCollection = objFolder.Files

   For Each objItem in objCollection
   
      strName = objItem.Name
      strFile = Server.HTMLEncode(Lcase(strName))

      intSizeB = objItem.Size
      intSizeK = Int((intSizeB/1024) + .5)
      If intSizeK = 0 Then intSizeK = 1

      strAttr = MakeAttr(objItem.Attributes)
      strName = Ucase(objItem.ShortName)
      If Instr(strName,".") Then strExt = Right(strName,Len(strName)-Instr(strName,".")) Else strExt = ""
      dtmDate = CDate(objItem.DateLastModified)
%>

<!--<%    If objItem.ShortName <> "DIRECT~1.ASP" then   %> -->
	
<tr>
   <th width="11" >&nbsp;</th>
   <td align="left"><a  target="_new" href="<%=strFile%>"><%=strFile%></a></td>
   <td align="right"><%=FormatNumber(intSizeB,0)%></td>
   <td align="right"><%=intSizeK%>Kb</td>
   <!--<td align="left"><tt><%=strAttr%></tt></td> -->
   <td align="center"><%=strExt%></td>
   <td align="left"><%=objItem.Type%></td>
   <td align="left"><%=FormatDateTime(dtmDate,vbShortDate)%></td>
   <td align="left"><%=FormatDateTime(dtmDate,vbLongTime)%></td>
</tr>

<!--<%    end if   %> -->

<% Next %>

</table>
</center></div>

</body>
</html>
<%
   Set objFSO = Nothing
   Set objFolder = Nothing

   ' this adds the IIf() function to VBScript
   Function IIf(i,j,k)
      If i Then IIf = j Else IIf = k
   End Function

   ' this function creates a string from the file atttributes
   Function MakeAttr(intAttr)
      MakeAttr = MakeAttr & IIf(intAttr And vbArchive,"A","-")
      MakeAttr = MakeAttr & IIf(intAttr And vbSystem,"S","-")
      MakeAttr = MakeAttr & IIf(intAttr And vbHidden,"H","-")
      MakeAttr = MakeAttr & IIf(intAttr And vbReadOnly,"R","-")
   End Function
%>
				

