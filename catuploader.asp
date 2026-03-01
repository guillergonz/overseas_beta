<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="freeASPUpload/freeaspupload.asp" -->
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<% 

If LEN(Request.QueryString) > 0 Then
	FileId = GetSecureVal(Request.QueryString("partno"))
	Session("MM_FileId") = FileId
End If

Response.Expires = -1
Server.ScriptTimeout = 300
' All communication must be in UTF-8, including the response back from the request
Session.CodePage  = 65001
%>
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


	if Session("MM_UserName") <> "Z099" then 
		MM_authFailedURL="index.asp"
		Response.Redirect(MM_authFailedURL)
	End if

	Dim FileId
	if Len(Session("MM_FileId")) = 0 or IsEmpty(Session("MM_FileId")) then
		if Len(Session("MM_Titulo")) = 0 then
			response.redirect("/index.asp")
		else
			
			FileID = Session("MM_Titulo") 
			Session("MM_FileId") = FileID
				
		end if			
	else
		FileID = Session("MM_FileId")
	end if	
	
  ' ****************************************************
  ' Change the value of the variable below to the pathname
  ' of a directory with write permissions, for example "C:\Inetpub\wwwroot"
  ' ****************************************************
  Dim uploadsDirVar
  uploadsDirVar = "C:\overseas_beta\parts_images\"

  ' Note: this file uploadTester.asp is just an example to demonstrate
  ' the capabilities of the freeASPUpload.asp class. There are no plans
  ' to add any new features to uploadTester.asp itself. Feel free to add
  ' your own code. If you are building a content management system, you
  ' may also want to consider this script: http://www.webfilebrowser.com/

function OutputForm()
%>

<div style="float:right;margin-right:50px">
    <a id="gobacktosource" align="center"  href="<%= Session("MM_referer") %>" >Regresar ...</a> 
    </div>
    
	    
	
    
<form style="margin-left:50px"  class="ui-widget" name="frmSend" method="POST" enctype="multipart/form-data" accept-charset="utf-8" action="/catuploader.asp" onSubmit="return onSubmitForm();" >
    <p style="width:290px" >File Name:&nbsp;&nbsp;<%= (Session("MM_FileId"))%>.jpg
	<br>
    File 1: <input id="attach1" name="attach1" type="file" size=35><p>
    
    <input style="float:left;margin-left:250px;clear:none" type=submit value="Subir / Upload">
    
    </form>
    
    
    
    
    
    
    
<%
end function

function TestEnvironment()
    Dim fso, fileName, testFile, streamTest
    TestEnvironment = ""
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    if not fso.FolderExists(uploadsDirVar) then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not exist.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
        exit function
    end if
    fileName = uploadsDirVar & "\test.txt"
    on error resume next
    Set testFile = fso.CreateTextFile(fileName, true)
    If Err.Number<>0 then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not have write permissions.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
        exit function
    end if
    Err.Clear
    testFile.Close
    fso.DeleteFile(fileName)
    If Err.Number<>0 then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not have delete permissions</B>, although it does have write permissions.<br>Change the permissions for IUSR_<I>computername</I> on this folder."
        exit function
    end if
    Err.Clear
    Set streamTest = Server.CreateObject("ADODB.Stream")
    If Err.Number<>0 then
        TestEnvironment = "<B>The ADODB object <I>Stream</I> is not available in your server.</B><br>Check the Requirements page for information about upgrading your ADODB libraries."
        exit function
    end if
    Set streamTest = Nothing
end function

function SaveFiles
    Dim Upload, fileName, fileSize, ks, i, fileKey

	Dim strDirectory, strFile, objFSO

    Set Upload = New FreeASPUpload
    Upload.Save(uploadsDirVar)

	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then Exit function

    SaveFiles = ""
    ks = Upload.UploadedFiles.keys
    if (UBound(ks) <> -1) then
        SaveFiles = "<br><B>Files uploaded:</B> "
        for each fileKey in Upload.UploadedFiles.keys
			
			
	
			SaveFiles = SaveFiles & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "

			Set objFSO = CreateObject("Scripting.FileSystemObject")

			strDirectory = "C:\overseas_beta\parts_images\" & Session("MM_FileId") & ".jpg"
			strFile = uploadsDirVar & "\" & Upload.UploadedFiles(fileKey).FileName 
		
			response.write "Part No " + Session("MM_FileId") + "<br>"
			 
			Image_exist( Session("MM_FileId") )
			response.write("<br>" & "From " & strFile & "    TO " & strDirectory  )
			
			response.write("<br>Solamente se aceptan archivos JPG!<br>Session FileId = " & Session("MM_FileId") & "<br>" )

			RESPONSE.WRITE("File " + strFile + " copied to " +  strDirectory)
			'objFSO.CopyFile strFile, strDirectory 

        next
    else
        SaveFiles = "No file selected for upload or the file name specified in the upload form does not correspond to a valid file in the system."
    end if
	'SaveFiles = SaveFiles & "<br>Enter a number = " & Upload.Form("enter_a_number") & "<br>"
'	SaveFiles = SaveFiles & "Checkbox values = " & Upload.Form("checkbox_values") & "<br>"
'	SaveFiles = SaveFiles & "List values = " & Upload.Form("list_values") & "<br>"
'	SaveFiles = SaveFiles & "Text area = " & Upload.Form("t_area") & "<br>"

	
	
end function
%>



<!doctype html><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">

<title>Overseas Import Corporation</title>
<link rel="icon" href="images/favicon.ico" type="image/x-icon" /> 
<link rel="icon" href="images/favicon.ico" type="image/x-icon" />

<link href="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/css/ui-darkness/jquery-ui-1.10.4.custom.css" rel="stylesheet">
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-1.10.2.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-ui-1.10.4.custom.js"></script>

   
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/jquery-1.10.2.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.core.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.widget.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.button.min.js"></script>
<!--<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.menu.min.js"></script>-->
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.position.min.js"></script>
<!--<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.tooltip.min.js"></script>-->
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/jquery.ui.autocomplete.js"></script>

<!--<script type="text/javascript" src="jquery-ui/jquery-ui-1.8.16-sunny.custom/js/jquery-1.6.2.min.js"></script>
<script type="text/javascript" src="jquery-ui/jquery-ui-1.8.16-sunny.custom/js/jquery-ui-1.8.16.custom.min.js"></script>
<link type="text/css" href="jquery-ui/jquery-ui-1.8.16-sunny.custom/css/sunny/jquery-ui-1.8.16.custom.css" rel="stylesheet" />
<script src="jquery-ui/jquery-ui-1.8.16-sunny.custom/development-bundle/ui/jquery.ui.core.js"></script>
<script src="jquery-ui/jquery-ui-1.8.16-sunny.custom/development-bundle/ui/jquery.ui.widget.js"></script>
<script src="jquery-ui/jquery-ui-1.8.16-sunny.custom/development-bundle/ui/jquery.ui.datepicker.js"></script>
<script src="jquery-ui-1.9.0.custom/jquery-ui-1.9.0.custom/development-bundle/ui/jquery.ui.button.js"></script>-->

<style type="text/css">
body {

	margin-left: 10px;
	margin-top: 10px;
	margin-right: 10px;
	margin-bottom: 10px;
	background-color: #000;
	color:white;
}
</style>

<script type="text/javascript">
$(document).ready(function() {
	
	$( "input[type=submit], input[type=button], #gobacktosource " )
		.button()
		.click(function( event ) {
			//event.preventDefault();
		});
		
});
</script>

<!-- check select, radio buttons, checkboxes -->
<link rel="stylesheet" type="text/css" href="style2.css" />
<style type="text/css">
body,td,th {
	font-size: 14px;
	color: #FFF;
}
.uploaderform {
	width: 1000px;
	margin-left: auto;
	margin-right: auto;
	height: 450px;
	display: block;
}
</style>


<script>
function onSubmitForm() {
    var formDOMObj = document.frmSend;
    if (formDOMObj.attach1.value == "" && formDOMObj.attach2.value == "" && formDOMObj.attach3.value == "" && formDOMObj.attach4.value == "" )
        alert("Please press the Browse button and pick a file.")
    else
        return true;
    return false;
}
</script>
</head>
<body>


<div class="ui-widget-content ui-corner-all" style="margin-top: 10px; border-width: thin; width: 1000px; margin-right: auto; margin-left: auto; background-image: url(/secondarytile.png)">
    
  <div class="uploaderform" >
  
  <h2 class="ui-state-active" >&nbsp;&nbsp;Subir imágenes del catálogo </h2>
  <%
Dim diagnostics
if Request.ServerVariables("REQUEST_METHOD") <> "POST" then
    diagnostics = TestEnvironment()
    if diagnostics<>"" then
        response.write "<div style=""margin-left:20; margin-top:30; margin-right:30; margin-bottom:30;"">"
        response.write diagnostics
        response.write "<p>After you correct this problem, reload the page."
        response.write "</div>"
    else
        response.write "<div style=""margin-left:150"">"
        OutputForm()
        response.write "</div>"
    end if
else
    response.write "<div style=""margin-left:150"">"
    OutputForm()
    response.write SaveFiles()
    response.write "<br><br></div>"
	response.redirect("catmaint.asp")
	
	'response.write ("<div><iframe align=""top"" src=""/parts_images/directory_browsing.asp"" width=""100%"" height=""500px"" style=""overflow-x:auto; overflow-y:auto;"" frameborder=""0"" scrolling=""Yes""></iframe></div>  ")
	
	
end if

%>
    
    
<br clear="all">
 
  <br clear="all">
 
  
   
    
</div>

 
    

</body>
</html>
