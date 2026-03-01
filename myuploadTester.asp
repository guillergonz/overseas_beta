<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% 
'option explicit

Response.Expires = -1
Server.ScriptTimeout = 300
' All communication must be in UTF-8, including the response back from the request
Session.CodePage  = 65001
%>

<!--#include file="freeASPUpload/freeaspupload.asp" -->
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->


<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "index.asp"
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

	Dim FileId
	
	FileId = GetSecureVal(Request.QueryString("nam"))
	if Len(FileId) > 0 then
		Session("MM_FileId") = FileId
	end if

	'response.write("Querystring " & FileId & "<br>" ) 
	'response.write("Session MM_FileId " & Session("FileId") & "<br>" ) 
	
	CreateUserDirectory() 
	
%>


<%


  ' ****************************************************
  ' Change the value of the variable below to the pathname
  ' of a directory with write permissions, for example "C:\Inetpub\wwwroot"
  ' ****************************************************

  Dim uploadsDirVar
  uploadsDirVar = "C:\Inetpub\ftproot\overseas_beta"
  

  ' Note: this file uploadTester.asp is just an example to demonstrate
  ' the capabilities of the freeASPUpload.asp class. There are no plans
  ' to add any new features to uploadTester.asp itself. Feel free to add
  ' your own code. If you are building a content management system, you
  ' may also want to consider this script: http://www.webfilebrowser.com/

function OutputForm()
%>
    <form name="frmSend" method="POST" enctype="multipart/form-data" accept-charset="utf-8" action="/cepr2/myuploadTester.asp" onSubmit="return onSubmitForm();">
	<B>File names:</B><br>
    File 1: <input id="attach1" name="attach1" type="file" size=35><br>
    
    <div style="display:none">File 2: <input id="attach2" name="attach2" type="file" size=35><br>
    File 3: <input id="attach3" name="attach3" type="file" size=35><br>
    File 4: <input id="attach4" name="attach4" type="file" size=35></div>
    
    <br>
    <br> 
	
    <input style="margin-top:4" type=submit value="Subir / Upload">
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;        
    <a id="goback1" align="center" href="javascript: history.go(-1)">Regresar ...</a> 
    
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
	
			strDirectory = "c:\dw\cepr2\uploadedfiles\" & Session("MM_UserID") & "\" & Session("MM_FileId")
			strFile = "c:\dw\cepr2\uploadedfiles\" & Upload.UploadedFiles(fileKey).FileName
		
			'response.write("<br>" & "From " & strFile & "    TO " & strDirectory  )
			
			response.write("<br>Solamente se aceptan archivos PDF!<br>Session FileId = " & Session("MM_FileId") & "<br>" )

			objFSO.CopyFile strFile, strDirectory 

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

<title>CEPR - PLES</title>
<link rel="icon" type="image/ico" href="/cepr2//images/faviconcepr.ico" title="favicon">

<style type="text/css">
body,td,th {
	font-size: 14px;
	font-family: "Segoe UI", "Trajan Pro", "Times New Roman";
}
body {
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
<script type="text/javascript">
$(document).ready(function() {
	
	$( "input[type=submit], input[type=button], #goback1 " )
		.button()
		.click(function( event ) {
			//event.preventDefault();
		});
		
 //       $('.jqmenu').hide();

//	$( "#JDtermino_desde" ).datepicker();
});
</script>

<!-- check select, radio buttons, checkboxes -->
<link rel="stylesheet" type="text/css" href="style2.css" />
<style type="text/css">
body,td,th {
	font-family: "Segoe UI", "Trajan Pro", "Times New Roman";
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
<body id="dt_example" >
<br clear="all">

<div class="wrapper2">
    
  <div style="margin:20px;">
  <br><br>
  <div style="border-bottom: #A91905 2px solid;font-size:16">Subir archivos a Overseas Import Corporation Web Server</div>
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
	
	
	response.write ("<div><iframe align=""top"" src=""/parts_images/directory_browsing.asp"" width=""100%"" height=""500px"" style=""overflow-x:auto; overflow-y:auto;"" frameborder=""0"" scrolling=""Yes""></iframe></div>  ")
	
	
end if

%>
    
    
  <br><br>
    
    
  
  </div>
  
  
  
   
    
</div>

 
    

</body>
</html>
