<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="freeASPUpload/freeaspupload.asp" -->
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
 if Session("MM_UserName") <> "Z099" then 
 	MM_authFailedURL="index.asp"
    Response.Redirect(MM_authFailedURL)
 End if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="overseas.css" rel="stylesheet" type="text/css" media="screen" />

<%
Dim dfilename, referer

response.write( "part " & GetSecureVal(Request.Querystring("part")) & "<br>" )

If GetSecureVal(Request.Querystring("part") <> "") Then
	dfilename = GetSecureVal(Request("part"))
	referer = GetSecureVal(Request.servervariables("http_referer"))
	If LEN(dfilename) > 0 Then 
		
		Session("MM_FileId") = dfilename	
		Session("MM_referer") = referer
		response.redirect( "/ggguploader.asp")
		
		response.write("MM_FileId " & dfilename & "<br>")
		response.write("MM_referer " & referer & "<br>")		
		
	End If

else
	
	response.redirect("catalog_beta.asp")
	
end if	
%>

</head>

<body>
</body>
</html>
