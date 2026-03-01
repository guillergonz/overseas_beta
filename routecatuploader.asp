<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
If GetSecureVal(Request.Querystring("id") <> "") Then
	titulo = GetSecureVal(Request("id"))
	If LEN(titulo) > 0 Then 
		
		Session("MM_FileId") = titulo	

		response.redirect( "catuploader.asp")
		
	
		
	End If

else
	
	response.redirect("index.asp")
	
end if	
%>