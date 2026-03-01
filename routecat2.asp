<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%



If GetSecureVal(Request.Querystring("titulo") <> "") Then
	vtitulo = GetSecureVal(Request("titulo"))
	If LEN(vtitulo) > 0 Then 
		Session("MM_Titulo") = vtitulo	
		response.redirect( "catmaint.asp")
	End If
else
	If GetSecureVal(Request.Querystring("id") <> "") Then
	
		If IsNumeric(request("id")) then
			nsubcatid = CInt(request("id"))
			If nsubcatid > 0 Then 
				Session("MM_SubCatID") = nsubcatid
				response.redirect( "catmaint.asp")
			End If
		End If
		
	Else
		response.redirect("index.asp")
	End If
end if	
%>