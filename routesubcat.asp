<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
If GetSecureVal(Request.Querystring("id") <> "") Then
	catid = GetSecureVal(Request("id"))
	If LEN(catid) > 0 Then 
		
		If GetSecureVal(Request.Querystring("titulo") <> "") Then
			vtitulo = GetSecureVal(Request("titulo"))
			If LEN(vtitulo) > 0 Then 
		
				Session("MM_Titulo") = vtitulo	
				Session("MM_CatID") = catid	
				response.redirect( "catmaint.asp")
					
			Else
				response.redirect("index.asp")		
			End If
		Else
			response.redirect("index.asp")
		End If	
		
	Else
		response.redirect("index.asp")	
	End If

else
	
	response.redirect("index.asp")
	
end if	
%>
