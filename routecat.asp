<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
If GetSecureVal(Request.Querystring("id") <> "") Then
	catid = GetSecureVal(Request("id"))
	If LEN(catid) > 0 Then 
		
		Session("MM_CatalogID") = catid	
		response.redirect( "catalog_maint_CAT.asp")
		
	End If

else
	
	response.redirect("index.asp")
	
end if	
%>