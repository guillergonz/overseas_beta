<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
If GetSecureVal(Request.Querystring("id") <> "") Then
	
	subcategoryid = GetSecureVal(Request("id"))
	If LEN(subcategoryid) > 0 Then 		
		Session("MM_SubCatID") = subcategoryid	
		response.redirect( "catalog_maint2.asp")
	End If

else
	
	response.redirect("index.asp")
	
end if	
%>