<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
If GetSecureVal(Request.Querystring("id") <> "") Then
	categoryid = GetSecureVal(Request.Querystring("id"))
	If categoryid > 0 Then 
		newvalue = GetSecureVal(Request.Querystring("newval"))
		if LEN(newvalue) > 0 then
			Set objConn = Server.CreateObject("ADODB.Connection")
			objConn.Open MM_overseaspr_STRING
			sql="UPDATE dbo.OIC_Category SET category = '" + newvalue + "' WHERE categoryid = " + CStr(categoryid) 
			objConn.Execute sql
			sError = err.description
			If len(sError) > 0 Then
				Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
			End if
			objConn.Close
			Set objConn = Nothing
			
			response.redirect("catalog_maint_CAT.asp")
			
		End If
				
	End If

else
	
	response.redirect("index.asp")
	
end if	
%>
