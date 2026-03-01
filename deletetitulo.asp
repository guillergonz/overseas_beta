<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%


If GetSecureVal(Request.Querystring("uid") <> "") Then
	titulo = GetSecureVal(Request.Querystring("uid"))
	If LEN(titulo) > 0 Then 
		
			Set objConn = Server.CreateObject("ADODB.Connection")
			objConn.Open MM_overseaspr_STRING
			sql="DELETE FROM dbo.OIC_PartsPerCategory WHERE titulo = '" + CStr(titulo) + "' " 
			objConn.Execute sql
			sError = err.description
			If len(sError) > 0 Then
				Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
			End if
			objConn.Close
			Set objConn = Nothing
			
			response.redirect("catalog_maint2.asp")
			
	
				
	End If

else
	
	response.redirect("index.asp")
	
end if	
%>
