<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%

If GetSecureVal(Request.Querystring("id") <> "") Then
	user_auto_id = GetSecureVal(Request.Querystring("id"))
	If LEN(user_auto_id) > 0 Then 
		'newvalue = GetSecureVal(Request.Querystring("newval"))
		'if LEN(newvalue) > 0 then
			Set objConn = Server.CreateObject("ADODB.Connection")
			objConn.Open MM_overseaspr_STRING
			sql="UPDATE dbo.users SET activestatus = 'A',user_pwd = '11111' WHERE user_auto_id = '" + CStr(user_auto_id) + "'" 
			objConn.Execute sql
			sError = err.description
			If len(sError) > 0 Then
				Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
			End if
			objConn.Close
			Set objConn = Nothing
			
			response.redirect("usuarios.asp")
			
		'End If
				
	End If

else
	
	response.redirect("index.asp")
	
end if	
%>
