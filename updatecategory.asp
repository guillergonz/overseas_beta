<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%

titulo = GetSecureVal(Request.Querystring("tit"))

oldcatvalue = GetSecureVal(Request.Querystring("oldcat"))
oldsubvalue = GetSecureVal(Request.Querystring("oldsub"))
catvalue = GetSecureVal(Request.Querystring("cat"))
subvalue = GetSecureVal(Request.Querystring("sub"))
if LEN(catvalue) > 0 AND LEN(subvalue) > 0 then

	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open MM_overseaspr_STRING
	sql="UPDATE dbo.OIC_PartsPerCategory SET categoryid = " + CStr(catvalue) + ", subcategoryid = " + CStr(subvalue) + " WHERE titulo = '" + trim(titulo) + "' AND categoryid = " + CStr(oldcatvalue) + " AND subcategoryid = " + CStr(oldsubvalue)
	objConn.Execute sql

	response.write( sql )
	
	sError = err.description
	objConn.Close
	Set objConn = Nothing

	If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
	else
		response.redirect("catalog_maint2.asp")
	End if

else
	
	response.redirect("index.asp")
	
end if	
%>
