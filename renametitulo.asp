<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%

newvalue = GetSecureVal(Request.Querystring("newval"))
oldvalue = GetSecureVal(Request.Querystring("oldval"))
if LEN(newvalue) > 0 then

	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open MM_overseaspr_STRING
	sql="UPDATE dbo.OIC_PartsPerCategory SET titulo = '" + trim(newvalue) + "' WHERE titulo = '" + trim(oldvalue) + "' AND categoryid = " + CStr(Session("MM_CatID")) + " AND subcategoryid = " + CStr(Session("MM_SubCatID")) + " AND IdCatalog = " + CStr(Session("MM_CatalogID"))
	objConn.Execute sql
	sError = err.description
	If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
	End if
	objConn.Close
	Set objConn = Nothing
	
	response.redirect("catalog_maint2.asp")

else
	
	response.redirect("index.asp")
	
end if	
%>
