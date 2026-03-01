<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%

If GetSecureVal(Request.Querystring("id") <> "") Then

	

	subcategoryid = GetSecureVal(Request.Querystring("id"))
	If subcategoryid = 0 Then 
		newvalue = GetSecureVal(Request.Querystring("newval"))
		if LEN(newvalue) > 0 AND LEN(Session("MM_CatID")) > 0 AND LEN(Session("MM_SubCatID")) > 0 AND LEN(Session("MM_CatalogID")) > 0 then

			Set objConn = Server.CreateObject("ADODB.Connection")
			objConn.Open MM_overseaspr_STRING
			
			sql="Insert INTO dbo.OIC_PartsPerCategory (partno,image,link,idcatalog,categoryid,subcategoryid,titulo,nota) VALUES ('','',''," + CStr(Session("MM_CatalogID")) + "," + CStr(Session("MM_CatID")) + "," + CStr(Session("MM_SubCatID")) + ",'" + newvalue + "','');"  
			objConn.Execute sql
			sError = err.description
			If len(sError) > 0 Then
				Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
			End if
			
			response.redirect("catalog_maint2.asp")
			
			objConn.Close
			Set objConn = Nothing
		
			oRS2.close
			Set oRS2 = null
						
		End If			
	End If

else
	
	response.redirect("index.asp")
	
end if	
%>
