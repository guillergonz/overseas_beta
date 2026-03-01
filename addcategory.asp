<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="index.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>

<% 



'--------------------------------------------------------------------------
' THE PURPOSE OF THIS PAGE IS:
'
' 1) ADD ITEM TO OIC_CATALOG_CATEGORY

'				Session("MM_CatID") = 1	
'				Session("MM_CatalogID") = 1
'				Session("MM_Titulo") = ""	
'--------------------------------------------------------------------------
vcatid 	  = Session("MM_CatalogID")

If LEN(request("nombrecat")) > 0 AND LEN(request("nombresubcat")) > 0 Then
	
	vcategory = request("nombrecat") 
	vcategoryid = GetCategoriadeNombre(vcategory)
	if vcategoryid = 0 then
		AddCat vcategory
		vcategoryid = GetCategoriadeNombre(vcategory)
	end if
	if vcategoryid > 0 then	
		vtitulo = request("nombresubcat")
		AddCategory vcatid,vcategoryid,vtitulo
		Session("MM_CatID") = vcategoryid
		Session("MM_Titulo") = vtitulo
	end if

End If
response.redirect("catalog_maint2.asp")
%>