<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%



Response.Expires = -1 

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

'--------------------------------------------------------------------------
' THE PURPOSE OF THIS PAGE IS:
'
' 1) DELETE ITEM FROM SHOPPING CART TABLE USING SHOP_AUTO_ID
' 2) SHOW SHOPPING CART IN TABLE FORMAT.
'
'--------------------------------------------------------------------------
shop_id		= Request("p")
page_source = Request("t")

DelFromCart shop_id

If IsNull(page_source) Then page_source = "0"
If page_source = "1" Then 
	CartDisplay("N")
Else
	CartDisplay("Y")
End if


%>