<%@LANGUAGE="VBSCRIPT"%>
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



%>
<% 
'--------------------------------------------------------------------------
' THE PURPOSE OF THIS PAGE IS:
'
' 1) ADD ITEM TO SHOPPING CART TABLE USING USER ID FROM SESSION VARIABLE
' 2) SHOW SHOPPING CART IN TABLE FORMAT.
'
'--------------------------------------------------------------------------
num_pieza	= Request("p")
amount		= Request("a")



AddToCart num_pieza,amount
CartDisplay("N")

Session("MM_Update_Message") = "OK"
'response.redirect("catmaint.asp")

%>