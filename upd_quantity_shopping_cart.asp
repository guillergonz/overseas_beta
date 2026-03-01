<%@LANGUAGE="VBSCRIPT"%>
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
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<% 
'--------------------------------------------------------------------------
' THE PURPOSE OF THIS PAGE IS:
'
' 1) UPDATE QUANTITY ONHAND CHANGES IN SHOPPING CART
' 2) SHOW SHOPPING CART IN TABLE FORMAT.
'
'--------------------------------------------------------------------------
shop_id		= Request("p")
page_source = Request("t")
quantity	= Request("q")

'response.write ("Shop id" + shop_id + "<br>")
'response.write ("Show type" + page_source + "<br>")
'response.write ("New qty" + quantity + "<br>")

    
if quantity = 0 Then
    DelFromCart shop_id
else
    UpdateCart shop_id,quantity
end if

If page_source = "1" Then 
	CartDisplay("N")
Else
	CartDisplay("Y")
End if


%>