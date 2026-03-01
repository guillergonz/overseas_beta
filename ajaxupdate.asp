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

userid = Request("userid")
pwd	= Request("pwd")


if LEN(userid) > 0 and LEN(pwd) > 0 then
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open MM_overseaspr_STRING
	sql="UPDATE dbo.multiuser SET password = '" + pwd + "' WHERE userid = '" + CStr(userid) + "'" 
	objConn.Execute sql
	sError = err.description
	If len(sError) > 0 Then
		response.Write("Ocurrió el siguiente error : <br><br>" & sError )
	Else
		response.write("<label style='color:red;font-siza:14px'>Record actualizado</label>")
	End if
	objConn.Close
	Set objConn = Nothing
end if	
%>