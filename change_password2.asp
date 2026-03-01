<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
pwd		= Request("pwd1")
user	= Session("MM_Username2")

'Response.Write( pwd )
'Response.Write( user )

If len(user) > 0 Then

	ChangePassword user,pwd
	Response.Redirect "index.asp"

Else

	Response.Redirect "change_password.asp"
	
End if

%>