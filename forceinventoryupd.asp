<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/overseaspr.asp" -->

<% 

'Declare our variables
Dim forcecheck, sql, objRS2, sError
forcecheck = trim(request.form("forcecheck"))

Response.Write("forcecheck #: "+Request.Form("forcecheck")+"<br>" )

if IsNull(forcecheck) or LEN(forcecheck) = 0 then
	Response.redirect "monitor_beta.asp" 
end if	

Set objRS2 = Server.CreateObject("ADODB.Connection")
objRS2.Open MM_overseaspr_STRING
sql=" UPDATE dbo.counters SET forceinventoryupd = '1', forcedatetime = '" + CSTR(Now()) + "' "
objRS2.Execute sql
objRS2.Close
Set objRS2 = nothing
sError = err.description
If len(sError) > 0 Then
	Response.Write("Ocurrió el siguiente error !<br><br><strong>" & sError & "</strong")
else
'	Response.Write("Ok")
	Response.redirect "monitor_beta.asp"
End if


%>