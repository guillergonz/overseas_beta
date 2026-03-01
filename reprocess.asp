<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/overseaspr.asp" -->

<% 

'Declare our variables
Dim vorder_number, vorder_id, sql, objRS2, sError
vorder_number = trim(request.form("orden"))
if IsNull(vorder_number) or LEN(vorder_number) = 0 then

	Response.Write("Orden #: "+Request.Form("orden")+"<br>" )
	Response.redirect "monitor_beta.asp" 

end if	

'UPDATE processed "P" to "O"
Set objRS2 = Server.CreateObject("ADODB.Connection")
objRS2.Open MM_overseaspr_STRING
sql="UPDATE dbo.clients_orders SET order_status = 'O' WHERE order_number = '" & vorder_number & "' ;"
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