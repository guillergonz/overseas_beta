<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" --> 
<%

Server.ScriptTimeout = 3600

Set oRS2 = Server.CreateObject("ADODB.Command")
oRS2.ActiveConnection = MM_overseaspr_STRING
oRS2.CommandText = "TRUNCATE TABLE dbo.invoice"
on error resume next
oRS2.Execute
sError = err.description
If len(sError) > 0 Then
	Response.Write("<h6 class='red'>TRUNCATE invoice: " + sError + "</h6>")
End if
oRS2.ActiveConnection.Close

'Set objConn55 = Server.CreateObject("ADODB.Connection")
'objConn55.Open MM_overseaspr_STRING
'
'sql="DELETE FROM dbo.invoice"
'objConn55.Execute sql
'sError = err.description
'objConn55.Close 
'
'If len(sError) > 0 Then
'	Response.Write("Ocurrió el siguiente error eliminando piezas del catálogo: <br><br>" & sError )
'	Response.End
'End if
	

	
Dim position
Dim fieldno
Dim field_1,field_2,field_3,field_4,field_5,field_6,field_7
Dim lineData,lineData2
Dim recprocessed
recprocessed = 0

Set fso = Server.CreateObject("Scripting.FileSystemObject") 
set fs = fso.OpenTextFile(Server.MapPath("../apps/OVERSEAS SQL/Webdata/invoice.csv"), 1, true) 
Do Until fs.AtEndOfStream 
    lineData = fs.ReadLine
	lineData2 = lineData
    'inventory
	'field_1 char(18)
	'field_2 float
	'field_3 char(20)
	'field_4 float
	'field_5 char(20)
	'field_6 float
	'field_7 float
	
	IF len(lineData) > 0 THEN
	
		recprocessed = recprocessed + 1
		fieldno = 0
		Do While InStr(1, lineData, "," ,vbTextCompare) > 0
			fieldno = fieldno + 1
			position = InStr(1, lineData, "," ,vbTextCompare)
			if fieldno = 1 then
				' field_1
				field_1 = trim(mid(lineData,1,position-1))
			elseif fieldno = 2 then
				' field_2
				field_2 = mid(lineData,1,position-1)	 
			elseif fieldno = 3 then
				' field_3
				field_3 = mid(lineData,1,position-1)	 
			elseif fieldno = 4 then
				' field_4
				field_4 = mid(lineData,1,position-1)
			elseif fieldno = 5 then
				' field_5
				field_5 = mid(lineData,1,position-1)	 
			elseif fieldno = 6 then
				' field_6
				field_6 = mid(lineData,1,position-1)
		
				
			else
				linedata = ""		 		
			end if
			if fieldno < 7 then
				lineData = mid(lineData,position+1)
			end if
		Loop
		' field_7
		if len(lineData) > 0 then

			field_7 = lineData
			Dim objConn5 
			Dim objRs 
				
			Set objConn5 = Server.CreateObject("ADODB.Command") 
			objConn5.ActiveConnection = MM_overseaspr_STRING
					
			objConn5.commandtext="INSERT into invoice(cust_inv,invamt,inv_date,payment_amt,payment_date,disc1,disc2,cust,inv)values(?,?,?,?,?,?,?,?,?)"
			
			objConn5.Parameters(0) = field_1
			objConn5.Parameters(1) = field_2
			objConn5.Parameters(2) = field_3
			objConn5.Parameters(3) = field_4
			objConn5.Parameters(4) = field_5
			objConn5.Parameters(5) = field_6
			objConn5.Parameters(6) = field_7
			objConn5.Parameters(7) = mid(field_1,1,4)
			objConn5.Parameters(8) = mid(field_1,5)
			
			set objRs = objConn5.execute
			sError = err.description

			If len(sError) > 0 Then
				Response.Write("Ocurrió el siguiente error añadiendo esta pieza al catálogo: <br><br>" & sError )
			End if

	
		Else
			Response.Write("EOL error!")
			response.end
		End If
		
	
	End If		
			
    'Response.Write lineData2
	'Exit Do
	
Loop 
response.Write("<h6>INVOICE.CSV : " + trim(recprocessed) + "</h6>")
fs.close: set fs = nothing 
%>
