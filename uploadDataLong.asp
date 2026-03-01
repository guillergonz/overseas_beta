<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" --> 
<%

Set oRS2 = Server.CreateObject("ADODB.Command")
oRS2.ActiveConnection = MM_overseaspr_STRING
oRS2.CommandText = "TRUNCATE TABLE dbo.long1"
on error resume next
oRS2.Execute
sError = err.description
If len(sError) > 0 Then
	Response.Write("<h6 class='red'>TRUNCATE long1: " + sError + "</h6>")
End if
oRS2.ActiveConnection.Close

'Set objConn55 = Server.CreateObject("ADODB.Connection")
'objConn55.Open MM_overseaspr_STRING
'
'sql="DELETE FROM dbo.long1"
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
Dim field_1,field_2
Dim lineData,lineData2
Dim recprocessed
recprocessed = 0

Set fso = Server.CreateObject("Scripting.FileSystemObject") 
set fs = fso.OpenTextFile(Server.MapPath("../apps/OVERSEAS SQL/Webdata/LONG1.csv"), 1, true) 
Do Until fs.AtEndOfStream 
    lineData = fs.ReadLine
	lineData2 = lineData
    'familicat/familicaten
	'field_1 char(20)
	'field_2 char(90)
		
	IF len(lineData) > 0 THEN
	
		recprocessed = recprocessed + 1
		fieldno = 0
		Do While InStr(1, lineData, "," ,vbTextCompare) > 0
			fieldno = fieldno + 1
			position = InStr(1, lineData, "," ,vbTextCompare)
			if fieldno = 1 then
				' fam_make_item
				field_1 = trim(mid(lineData,1,position-1))
			else
				linedata = ""		 		
			end if
			if fieldno < 2 then
				lineData = trim(mid(lineData,position+1))
			end if
		Loop
		' itemno
		if len(lineData) > 0 then
			field_2 = lineData
		
			Dim objConn5 
			Dim objRs 
				
			Set objConn5 = Server.CreateObject("ADODB.Command") 
			objConn5.ActiveConnection = MM_overseaspr_STRING
					
			objConn5.commandtext="INSERT into dbo.long1(fam_make_item,item_no)values(?,?)"
			
			objConn5.Parameters(0) = field_1
			objConn5.Parameters(1) = field_2
			
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
	
Loop 


response.Write("<h6>LONG1.CSV : " + trim(recprocessed) + "</h6>" )
fs.close: set fs = nothing 
%>
