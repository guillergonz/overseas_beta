<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" --> 
<%


Set objConn55 = Server.CreateObject("ADODB.Connection")
objConn55.Open MM_overseaspr_STRING

sql="TRUNCATE TABLE dbo.familicat"
objConn55.Execute sql
sError = err.description
objConn55.Close 

If len(sError) > 0 Then
	Response.Write("Ocurrió el siguiente error eliminando piezas del catálogo: <br><br>" & sError )
	Response.End
End if
	
	
Dim position
Dim fieldno
Dim field_1,field_2
Dim lineData,lineData2
Dim recprocessed
recprocessed = 0

Set fso = Server.CreateObject("Scripting.FileSystemObject") 
set fs = fso.OpenTextFile(Server.MapPath("../apps/OVERSEAS SQL/Webdata/FAMILYCAT.csv"), 1, true) 
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
				' field_1 description spanish / english
				field_1 = mid(lineData,1,position-1)
			
			else
				linedata = ""		 		
			end if
			if fieldno < 2 then
				lineData = mid(lineData,position+1)
			end if
		Loop
		' field_2
		if len(lineData) > 0 then
			field_2 = lineData
		
			Dim objConn5 
			Dim objRs 
				
			Set objConn5 = Server.CreateObject("ADODB.Command") 
			objConn5.ActiveConnection = MM_overseaspr_STRING
			
			' Connect to the data source.
			'Set objConn = GetNewConnection
			'objConn5.ActiveConnection = objConn
		
			objConn5.commandtext="INSERT into familicat(field_1,field_2)values(?,?)"
			
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

Dim recprocessed2
Set fso = Server.CreateObject("Scripting.FileSystemObject") 
set fs = fso.OpenTextFile(Server.MapPath("../apps/OVERSEAS SQL/Webdata/FAMILYCATEN.csv"), 1, true) 
Do Until fs.AtEndOfStream 
    lineData = fs.ReadLine
	lineData2 = lineData
    'familicat/familicaten
	'field_1 char(20)
	'field_2 char(90)
	
	
	IF len(lineData) > 0 THEN
	
		recprocessed2 = recprocessed2 + 1
		fieldno = 0
		Do While InStr(1, lineData, "," ,vbTextCompare) > 0
			fieldno = fieldno + 1
			position = InStr(1, lineData, "," ,vbTextCompare)
			if fieldno = 1 then
				' field_1 description spanish / english
				field_1 = mid(lineData,1,position-1)
			
			else
				linedata = ""		 		
			end if
			if fieldno < 2 then
				lineData = mid(lineData,position+1)
			end if
		Loop
		' field_2
		if len(lineData) > 0 then
			field_2 = lineData
		end if
		
		if LEN(lineData) > 0 then
			field_2 = lineData
			
			Set objConn5 = Server.CreateObject("ADODB.Command") 
			objConn5.ActiveConnection = MM_overseaspr_STRING
			
			' Connect to the data source.
			'Set objConn = GetNewConnection
			'objConn5.ActiveConnection = objConn
		
			objConn5.commandtext="UPDATE familicat SET english_desc = ? WHERE field_2 = ? "
			
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

response.Write("<h6>FAMILYCAT.CSV : " + trim(recprocessed) + "&nbsp;&nbsp;&nbsp;FAMILYCATEN.CSV&nbsp;&nbsp;" + trim(recprocessed2) + "</h6>" )
fs.close: set fs = nothing 
%>
