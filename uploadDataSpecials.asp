<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" --> 
<%


Set objConn55 = Server.CreateObject("ADODB.Connection")
objConn55.Open MM_overseaspr_STRING

sql="TRUNCATE TABLE dbo.prespecials"
on error resume next
objConn55.Execute sql
sError = err.description
If len(sError) > 0 Then
	Response.Write("<h6 class='red'>DELETE error prespecials: " & sError & "</h6>")
End if
objConn55.Close 

		
Dim position
Dim fieldno
Dim field_1,field_2,field_3,field_4,field_5
Dim lineData,lineData2
Dim recprocessed
recprocessed = 0

Set fso = Server.CreateObject("Scripting.FileSystemObject") 
set fs = fso.OpenTextFile(Server.MapPath("../apps/OVERSEAS SQL/Webdata/SPECIALS.CSV"), 1, true) 
Do Until fs.AtEndOfStream 
    lineData = fs.ReadLine
	lineData2 = lineData
    'inventory
	'specials char(20)
	'qty char(90)
	'description float
	'regular float
	'especial Integer
	
	IF len(lineData) > 0 THEN
	
		recprocessed = recprocessed + 1
		fieldno = 0
		Do While InStr(1, lineData, "," ,vbTextCompare) > 0
			fieldno = fieldno + 1
			position = InStr(1, lineData, "," ,vbTextCompare)
			if fieldno = 1 then
				' partno
				field_1 = mid(lineData,1,position-1)
			elseif fieldno = 2 then
				' description
				field_2 = mid(lineData,1,position-1)	 
			elseif fieldno = 3 then
				' price level1
				field_3 = mid(lineData,1,position-1)	 
			elseif fieldno = 4 then
				' price level3
				field_4 = mid(lineData,1,position-1)
			else
				lineData = ""					
			end if
			if fieldno < 5 then
				lineData = mid(lineData,position+1)
			end if
		Loop
		' Especial
		if LEN(lineData) > 0 then
			field_5 = lineData
		
			Dim objConn5 
			Dim objRs 
	
			Set objConn5 = Server.CreateObject("ADODB.Command") 
			objConn5.ActiveConnection = MM_overseaspr_STRING
			
			objConn5.commandtext="INSERT into prespecials(specials,qty,description,regular,especial)values(?,?,?,?,?)"
			
			objConn5.Parameters(0) = field_1
			objConn5.Parameters(1) = field_2
			objConn5.Parameters(2) = field_3
			objConn5.Parameters(3) = field_4
			objConn5.Parameters(4) = field_5
			
			on error resume next	
			set objRs = objConn5.execute
			sError = err.description
			If len(sError) > 0 Then
				Response.Write("<h6 class='red'>DELETE error prespecials: " & sError & "</h6>")
			End if

		Else
			Response.Write("EOL error!")
			response.end
		End If
		
	End If		
	
Loop 
response.Write("<h6>SPECIALS.CSV : " + trim(recprocessed) + "</h6>")
fs.close: set fs = nothing 


''PROD_LIQUI
'recprocessed = 0
'
'Set objConn56 = Server.CreateObject("ADODB.Connection")
'objConn56.Open MM_overseaspr_STRING
'
'sql="TRUNCATE TABLE dbo.prod_liqui"
'on error resume next
'objConn56.Execute sql
'sError = err.description
'If len(sError) > 0 Then
'	Response.Write("<h6 class='red'>DELETE error prod_liqui: " & sError & "</h6>")
'End if
'objConn56.Close 
'
'		
'
'recprocessed = 0
'
'Set fso = Server.CreateObject("Scripting.FileSystemObject") 
'set fs = fso.OpenTextFile(Server.MapPath("../apps/OVERSEAS SQL/Webdata/PROD_LIQUI.CSV"), 1, true) 
'Do Until fs.AtEndOfStream 
'    lineData = fs.ReadLine
'	lineData2 = lineData
'    'prod_liqui
'	'field_1 char(20)
'	'description varchar(50)
'	'regular money
'	'price money
'	
'	IF len(lineData) > 0 THEN
'	
'		recprocessed = recprocessed + 1
'		fieldno = 0
'		Do While InStr(1, lineData, "," ,vbTextCompare) > 0
'			fieldno = fieldno + 1
'			position = InStr(1, lineData, "," ,vbTextCompare)
'			if fieldno = 1 then
'				' field_1
'				field_1 = mid(lineData,1,position-1)	 				
'			else
'				lineData = ""					
'			end if	
'			if fieldno < 2 then
'				lineData = mid(lineData,position+1)
'			end if
'		Loop
'		if LEN(lineData) > 0 then
'			field_2 = lineData	 	
'		end if
'		
'		if LEN(field_2) > 0 then
'			Dim objConn555 
'	
'			Set objConn555 = Server.CreateObject("ADODB.Command") 
'			objConn555.ActiveConnection = MM_overseaspr_STRING
'			
'			objConn555.commandtext="INSERT into dbo.prod_liqui(field_1,price)values(?,?)"
'			
'			objConn555.Parameters(0) = field_1
'			objConn555.Parameters(1) = Cdbl(field_2)
'			
'			on error resume next	
'			set objRs = objConn555.execute
'			sError = err.description
'			If len(sError) > 0 Then
'				Response.Write("<h6 class='red'>INSERT error prod_liqui: " & sError & "</h6>")
'			End if
'
'		End If
'		
'	End If		
'			
'    'Response.Write lineData2
'	'Exit Do
'	
'Loop 
'response.Write("<h6>PROD_LIQUI.CSV : " + trim(recprocessed) + "</h6>")
'fs.close: set fs = nothing 

%>
