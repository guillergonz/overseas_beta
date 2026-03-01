<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" --> 
<%
Server.ScriptTimeout = 3600

dim sauto
sauto = request("auto")

if sauto = "true" then

	Session("MM_auto") = true	
	
	
	
	''DECLARE	@return_value int
'	'EXEC	@return_value = [dbo].[Kill_Overseas_Users]
'	'SELECT	'Return Value' = @return_value
'
'	Set Conn = CreateObject("ADODB.Connection")
'	
'	DIM cmd
'	SET cmd = Server.CreateObject("ADODB.Command")
'	SET cmd.ActiveConnection = Conn
'	
'	'Prepare the stored procedure
'	cmd.CommandText = "Kill_Overseas_Users"
'	cmd.CommandType = 4  'adCmdStoredProc
'	
'	'cmd.Parameters("@RECORD_NUMBER") = Request.Form("Record_Number") 
'	
'	'Execute the stored procedure
'	'This returns recordset but you dont need it
'	cmd.Execute
'	
'	Conn.Close
'	SET Conn = Nothing

										' Set oRS2 = Server.CreateObject("ADODB.Command") 
										' oRS2.ActiveConnection = MM_overseaspr_STRING	
										' oRS2.commandtext="ALTER DATABASE overseas_prod SET SINGLE_USER WITH ROLLBACK IMMEDIATE;"
										' on error resume next
										' oRS2.Execute
										' sError = err.description			
										' If len(sError) > 0 Then
											' Response.Write("<h6 class='red'>System error</h6>")
											' response.end
										' End if
										' oRS2.ActiveConnection.Close
										
										' Set oRS2 = Server.CreateObject("ADODB.Command") 
										' oRS2.ActiveConnection = MM_overseaspr_STRING	
										' oRS2.commandtext="ALTER DATABASE overseas_prod SET MULTI_USER;"
										' on error resume next
										' oRS2.Execute
										' sError = err.description			
										' If len(sError) > 0 Then
											' Response.Write("<h6 class='red'>System error</h6>")
											' response.end
										' End if
										' oRS2.ActiveConnection.Close


	' File permission error not working!
	
	'Set oRS = Server.CreateObject("ADODB.Connection")
'	oRS.Open MM_overseaspr_STRING
'	oRS.CommandTimeout = 60000
'	Set rs = oRS.Execute("EXEC [overseas_prod].[dbo].[MyBackup]")
'	sError = err.description
'	If len(sError) > 0 Then
'		Response.Write("Ocurrió el siguiente error creando backup: <br><br><strong>" & sError & "</strong")
'	End if
'	response.write("Database backup ended ...")


	' Enable website under maintenance
	Set oRS2 = Server.CreateObject("ADODB.Command") 
	oRS2.ActiveConnection = MM_overseaspr_STRING	
	oRS2.commandtext="UPDATE dbo.counters SET forceinventoryupd = ? "
	oRS2.Parameters(0) = "0"
	on error resume next
	oRS2.Execute
	'set oRecordSet = oRS.execute
	sError = err.description			
	If len(sError) > 0 Then
		Response.Write("<h6 class='red'>System error</h6>")
		response.end
	End if
	oRS2.ActiveConnection.Close
	
end if

'close oRecordSet
'set oRecordSet = nothing	

response.write("Cleanup Started ...")

Set oRS2 = Server.CreateObject("ADODB.Command")
oRS2.ActiveConnection = MM_overseaspr_STRING
oRS2.CommandText = "TRUNCATE TABLE dbo.partmst1_distinct"
on error resume next
oRS2.Execute
sError = err.description
If len(sError) > 0 Then
	Response.Write("<h6 class='red'>TRUNCATE partmst1_distinct: " + sError + "</h6>")
End if
oRS2.ActiveConnection.Close

'Set objConn55 = Server.CreateObject("ADODB.Connection")
'objConn55.Open MM_overseaspr_STRING

'sql="DELETE FROM dbo.partmst1_distinct"
'on error resume next
'objConn55.Execute sql

'objConn55.Close 
	
	
Dim position
Dim fieldno
Dim field_1,field_2,field_3,field_4,field_5,field_6,field_7,field_8
Dim lineData,lineData2
Dim recprocessed
recprocessed = 0

' CHECK FILE PERMISSION ACCESS   authenticated users
' Server.MapPath("../apps/OVERSEAS SQL/Webdata/INVENTORY.CSV")

path = "C:/apps/OVERSEAS SQL/Webdata/INVENTORY.CSV"

'path = "C:/apps/OVERSE~1/Webdata/INVENTORY.CSV"



Set fso = CreateObject("Scripting.FileSystemObject")
Set objFile = fso.GetFile(path)
sempty = "EMPTY"
If objFile.Size > 0 AND objFile.Size <> "Empty" Then
	
	sempty = objFile.Size
	response.write("File Read")
	
	set fs = fso.OpenTextFile(path, 1, true) 

 	'Set fso = Server.CreateObject("Scripting.FileSystemObject") 
 	'set fs = fso.OpenTextFile(Server.MapPath("../Webdata/INVENTORY.CSV"), 1, true) 

 Do Until fs.AtEndOfStream 
    lineData = fs.ReadLine
	lineData2 = lineData
    'inventory
	'field_1 char(20)
	'field_2 char(90)
	'field_3 float
	'field_4 float
	'QOH Integer
	'FamilyCat Integer
	'Precio4 float
	'Precio5 float
	
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
			elseif fieldno = 5 then
				' qoh
				field_5 = mid(lineData,1,position-1)	 
			elseif fieldno = 6 then
				' familycat
				field_6 = mid(lineData,1,position-1)
			elseif fieldno = 7 then
				' price level4
				field_7 = mid(lineData,1,position-1)
			else
				linedata = ""		 		
			end if
			if fieldno < 8 then
				lineData = mid(lineData,position+1)
			end if
		Loop
		' price level6
		if LEN(lineData) > 0 then
			field_8 = lineData
		
			Dim objConn5 
			Dim objRs 
	
			'Set objConn5 = Server.CreateObject("ADODB.Connection")
			'objConn5.Open MM_overseaspr_STRING
			
			Set objConn5 = Server.CreateObject("ADODB.Command") 
			objConn5.ActiveConnection = MM_overseaspr_STRING
		
			objConn5.commandtext="INSERT into partmst1_distinct(field_1,field_2,field_3,field_4,field_5,field_6,field_7,field_8)values(?,?,?,?,?,?,?,?)"
			
			objConn5.Parameters(0) = field_1
			objConn5.Parameters(1) = field_2
			objConn5.Parameters(2) = field_3
			objConn5.Parameters(3) = field_4
			objConn5.Parameters(4) = field_5
			objConn5.Parameters(5) = field_6
			objConn5.Parameters(6) = field_7
			objConn5.Parameters(7) = field_8
			
			on error resume next
			set objRs = objConn5.execute
			sError = err.description
			If len(sError) > 0 Then
				Response.Write("<h6 class='red'>INSERT error: " + field_1 + " - " + sError + "</h6>")
			End if

		Else
			Response.Write("EOL error!")
			response.end
		End If
		
	End If		
			
    'Response.Write lineData2
	'Exit Do
	
 Loop 
 response.Write("<h6>INVENTORY.CSV : " + trim(recprocessed) + "</h6>")
 fs.close
 set fs = nothing 
 set fso = nothing
 
 
else

	response.write("File NOT Read by System")
 	
end if





Set objConn55 = Server.CreateObject("ADODB.Connection")
objConn55.Open MM_overseaspr_STRING

sql="TRUNCATE TABLE dbo.english_names"
on error resume next
objConn55.Execute sql
sError = err.description
If len(sError) > 0 Then
	Response.Write("<h6 class='red'>DELETE ALL english_names error: " + sError + "</h6>")
	Response.End
End if
objConn55.Close 

recprocessed = 0

Set fso = Server.CreateObject("Scripting.FileSystemObject") 
set fs = fso.OpenTextFile(Server.MapPath("../apps/OVERSEAS SQL/Webdata/inventoryen.csv"), 1, true) 
Do Until fs.AtEndOfStream 
    lineData = fs.ReadLine
	lineData2 = lineData
    'inventory
	'field_1 char(20)
	'field_2 char(90)
	'field_3 float
	'field_4 float
	'QOH Integer
	'FamilyCat Integer
	'Precio4 float
	'Precio5 float
	
	IF len(lineData) > 0 THEN
	
		recprocessed = recprocessed + 1
		fieldno = 0
		Do While InStr(1, lineData, "," ,vbTextCompare) > 0
			fieldno = fieldno + 1
			position = InStr(1, lineData, "," ,vbTextCompare)
			if fieldno = 1 then
				' partno
				field_1 = mid(lineData,1,position-1)		
				lineData = mid(lineData,position+1)
			elseif fieldno = 2 then
				field_2 = mid(lineData,1,position-1)	
				linedata = ""		 		
			end if
		Loop


		Set oRS2 = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT COUNT(*) AS counter FROM dbo.english_names WHERE partno = '" + field_1 + "' " 
		oRS2.Open strSQL, MM_overseaspr_STRING
		If Not oRS2.EOF AND Not IsNull(oRS2.Fields.Item("counter")) then
			counter = trim(oRS2.Fields.Item("counter"))
		else
			counter = 0
		End if
		oRS2.Close
		Set oRS2 = Nothing

		Dim objConn6 
		Dim objRs2 
		
		if counter = 0 then
				
			Set objConn6 = Server.CreateObject("ADODB.Command") 
			'objConn6.ActiveConnection = "Driver={SQL Server};Server=GGG-LENOVO\EXPRESS2014;Database=overseas_prod;UID=sa;PWD=Akita9013"
			
			
			objConn6.ActiveConnection = MM_overseaspr_STRING
					
			objConn6.commandtext="INSERT into english_names(partno,description)values(?,?)"
			
			objConn6.Parameters(0) = field_1
			objConn6.Parameters(1) = field_2
			
			on error resume next
			set objRs2 = objConn6.execute
			sError = err.description
			If len(sError) > 0 Then
				Response.Write("<h6 class='red'>INSERT error english_names : " + sError + "</h6>")
			End if

		Else
		
			Set objConn6 = Server.CreateObject("ADODB.Command") 
			objConn6.ActiveConnection = MM_overseaspr_STRING
					
			objConn6.commandtext="UPDATE dbo.english_names SET description = ? WHERE partno = ?"
			
			objConn6.Parameters(0) = field_2
			objConn6.Parameters(1) = field_1
			
			on error resume next
			set objRs2 = objConn6.execute
			sError = err.description			
			If len(sError) > 0 Then
				Response.Write("<h6 class='red'>UPDATE error english_names : " + sError + "</h6>")
			End if

		End If
		
	End If		
	
Loop 
response.Write("<h6>INVENTORYEN.CSV : " + trim(recprocessed) + "</h6>")
fs.close: set fs = nothing 
%>
