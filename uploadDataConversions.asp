<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" --> 
<%

Server.ScriptTimeout = 3600
Dim recprocessed
recprocessed = 0

Set oRS = Server.CreateObject("ADODB.Recordset")	
strSQL = "SELECT * FROM dbo.similar "	
oRS.Open strSQL, MM_overseaspr_STRING
Do While Not oRS.EOF
	
	sreplacement = trim(oRS("replacement"))
	spartid = trim(oRS("partid"))
	recprocessed = recprocessed + 1		
	Set oRS2 = Server.CreateObject("ADODB.Connection")
	oRS2.Open MM_overseaspr_STRING
	sql="UPDATE dbo.partmst1_distinct SET replacement = '" + sreplacement + "' WHERE field_1 = '" + spartid + "' " 
	oRS2.Execute sql
	sError = err.description
	oRS2.Close 
	
	If len(sError) > 0 Then
		Response.Write("<h6 class='red'>replacements error : " & sError & "</h6>" )
	End if
		
	oRS.MoveNext
Loop
response.Write("<h6>replacements : " + trim(recprocessed) + "</h6>")

recprocessed = 0
Set oRS = Server.CreateObject("ADODB.Recordset")	
strSQL = "SELECT * FROM dbo.familicat "	
oRS.Open strSQL, MM_overseaspr_STRING
Do While Not oRS.EOF
	
	sfamdesc 	= trim(oRS("field_1"))
	if not isnull(oRS("english_desc")) then
		sengfamdesc = trim(oRS("english_desc"))	
	else
		sengfamdesc = ""	
	end if
	sfam = trim(oRS("field_2"))
	recprocessed = recprocessed + 1
	Set oRS2 = Server.CreateObject("ADODB.Connection")
	oRS2.Open MM_overseaspr_STRING
	sql="UPDATE dbo.partmst1_distinct SET familia_descripcion = '" + sfamdesc + "', family_description = '" + sengfamdesc + "' WHERE field_6 = '" + sfam + "' " 
	oRS2.Execute sql
	sError = err.description
	oRS2.Close 
	
	If len(sError) > 0 Then
		Response.Write("<h6 class='red'>family descriptions error : " & sError & "</h6>" )
	End if
				
	oRS.MoveNext
Loop
response.Write("<h6>family descriptions : " + trim(recprocessed) + "</h6>")

recprocessed = 0
Set oRS = Server.CreateObject("ADODB.Recordset")	
strSQL = "SELECT * FROM dbo.long1 "	
oRS.Open strSQL, MM_overseaspr_STRING
Do While Not oRS.EOF
	
	sfam_make_item 	= trim(oRS("fam_make_item"))
	sitem_no = trim(oRS("item_no"))
	recprocessed = recprocessed + 1
		
	Set oRS2 = Server.CreateObject("ADODB.Connection")
	oRS2.Open MM_overseaspr_STRING
	sql="UPDATE dbo.partmst1_distinct SET fam_make_item = '" + sfam_make_item + "' WHERE field_1 = '" + sitem_no + "' " 
	oRS2.Execute sql
	sError = err.description
	oRS2.Close 
	
	If len(sError) > 0 Then
		Response.Write("<h6 class='red'>fam_make_item ORDERING error : " & sError & "</h6>")
	End if
			
	oRS.MoveNext
Loop
response.Write("<h6>fam_make_item ORDERING: " + trim(recprocessed) + "</h6>")



recprocessed = 0
Set oRS = Server.CreateObject("ADODB.Recordset")	
strSQL = "SELECT partno,description FROM dbo.english_names "	
oRS.Open strSQL, MM_overseaspr_STRING
Do While Not oRS.EOF
	
	sdescription = trim(oRS("description"))
	spartno = trim(oRS("partno"))
	recprocessed = recprocessed + 1
	Set oRS2 = Server.CreateObject("ADODB.Command") 
	oRS2.ActiveConnection = MM_overseaspr_STRING		
	oRS2.commandtext="UPDATE dbo.partmst1_distinct SET english_version = ? WHERE field_1 = ?"	
	oRS2.Parameters(0) = sdescription
	oRS2.Parameters(1) = spartno	
	set objRs2 = oRS2.execute
	sError = err.description
	'objRs2.Close 
	
	If len(sError) > 0 Then
		Response.Write("<h6 class='red'>Part english description error : " & sError & "</h6>" )
	End if
			
	oRS.MoveNext
Loop
response.Write("<h6>Part english description : " + trim(recprocessed) + "</h6>")

recprocessed = 0
Set oRS3 = Server.CreateObject("ADODB.Recordset")	
strSQL = "SELECT userid from dbo.users_lockedout WHERE activestatus = 'I' " 
oRS3.Open strSQL, MM_overseaspr_STRING
Do While Not oRS3.EOF
	
	suserid = trim(oRS3("userid"))
	recprocessed = recprocessed + 1	
	if trim(GetUserName(suserid)) <> "N/A" then
		susername = susername + GetUserName(suserid) + "<br>"
	else
		susername = susername + suserid + "<br>"
	end if
	Set oRS4 = Server.CreateObject("ADODB.Command") 
	oRS4.ActiveConnection = MM_overseaspr_STRING		
	oRS4.commandtext="UPDATE dbo.users SET activestatus = 'I' WHERE user_auto_id = ? "	
	oRS4.Parameters(0) = suserid
	set objRs2 = oRS4.execute
	sError = err.description
	'objRs2.Close 
	
	If len(sError) > 0 Then
		Response.Write("<h6 class='red' >Users locked OUT error : " & sError & "</h6>")
	End if	
			
	oRS3.MoveNext
Loop
response.Write("<h6 class='red'>Users locked OUT : " + trim(recprocessed) + "<br>" + susername + "<br></h6>")
oRS3.close



'lang = null set to S spanish
Set oRS = Server.CreateObject("ADODB.Command") 
oRS.ActiveConnection = MM_overseaspr_STRING		
oRS.commandtext="UPDATE dbo.users SET lang = 'S' WHERE lang is null "	
set objRs = oRS.execute
sError = err.description
If len(sError) > 0 Then
	Response.Write("<h6 class='red' >English users error : " & sError & "</h6>")
End if	
	

recprocessed = 0
Set oRS4 = Server.CreateObject("ADODB.Recordset")	
strSQL = "SELECT user_id from dbo.english_users " 
oRS4.Open strSQL, MM_overseaspr_STRING
Do While Not oRS4.EOF
	
	suserid = trim(oRS4("user_id"))
	recprocessed = recprocessed + 1	
	if trim(GetUserName(suserid)) <> "N/A" then
		susername = susername + GetUserName(suserid) + "<br>"
	else
		susername = susername + suserid + "<br>"
	end if
	Set oRS5 = Server.CreateObject("ADODB.Command") 
	oRS5.ActiveConnection = MM_overseaspr_STRING		
	oRS5.commandtext="UPDATE dbo.users SET lang = 'E' WHERE user_auto_id = ? "	
	oRS5.Parameters(0) = suserid
	set objRs2 = oRS5.execute
	sError = err.description
	'objRs2.Close 
	
	If len(sError) > 0 Then
		Response.Write("<h6 class='red' >English users error : " & sError & "</h6>")
	End if	
			
	oRS4.MoveNext
Loop
response.Write("<h6 class='red'>English users : " + trim(recprocessed) + "<br>" + susername + "<br></h6>")
oRS4.close



' Disable website under maintenance
Set oRS = Server.CreateObject("ADODB.Command") 
oRS.ActiveConnection = MM_overseaspr_STRING	
oRS.commandtext="UPDATE dbo.counters SET forceinventoryupd = ? "
oRS.Parameters(0) = "0"
on error resume next
set oRecordSet = oRS.execute
sError = err.description			
If len(sError) > 0 Then
	Response.Write("<h6 class='red'>System error</h6>")
End if
close oRecordSet
set oRecordSet = nothing	




%>
