<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" --> 
<%

Server.ScriptTimeout = 300  ' 5 minutes instead of default 90 seconds

Dim position
Dim fieldno
Dim field_1,field_2,field_3,field_4,field_5,field_6,field_7,field_8,field_9,field_10,field_11,field_12,field_13,field_14,field_15,field_16,field_17
Dim counter, sError, strSQL, oRS2, oRS5, fso, fs, rs

Dim lineData,lineData2
Dim recprocessed
recprocessed = 0

Dim nopricelevel
nopricelevel = 0
Set fso = Server.CreateObject("Scripting.FileSystemObject") 

' PROD server IIS works with MapPath 
'set fs = fso.OpenTextFile(Server.MapPath("../apps/OVERSEAS SQL/Webdata/customer.csv"), 1, true) 
' IIS Express local dev
set fs = fso.OpenTextFile("C:/apps/OVERSEAS SQL/Webdata/customer.csv", 1, true)
Do Until fs.AtEndOfStream 
    lineData = fs.ReadLine
	lineData2 = lineData
    'customer => users
	'1 user_auto_id char(6)
	'2 desc char(22)
	'3 pricetype integer
	'4 addr1 char(22)
	'5 addr2 char(22)
	'6 addr3 char(22)
	'7 salesman char(3)
	'8 term char(10)
	'9 fax char(14)
	'10 email char(150)
	'11 tel char(15)
	'12 discount char(4)
	'13 currentcredits char(8)
	'14 commchar char(1)
	'15 citytax char(1)
	'16 statetax char(1)

	'17 sameday char(1)  new column
	
	IF len(lineData) > 0 THEN
		
		recprocessed = recprocessed + 1
		fieldno = 0
		Do While InStr(1, lineData, "," ,vbTextCompare) > 0
			fieldno = fieldno + 1
			position = InStr(1, lineData, "," ,vbTextCompare)
			if fieldno = 1 then
				' custid
				field_1 = mid(lineData,1,position-1)						
			elseif fieldno = 2 then
				' desc
				field_2 = mid(lineData,1,position-1)
			elseif fieldno = 3 then
				' pricetype => user_level
				if Not IsNumeric(mid(lineData,1,position-1)) then
					field_3 = 0
					nopricelevel = nopricelevel + 1
				else	
					field_3 = CInt(mid(lineData,1,position-1)) 
				end if
			elseif fieldno = 4 then
				' addr1
				field_4 = trim(mid(lineData,1,position-1))
			elseif fieldno = 5 then
				' addr2
				field_5 = trim(mid(lineData,1,position-1))	 
			elseif fieldno = 6 then
				' addr3
				field_6 = trim(mid(lineData,1,position-1))
			elseif fieldno = 7 then
				' salesman
				field_7 = trim(mid(lineData,1,position-1))
			elseif fieldno = 8 then
				' term
				field_8 = trim(mid(lineData,1,position-1))	 
				field_8 = mid(field_8,1,10)
			elseif fieldno = 9 then
				' fax
				field_9 = trim(mid(lineData,1,position-1))	 
			elseif fieldno = 10 then
				' email
				field_10 = trim(mid(lineData,1,position-1))
			elseif fieldno = 11 then
				' tel
				field_11 = trim(mid(lineData,1,position-1))	 
			elseif fieldno = 12 then
				' discount
				if Not IsNumeric(mid(lineData,1,position-1)) then
					field_12 = 0
				else
					field_12 = CDbl(mid(lineData,1,position-1))
				end if
			elseif fieldno = 13 then
				' currentcredits
				if Not IsNumeric(mid(lineData,1,position-1)) then
					field_13 = 0
				else	
					field_13 = CDbl(mid(lineData,1,position-1))
				end if	
			elseif fieldno = 14 then
				' commchar
				field_14 = trim(mid(lineData,1,position-1))
				if trim(field_14) = "" then
					field_14 = "C"
				end if					
			elseif fieldno = 15 then
				' citytax
				field_15 = mid(lineData,1,1)
				if field_15 = "" or (field_15 <> "Y" AND field_15 <> "N") then
					field_15 = "Y"
				end if	

			elseif fieldno = 16 then
				' statetax
				field_16 = mid(lineData,1,1)
				if field_16 = "" or (field_16 <> "Y" AND field_16 <> "N") then
					field_16 = "Y"
				end if	

			else
				lineData = ""		 		
			end if
			
			if fieldno < 17 then
				lineData = mid(lineData,position+1)
			end if
		Loop
		
		' samedayflag (campo 17)
		field_17 = UCase(Trim(lineData))

		if field_17 = "" OR (field_17 <> "Y" AND field_17 <> "N") then
			Response.Write("<h6 class='red'>No valid samedayflag for item " & field_1 & "</h6>")
			field_17 = "Y"
		end if

		Dim objConn5 
		Dim objRs 
	
		Set oRS2 = Server.CreateObject("ADODB.Command")
		oRS2.ActiveConnection = MM_overseaspr_STRING
		oRS2.CommandText = "SELECT COUNT(*) AS counter FROM dbo.users WHERE user_auto_id = ?"
		oRS2.Parameters.Append oRS2.CreateParameter("@user_id", 200, 1, 6, field_1)
		Set rs = oRS2.Execute

		If Not rs.EOF AND Not IsNull(rs.Fields.Item("counter")) then
			counter = CInt(rs.Fields.Item("counter"))
		else
			counter = 0
		End if
		rs.Close
		Set rs = Nothing
		Set oRS2 = Nothing
		
		if counter = 0 then
			
			Set objConn5 = Server.CreateObject("ADODB.Command") 
			objConn5.ActiveConnection = MM_overseaspr_STRING		
			objConn5.commandtext="INSERT into users(user_auto_id,user_name,user_level,address,address2,city,salesman,term,fax,user_email,tel,discount,currentcredits,commchar,citytax,statetax,samedayflag,user_client_code,user_pwd,activestatus,lang,user_id)values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

			objConn5.Parameters.Append objConn5.CreateParameter("@p1", 200, 1, 6, field_1)      'user_auto_id
			objConn5.Parameters.Append objConn5.CreateParameter("@p2", 200, 1, 22, field_2)     'user_name
			objConn5.Parameters.Append objConn5.CreateParameter("@p3", 3, 1, 0, field_3)         'user_level (integer)
			objConn5.Parameters.Append objConn5.CreateParameter("@p4", 200, 1, 22, field_4)     'address
			objConn5.Parameters.Append objConn5.CreateParameter("@p5", 200, 1, 22, field_5)     'address2
			objConn5.Parameters.Append objConn5.CreateParameter("@p6", 200, 1, 22, field_6)     'city
			objConn5.Parameters.Append objConn5.CreateParameter("@p7", 200, 1, 3, field_7)      'salesman
			objConn5.Parameters.Append objConn5.CreateParameter("@p8", 200, 1, 10, field_8)     'term
			
			' Sanitize fax - force to string and truncate
			field_9 = CStr(field_9 & "")
			If Len(field_9) > 14 Then
				field_9 = Left(field_9, 14)
			End If
			objConn5.Parameters.Append objConn5.CreateParameter("@p9", 200, 1, 14, field_9)     'fax
			
			objConn5.Parameters.Append objConn5.CreateParameter("@p10", 200, 1, 150, field_10)  'user_email
			objConn5.Parameters.Append objConn5.CreateParameter("@p11", 200, 1, 15, field_11)   'tel
			objConn5.Parameters.Append objConn5.CreateParameter("@p12", 5, 1, 0, field_12)       'discount (double)
			objConn5.Parameters.Append objConn5.CreateParameter("@p13", 5, 1, 0, field_13)       'currentcredits (double)
			objConn5.Parameters.Append objConn5.CreateParameter("@p14", 200, 1, 1, field_14)    'commchar
			objConn5.Parameters.Append objConn5.CreateParameter("@p15", 200, 1, 1, field_15)    'citytax
			objConn5.Parameters.Append objConn5.CreateParameter("@p16", 200, 1, 1, field_16)    'statetax
			objConn5.Parameters.Append objConn5.CreateParameter("@p17", 200, 1, 1, field_17)    'samedayflag
			objConn5.Parameters.Append objConn5.CreateParameter("@p18", 200, 1, 6, field_1)     'user_client_code
			objConn5.Parameters.Append objConn5.CreateParameter("@p19", 200, 1, 5, "11111")     'user_pwd
			objConn5.Parameters.Append objConn5.CreateParameter("@p20", 200, 1, 1, "A")         'activestatus
			objConn5.Parameters.Append objConn5.CreateParameter("@p21", 200, 1, 1, "S")         'lang
			objConn5.Parameters.Append objConn5.CreateParameter("@p22", 200, 1, 6, field_1)     'user_id
			
			On Error Resume Next
			set objRs = objConn5.execute
			sError = Err.Description
			On Error Goto 0

			If len(sError) > 0 Then
				Response.Write("<h6 class='red'>Insert:&nbsp;&nbsp;" & trim(field_1) & "&nbsp;&nbsp;" & sError & "</h6>" )
			End if
			
			Set objRs = Nothing
			Set objConn5 = Nothing

		Else
			
			Set objConn5 = Server.CreateObject("ADODB.Command") 
			objConn5.ActiveConnection = MM_overseaspr_STRING		
			objConn5.commandtext="UPDATE dbo.users SET user_name = ?,user_level = ?,address = ?,address2 = ?,city = ?,salesman = ?,term = ?,fax = ?,user_email = ?,tel = ?,discount = ?,currentcredits = ?,commchar = ?,citytax = ?,statetax = ?,samedayflag = ?,user_client_code = ?,user_id = ?,activestatus = ? WHERE user_auto_id = ?"

			objConn5.Parameters.Append objConn5.CreateParameter("@p1", 200, 1, 22, field_2)     'user_name
			objConn5.Parameters.Append objConn5.CreateParameter("@p2", 3, 1, 0, field_3)         'user_level
			objConn5.Parameters.Append objConn5.CreateParameter("@p3", 200, 1, 22, field_4)     'address
			objConn5.Parameters.Append objConn5.CreateParameter("@p4", 200, 1, 22, field_5)     'address2
			objConn5.Parameters.Append objConn5.CreateParameter("@p5", 200, 1, 22, field_6)     'city
			objConn5.Parameters.Append objConn5.CreateParameter("@p6", 200, 1, 3, field_7)      'salesman
			objConn5.Parameters.Append objConn5.CreateParameter("@p7", 200, 1, 10, field_8)     'term

			' Sanitize fax - force to string and truncate  
			field_9 = CStr(field_9 & "")
			If Len(field_9) > 14 Then
				field_9 = Left(field_9, 14)
			End If
			objConn5.Parameters.Append objConn5.CreateParameter("@p8", 200, 1, 14, field_9)     'fax
			
			objConn5.Parameters.Append objConn5.CreateParameter("@p9", 200, 1, 150, field_10)   'user_email
			objConn5.Parameters.Append objConn5.CreateParameter("@p10", 200, 1, 15, field_11)   'tel
			objConn5.Parameters.Append objConn5.CreateParameter("@p11", 5, 1, 0, field_12)       'discount
			objConn5.Parameters.Append objConn5.CreateParameter("@p12", 5, 1, 0, field_13)       'currentcredits
			objConn5.Parameters.Append objConn5.CreateParameter("@p13", 200, 1, 1, field_14)    'commchar
			objConn5.Parameters.Append objConn5.CreateParameter("@p14", 200, 1, 1, field_15)    'citytax
			objConn5.Parameters.Append objConn5.CreateParameter("@p15", 200, 1, 1, field_16)    'statetax
			objConn5.Parameters.Append objConn5.CreateParameter("@p16", 200, 1, 1, field_17)    'samedayflag
			objConn5.Parameters.Append objConn5.CreateParameter("@p17", 200, 1, 6, field_1)     'user_client_code
			objConn5.Parameters.Append objConn5.CreateParameter("@p18", 200, 1, 6, field_1)     'user_id
			objConn5.Parameters.Append objConn5.CreateParameter("@p19", 200, 1, 1, "A")         'activestatus
			objConn5.Parameters.Append objConn5.CreateParameter("@p20", 200, 1, 6, field_1)     'user_auto_id (WHERE clause)
			
			On Error Resume Next
			set objRs = objConn5.execute
			sError = Err.Description
			On Error Goto 0

			If len(sError) > 0 Then
				Response.Write("<h6 class='red'>Update:&nbsp;&nbsp;" & trim(field_1) & "&nbsp;&nbsp;" & sError & "</h6>" )
			End if
		
			
			Set objRs = Nothing
			Set objConn5 = Nothing
    
		End IF  ' This closes "if counter = 0"
    
	End If  ' This closes "IF len(lineData) > 0"

Loop 
response.Write("<h6>CUSTOMER.CSV : " + trim(recprocessed) + "</h6>")

if nopricelevel > 0 then
	response.Write("<h6 class='red' >No price level assigned : " + trim(nopricelevel) + "</h6>")

	Set oRS5 = Server.CreateObject("ADODB.Recordset")	
	strSQL = "SELECT user_auto_id,user_name FROM dbo.users WHERE user_level = 0 "	
	oRS5.Open strSQL, MM_overseaspr_STRING
	Do While Not oRS5.EOF
	
		response.write("<h6 class='red'> " + trim(oRS5("user_auto_id")) + " - " + trim(oRS5("user_name")) + "  No price level! </h6>")
				
		oRS5.MoveNext
	loop
	oRS5.close
	
end if

fs.close: set fs = nothing 
%>
