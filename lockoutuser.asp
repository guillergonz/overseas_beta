<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%


If GetSecureVal(Request.Querystring("id") <> "") Then
	user_auto_id = GetSecureVal(Request.Querystring("id"))
	action 	= GetSecureVal(Request.Querystring("action"))
	check = GetSecureVal(Request.Querystring("check"))
	
	if LEN(user_auto_id) > 0 AND check = "S" then
	
		if action <> "1" then
			check = "E"
		end if
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open MM_overseaspr_STRING
		sql="UPDATE dbo.users SET lang = '" + check + "' WHERE user_auto_id = '" + CStr(user_auto_id) + "'" 
		objConn.Execute sql
		sError = err.description
		If len(sError) > 0 Then
			Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
		End if
		objConn.Close
		Set objConn = Nothing
		
		response.redirect("usuarios.asp")
		
	elseif LEN(user_auto_id) > 0 AND check = "E" then 
		 
		if action <> "1" then
			check = "S"
		end if	
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open MM_overseaspr_STRING
		sql="UPDATE dbo.users SET lang = '" + check + "' WHERE user_auto_id = '" + CStr(user_auto_id) + "'" 
		objConn.Execute sql
		sError = err.description
		If len(sError) > 0 Then
			Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
		End if
		objConn.Close
		Set objConn = Nothing
		 
		response.redirect("usuarios.asp") 
		 
	' A / I
	
	ElseIf LEN(user_auto_id) > 0 AND (action = "A" or action = "I") Then 
		
					
		Set oRS = Server.CreateObject("ADODB.Recordset")
		strSQL="SELECT user_id from dbo.users WHERE user_auto_id = '" + CStr(user_auto_id) + "'" 
		oRS.Open strSQL, MM_overseaspr_STRING
		If Not oRS.EOF Then
			If Not IsNull(oRS.Fields.Item("user_id")) Then
				
				user_id = oRS("user_id")
				
				Set oRS = Server.CreateObject("ADODB.Recordset")
				strSQL="SELECT userid from dbo.users_lockedout WHERE userid = '" + CStr(user_id) + "'" 
				oRS.Open strSQL, MM_overseaspr_STRING
				If Not oRS.EOF Then
					If Not IsNull(oRS.Fields.Item("userid")) Then
						' UPDATE
						
						Set objConn = Server.CreateObject("ADODB.Connection")
						objConn.Open MM_overseaspr_STRING
						sql="UPDATE dbo.users SET activestatus = '" + action + "' WHERE user_auto_id = '" + CStr(user_auto_id) + "'" 
						objConn.Execute sql
						sError = err.description
						If len(sError) > 0 Then
							Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
						End if
						objConn.Close
						Set objConn = Nothing
						

						Set objConn = Server.CreateObject("ADODB.Connection")
						objConn.Open MM_overseaspr_STRING
						
						if action = "I" then
							sql="UPDATE dbo.users_lockedout SET activestatus = '" + action + "' WHERE userid = '" + CStr(user_id) + "'" 
						elseif action = "A" then
							sql="DELETE FROM dbo.users_lockedout WHERE userid = '" + CStr(user_id) + "'" 
						end if
	
						objConn.Execute sql
						sError = err.description
						If len(sError) > 0 Then
							Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
						End if
						objConn.Close
						Set objConn = Nothing
					
					End If				
				
				Else
					' INSERT
				
					Set objConn = Server.CreateObject("ADODB.Connection")
					objConn.Open MM_overseaspr_STRING
					sql="INSERT INTO dbo.users_lockedout (userid,activestatus) VALUES ('" + user_id + "','" + action + "')" 
					objConn.Execute sql
					sError = err.description
					If len(sError) > 0 Then
						Response.Write("Ocurrió el siguiente error : <br><br>" & sError )
					End if
					objConn.Close
					Set objConn = Nothing
						
				End IF			
					
				response.redirect("usuarios.asp")

			End If
		End If	
				
	End If

else
	
	response.redirect("index.asp")
	
end if	
%>
