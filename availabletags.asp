<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
Response.Expires = -1
Server.ScriptTimeout = 300
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="index.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<% 
dim vparm , vchoices

vparm = request("term")
If LEN(vparm) > 0 Then



	Set keywords = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT distinct field_1 FROM dbo.partmst1_distinct WHERE field_1 like '" + vparm + "%' Order by field_1 ;"
	keywords.Open strSQL, MM_overseaspr_STRING
	
	If Not keywords.EOF then
		do while not keywords.eof
			vchoices = vchoices & "" & keywords("field_1") & ""","""
			keywords.movenext
		loop
	end if		
	keywords.close
	Set keywords = Nothing   

	vchoices = Replace(vchoices,""","""",""","")
	vchoices = Left(vchoices,Len(vchoices)-2)
	'vchoices = """["""&vchoices &"]"""
	vchoices = "["""&vchoices &"]"
	
	response.write vchoices

	

	
	
'	output = "["
'
'	While (NOT keywords.EOF) 
'    	output = output & "{""id"":""" & keywords.Fields.item("field_1") & """,""value"":""" & keywords.Fields.Item("field_1") & """},"
'     	keywords.MoveNext()
'	Wend
'
'	keywords.Close()
'	Set keywords = Nothing
'
'	output=Left(output,Len(output)-1)
'	output = output & "]"
'	
'	response.write output

End If


%>