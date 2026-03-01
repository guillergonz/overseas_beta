<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%

Response.Expires = -1
Response.addHeader "Cache-Control", "max-age=0,no-cache,no-store"
Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8"

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
'--------------------------------------------------------------------------
' THE PURPOSE OF THIS PAGE IS:
'
' 1) SUBMIT THE CART FOR A NEW ORDER BASED ON THE PARTS AVAILABLES IN USER_CART
' 2) CREATE A NEW ORDER WITH A NEW NUMBER BASED ON THE NEXT SEQUENCE AVAILABLE ON THE TABLE NEXT_ORDER_NUMBER.
' 3) SUBSTRACT QOH FOR EACH ORDERED ITEM IN PARTMST1_DISTINCT BASED ON ORDERED QUANTITY.
'
'--------------------------------------------------------------------------

' 1) SUBMIT ORDER FUNCTION
' 2) SUBSTRACK QOH FUNCTION
' 3) EMAIL NOTIFICATION FUNCTION

delivery_type = Request("delivery_type")

instructions = trim(Request("instructions"))

' response.write "DT " + delivery_type + "<br>"
' response.write "I  " + instructions + "<br>"

OrderSubmit delivery_type,instructions

'Process order fftemp

Set oRS = Server.CreateObject("ADODB.Recordset")	
strSQL = "SELECT order_number,order_user,order_date, count(*) as lines,sum(order_qty) as tot_item, sum(order_qty * item_price) as sale FROM dbo.clients_orders WHERE order_status <> 'P' and order_status <> 'H' GROUP BY order_number,order_user,order_date; "	
oRS.Open strSQL, MM_overseaspr_STRING
Do While Not oRS.EOF
	
	sorder_number 	= trim(oRS("order_number"))
	suserid 		= UCase(trim(oRS("order_user")))
	sorderdate 		= CDate(oRS("order_date"))
	ilines 			= CInt(trim(oRS("lines")))
	itotqty 		= CInt(trim(oRS("tot_item")))	
	dtotsale	 	= Cdbl(trim(oRS("sale")))
	
	Set oRS2 = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT order_qty,item_price,order_type FROM dbo.clients_orders WHERE order_id = ( SELECT max(order_id) from dbo.clients_orders where order_number = '" + trim(sorder_number) + "' ) "
	oRS2.Open strSQL, MM_overseaspr_STRING
	if not oRS2.eof then
	
		iqty = oRS2("order_qty")
		dprice = oRS2("item_price")
		sorder_type = oRS2("order_type")
				
		if len(trim(sorder_number)) = 1 then
			sordernumber = "0000" + trim(sorder_number)
		elseif len(trim(sorder_number)) = 2 then
			sordernumber = "000" + trim(sorder_number)
		elseif len(trim(sorder_number)) = 3 then
			sordernumber = "00" + trim(sorder_number)
		elseif len(trim(sorder_number)) = 4 then
			sordernumber = "0" + trim(sorder_number)
		else
			sordernumber = trim(sorder_number)
		end if			
	
		
		
		Set oRS3 = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT count(*) as icount FROM dbo.clients_orders WHERE order_number = '" & trim(sorder_number) & "' AND order_part in ( SELECT specials FROM dbo.prespecials )"
		oRS3.Open sql, MM_overseaspr_STRING
		if not oRS3.eof then
			icount = CInt(oRS3("icount"))
		else
			icount = 0	
		end if
		oRS3.close
		SET oRS3 = nothing
		
		if icount > 0 then
			sordernumber = "I" + sordernumber + "E"
		else	
			sordernumber = "I" + sordernumber + "W"		
		end if	
		sinvno = mid(sordernumber,2,5)
	
		Set oRS4 = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT fax,salesman,term,discount,user_name,address,address2,city,commchar,citytax,statetax,samedayflag FROM dbo.users WHERE user_id = '" + suserid + "' "
		oRS4.Open sql, MM_overseaspr_STRING
		if not oRS4.eof then
					
			if IsNull(oRS4("fax")) then
				sfax = ""
			else
				sfax = trim(oRS4("fax"))  
			end if	
			
			if IsNull(oRS4("salesman")) then
				ssalesman = ""
			else
				ssalesman = mid(oRS4("salesman"),1,2)  
			end if	
			 
         	if IsNull(oRS4("term")) then
				sterm = ""
			else
				sterm = trim(oRS4("term"))  
			end if
			
			if IsNull(oRS4("discount")) then
				sdisc = ""
			else
				sdisc = trim(oRS4("discount"))  
			end if	 
			
			suser_name = trim(oRS4("user_name"))
			if InStr(suser_name,"'") > 0 then
				suser_name = Replace(suser_name,"'","")
			end if	
			if LEN(suser_name) > 24 then
				suser_name = mid(suser_name,1,24)
			end if	
			
			if IsNull(oRS4("address")) then
				sadd1 = ""
			else
				sadd1 = trim(oRS4("address"))  
			end if	 
			if LEN(sadd1) > 24 then
				sadd1 = mid(sadd1,1,24)
			end if	
			
			if IsNull(oRS4("address2")) then
				sadd2 = ""
			else
				sadd2 = trim(oRS4("address2"))  
			end if	 
			
			if IsNull(oRS4("city")) then
				scity = ""
			else
				scity = trim(oRS4("city"))  
			end if	 
			if LEN(scity) > 24 then
				scity = mid(scity,1,24)
			end if	
			
			if IsNull(oRS4("commchar")) then
				scommchar = ""
			else
				scommchar = trim(oRS4("commchar"))  
			end if	
				
			if scommchar = "R" or scommchar = "E" then
				scommchar = "6.0"
			else
				scommchar = "7.0"
			end if
			' 1 char field		
			sline = mid(ilines,1,1)
			stotqty = trim(itotqty) 
			stotsale = trim(dtotsale)
			
			if IsNull(oRS4("citytax")) then
				scitytax = ""
			else
				scitytax = trim(oRS4("citytax"))  
			end if	
			
			if IsNull(oRS4("statetax")) then
				sstatetax = ""
			else
				sstatetax = trim(oRS4("statetax"))  
			end if	
				
			if scitytax = "" then scitytax = "Y"
			if sstatetax = "" then sstatetax = "Y"
			
			if scitytax = "Y" then
				dcitytax = FormatNumber(dtotsale * .01,2)
			end if

			if sstatetax = "Y" then
				dstatetax = FormatNumber(dtotsale * .105,2)
			end if
		
			
			seffdate = Cdate(now())
			seffdate = mid(seffdate,1,InStr(seffdate," "))
			
			if IsNull(sadd2) then sadd2 = ""
			if IsNull(scity) then scity = ""
			
						
			Set oRS5 = Server.CreateObject("ADODB.Connection")
			oRS5.Open MM_overseaspr_STRING
			sql="DELETE FROM dbo.FFTEM1 WHERE ord_number = '" + sordernumber + "' "
			on error resume next
			oRS5.Execute sql
			sError = err.description
			If len(sError) > 0 Then
				Response.Write("<h5 style='color:red'>DELETE FFTEM1 : " & trim(sorder_number) & "&nbsp;&nbsp;" & sError & "</h6>" )
				
			End if
	
			Set oRS5 = Server.CreateObject("ADODB.Connection")
			oRS5.Open MM_overseaspr_STRING
			
			sql="INSERT into FFTEM1 ( customer_no, ord_number, cust_name,addr1,addr2,addr3,salesperson,bags,sec,order_date,po_no,pay_term,invno,num_lines,tot_item_qty,total_sale,delivered,routing,citytax,statetax) VALUES ('" + trim(suserid) + "','" + trim(sordernumber) + "','" + trim(suser_name) + "','" + trim(add1) + "','" + trim(sadd2) + "','" + trim(scity) + "','" + trim(ssalesman) + "',1,'','" + trim(seffdate) + "',0,'" + trim(sterm) + "','" + trim(sinvno) + "','" + trim(sline) + "','" + trim(stotqty) + "','" + trim(stotsale) + "','" + trim(sorder_type) + "','" + trim(sdisc) + "','" + trim(dcitytax) + "','" + trim(dstatetax) + "' ) "
			
			on error resume next
			oRS5.Execute sql
			sError = err.description
			If len(sError) > 0 Then
				'Response.Write("<h5 style='color:red'>Insert FFTEM1 : " & trim(sorder_number) & "&nbsp;Order Number&nbsp;" & trim(sinvno) & "&nbsp;Invoice Number&nbsp;"  & sError & "</h6>" )				
			End if
			oRS5.close
			SET oRS5 = nothing

			Set oRS6 = Server.CreateObject("ADODB.Connection")
			oRS6.Open MM_overseaspr_STRING
			sql="DELETE FROM dbo.FFTEM2 WHERE invoice = '" + trim(sinvno) + "' "
			on error resume next
			oRS6.Execute sql
			sError = err.description
			If len(sError) > 0 Then
				Response.Write("<h6 class='red'>DELETE FFTEM2 : " & trim(sinvno) & "&nbsp;&nbsp;" & sError & "</h6>" )
				
			End if
			oRS6.close
			SET oRS6 = nothing
			
			Dim iline
			iline = 1	
			Set oRS5 = Server.CreateObject("ADODB.Recordset")	
			strSQL = "SELECT order_user,order_date,order_part,order_qty,item_price,comments FROM dbo.clients_orders WHERE order_number = '" + trim(sorder_number) + "' "	
			oRS5.Open strSQL, MM_overseaspr_STRING
			Do While Not oRS5.EOF
	
				spart 		= trim(oRS5("order_part"))
				sordqty 	= trim(oRS5("order_qty"))
				sitemprice 	= trim(oRS5("item_price")) 
				scomments 	= trim(oRS5("comments"))
			
				scomments = Replace(scomments, vbCrLf, "")
				scomments = Replace(scomments, vbLf, "")
				scomments = Replace(scomments, vbTab, "")
				scomments = Replace(scomments, "'", "") 
				
				if spart = "ZZZNOF" then
					spartdesc = scomments
					if len(spartdesc) > 25 then
						spartdesc = mid(spartdesc,1,25)
					end if	
				else	
				
					Set oRS6 = Server.CreateObject("ADODB.Recordset")
					sql = "SELECT field_2 FROM dbo.partmst1_distinct WHERE field_1 = '" + spart + "' "
					oRS6.Open sql, MM_overseaspr_STRING
					if not oRS6.eof then
						spartdesc = mid(oRS6("field_2"),1,25)
					else
						spartdesc = "N/A"
					end if
					oRS6.close
					SET oRS6 = nothing
						
				end if	
				
				Set oRS6 = Server.CreateObject("ADODB.Command") 
				'oRS6.ActiveConnection = "Driver={SQL Server};Server=GGG-LENOVO\EXPRESS2014;Database=overseas_prod;UID=sa;PWD=Akita9013"		
				
				oRS6.ActiveConnection = MM_overseaspr_STRING
				oRS6.commandtext="INSERT into FFTEM2 ( invoice,line_no,item,ordqty,price, descrip1,commission) VALUES (?,?,?,?,?,?,?) "
				
				oRS6.Parameters(0) = sinvno
				oRS6.Parameters(1) = iline
				oRS6.Parameters(2) = spart
				oRS6.Parameters(3) = sordqty
				oRS6.Parameters(4) = sitemprice
				oRS6.Parameters(5) = spartdesc
				oRS6.Parameters(6) = scommchar
				on error resume next
				set objRs = oRS6.execute
				sError = err.description
	
				If len(sError) > 0 Then
					Response.Write("<h5 style='color:red'>INSERT FFTEM2 : " & sordernumber & "  " & trim(spart) & " " & sError & "</h6>" )
				End if
				oRS6.close	
				SET oRS6 = nothing
				
				iline = iline + 1
				
	
				oRS5.MoveNext
			Loop
			oRS5.close	
			SET oRS5 = nothing
	
			Set oRS5 = Server.CreateObject("ADODB.Connection")
			oRS5.Open MM_overseaspr_STRING
			sql="DELETE FROM dbo.FFTEM3 WHERE ord_number = '" + sordernumber + "' "
			on error resume next
			oRS5.Execute sql
			sError = err.description
			If len(sError) > 0 Then
				Response.Write("<h5 style='color:red'>DELETE FFTEM3 : " & sordernumber & "  " & sError & "</h6>" )
				
			End if
			oRS5.close
			SET oRS5 = nothing
			
			Set oRS6 = Server.CreateObject("ADODB.Command") 
			oRS6.ActiveConnection = MM_overseaspr_STRING		
			oRS6.commandtext="INSERT into dbo.FFTEM3 ( invno,ord_number,customer_no,order_stat,commchar) VALUES (?,?,?,?,?) "			
			oRS6.Parameters(0) = sinvno
			oRS6.Parameters(1) = sordernumber
			oRS6.Parameters(2) = suserid
			oRS6.Parameters(3) = "OK"
			oRS6.Parameters(4) = mid(scommchar,1,1)
			set objRs = oRS6.execute
			sError = err.description
			
			
			
			If len(sError) > 0 Then
				Response.Write("<h5 style='color:red'>INSERT FFTEM3" & trim(sinvno) & sError & "</h6>" )
			End if
			objRs.close
			SET objRs = nothing
			
			oRS6.close	
			Set oRS6 = nothing
				
			Set oRS5 = Server.CreateObject("ADODB.Connection")
			oRS5.Open MM_overseaspr_STRING
			sql="DELETE FROM dbo.FFTEM4 WHERE ord_number = '" + sordernumber + "' "
			on error resume next
			oRS5.Execute sql
			sError = err.description
			If len(sError) > 0 Then
				Response.Write("<h6 class='red'>DELETE FFTEM4 : " & sordernumber & "&nbsp;&nbsp;" & sError & "</h6>" )
			End if
			oRS5.close
			Set oRS5 = nothing
	
			Dim filename
			Dim wrote1,wrote2,wrote3
			wrote1 = false
			wrote2 = false
			wrote3 = false
			
			
			
			filename	= "fftem1" + sordernumber + ".csv"
			Set fs=Server.CreateObject("Scripting.FileSystemObject")
			Set fname=fs.CreateTextFile("C:\apps\zorrilla\webords\rdata\" & filename  ,true)
			
			Set RecordsetFFTEM1 = Server.CreateObject("ADODB.Recordset")	
			strSQL = "SELECT customer_no,ord_number,cust_name, addr1,addr2,addr3,salesperson,bags,sec,order_date,po_no,pay_term,invno,num_lines,tot_item_qty,total_sale,delivered,routing,citytax,statetax FROM dbo.FFTEM1 WHERE ord_number = '" + sordernumber + "' "	
			RecordsetFFTEM1.Open strSQL, MM_overseaspr_STRING
			Do While Not RecordsetFFTEM1.EOF
			
				'  & "," & trim(RecordsetFFTEM1("routing"))  NOT USED
				fname.WriteLine( trim(RecordsetFFTEM1("customer_no")) & "," & trim(RecordsetFFTEM1("ord_number")) & "," & trim(RecordsetFFTEM1("cust_name")) & "," & trim(RecordsetFFTEM1("addr1")) & "," & trim(RecordsetFFTEM1("addr2")) & "," & trim(RecordsetFFTEM1("addr3")) & "," & trim(RecordsetFFTEM1("salesperson")) & "," & trim(RecordsetFFTEM1("bags")) & "," & trim(RecordsetFFTEM1("sec")) & "," & trim(RecordsetFFTEM1("order_date")) & "," & trim(RecordsetFFTEM1("po_no")) & "," & trim(RecordsetFFTEM1("pay_term")) & "," & trim(RecordsetFFTEM1("invno")) & "," & trim(RecordsetFFTEM1("num_lines")) & "," & trim(RecordsetFFTEM1("tot_item_qty")) & "," & trim(RecordsetFFTEM1("total_sale")) & "," & trim(RecordsetFFTEM1("delivered")) & "," & trim(RecordsetFFTEM1("citytax")) & "," & trim(RecordsetFFTEM1("statetax")) )
				
				wrote1 = true
				
				RecordsetFFTEM1.MoveNext
			loop
			RecordsetFFTEM1.close
			SET RecordsetFFTEM1 = Nothing
			
			set fname=nothing
			set fs=nothing
			
			
			
			
			
			filename	= "fftem2" + sordernumber + ".csv"
			Set fs=Server.CreateObject("Scripting.FileSystemObject")
			Set fname=fs.CreateTextFile("C:\apps\zorrilla\webords\rdata\" & filename ,true)
			
			Set RecordsetFFTEM2 = Server.CreateObject("ADODB.Recordset")	
			strSQL = "SELECT invoice,line_no,item,ordqty,price, descrip1,commission FROM dbo.FFTEM2 WHERE invoice = '" + sinvno + "' "	
			RecordsetFFTEM2.Open strSQL, MM_overseaspr_STRING
			Do While Not RecordsetFFTEM2.EOF
			
				' LEFT(strOrdNo & Space(8), 8)

				fname.WriteLine( trim(RecordsetFFTEM2("invoice")) & "," & trim(RecordsetFFTEM2("line_no")) & "," & LEFT(RecordsetFFTEM2("item") & Space(15),15) & "," & trim(RecordsetFFTEM2("descrip1")) & "," & trim(RecordsetFFTEM2("ordqty")) & "," & trim(RecordsetFFTEM2("price")) & "," & "," & "," & trim(RecordsetFFTEM2("commission")) )
				
				wrote2 = true
				
				RecordsetFFTEM2.MoveNext
			loop
			RecordsetFFTEM2.close
			Set RecordsetFFTEM2 = nothing
			
			set fname=nothing
			set fs=nothing
			
			
			
			
			filename	= "fftem3" + sordernumber + ".csv"
			Set fs=Server.CreateObject("Scripting.FileSystemObject")
			Set fname=fs.CreateTextFile("C:\apps\zorrilla\webords\rdata\" & filename  ,true)
			
			Set RecordsetFFTEM3 = Server.CreateObject("ADODB.Recordset")	
			strSQL = "SELECT invno,ord_number,customer_no,order_stat,commchar FROM dbo.FFTEM3 WHERE invno = '" + sinvno + "' "	
			RecordsetFFTEM3.Open strSQL, MM_overseaspr_STRING
			Do While Not RecordsetFFTEM3.EOF
			
				fname.WriteLine( trim(RecordsetFFTEM3("invno")) & "," & trim(RecordsetFFTEM3("ord_number")) & "," & trim(RecordsetFFTEM3("customer_no")) & "," & trim(RecordsetFFTEM3("order_stat")) & "," & trim(RecordsetFFTEM3("commchar"))  )
				
				wrote3 = true
				
				RecordsetFFTEM3.MoveNext
			loop
			RecordsetFFTEM3.close
			SET RecordsetFFTEM3 = Nothing
			
			set fname=nothing
			set fs=nothing
			
			
			if wrote1 AND wrote2 AND wrote3 then
			
				Set oRS5 = Server.CreateObject("ADODB.Connection")
				oRS5.Open MM_overseaspr_STRING
				sql="UPDATE dbo.clients_orders SET order_status = 'P' WHERE order_number = '" + trim(sorder_number) + "' "
				on error resume next
				oRS5.Execute sql
				sError = err.description
				If len(sError) > 0 Then
					Response.Write("<h5 style='color:red'>DELETE FFTEM3 : " & sordernumber & "&nbsp;&nbsp;" & sError & "</h6>" )
					
				End if
				oRS5.close
				SET oRS5 = Nothing
				
			end if
			
			
	
	
	
			
			
			
	
		else
		
			Response.Write("<h5 style='color:red'>EOF USERS : " & trim(suserid) & "</h6>" )
						
        end if
		oRS4.close
		SET oRS4 = Nothing
		
	else
		
		Response.Write("<h5 style='color:red'>EOF clients_orders : " & trim(order_number) & "</h6>" )
		
	end if	
	oRS2.close
	SET oRS2 = Nothing
		
	oRS.MoveNext
Loop

oRS.close
SET oRS = Nothing

%>