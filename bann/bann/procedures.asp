<%
SUB CheckCredentials()
	' ONLY AFTER SUCESSFUL LOGIN, TO CREATE SESSION CREDENTIALS
	If Len(Session("MM_Username")) > 0 Then
		Set oRS = Server.CreateObject("ADODB.Recordset")
		' IF USER ACCESS IS VALIDATED IN TABLE, PLEASE ADD COLUMN IN WHERE CONDITION
		strSQL = "SELECT users.user_name, users.user_level, users.user_pwd, users.lang FROM users WHERE users.user_auto_id = '" + Session("MM_Username") + "' ;"
'		strSQL = "SELECT users.user_name, users.user_level, users.user_pwd FROM users WHERE users.user_name = '" + Session("MM_Username") + "' ;"		
		oRS.Open strSQL, MM_overseaspr_STRING
		
		If Not oRS.EOF Then
			' ACCESS GRANTED
			
			If oRS.Fields.Item("user_pwd") <> "11111" Then
			
				If oRS.Fields.Item("lang") = "" Then
					Session("lang") = "E"
					' tenía "S"
				Else
					Session("lang") = oRS.Fields.Item("lang")
				End if
				
				Session("user_level") = oRS.Fields.Item("user_level")
				Response.Redirect("part_search.asp")
			Else
				Session("MM_UserAuthorization") = ""
				Session("user_level") = ""
				Session("MM_Username2") = Session("MM_Username")
				Session("MM_Username") = ""
				
				Session("Last_Order_Number") = 0

				Response.Redirect("change_password.asp")
			End if

		Else
			' ACCESS NOT GRANTED
			ResetUserSession()
		End if
		
	End if
	
 END SUB

FUNCTION Lang(header)

		if LEN(Session("lang")) > 0 then
			language = Session("lang")
		else
			language = "E"
			Session("lang") = "E"
		end if		

		Select Case header
			'  See Order just processed				
			Case "VerOrden"
								
				If language = "E" Then Lang = "See order # " + Session("Last_Order_Number") Else Lang = "Ver orden # " + Session("Last_Order_Number") End If
			'Toggle specials items only in serach list

			Case "Liquidations"
				If language = "E" Then Lang = "Liquidation" Else Lang = "Liquidación" End If

			Case "Specials"
				If language = "E" Then Lang = "Disc." Else Lang = "Especial" End If
			'INDEX
			Case "password"
				If language = "E" Then Lang = "Verificar Credenciales de Acceso!" Else Lang = "Invalid user or password!" End If
			' TEMPLATE HEADERS
			Case "fecha"
				If language = "E" Then Lang = "Date" Else Lang = "Fecha" End If
			Case "completar_orden"
				If language = "E" Then Lang ="Finish Order" Else Lang = "Terminar Orden" End if
			Case "ACCOUNTSTATEMENT"
				If language = "E" Then Lang ="Account Statement" Else Lang = "Estado de Cuenta" End if	
			Case "salir"
				If language = "E" Then Lang ="Log Out" Else Lang = "Salir del Sistema" End if
			Case "preparado"
				If language = "E" Then Lang = "CART" Else Lang = "PREPARADO" End if
			Case "entre_10_piezas"
				If language = "E" Then Lang = "Enter up to 10 items" Else Lang = "Entre hasta 10 piezas" End if
			Case "family_keyword"
				If language = "E" Then Lang = "FAMILY/KEYWORD" Else Lang = "FAMILIA/FRASE" End if
			Case "family_category"
				If language = "E" Then Lang = "CATEGORY" Else Lang = "CATEGORIA" End if
			Case "modelo"
				If language = "E" Then Lang = "MODEL" Else Lang = "MODELO" End if
			Case "buscar"
				If language = "E" Then Lang = "Search" Else Lang = "Buscar" End if
			Case "ordenes"
				If language = "E" Then Lang = "ORDERS" Else Lang = "ORDENES" End if
			Case "estado_de_cuenta"
				If language = "E" Then Lang = "Account Payment" Else Lang = "Estado de Cuenta" End if
			Case "estado_de_cuenta_actual"
				If language = "E" Then Lang = "Account Statement" Else Lang = "Estado de Cuenta" End if
			Case "catalogo"
				If language = "E" Then Lang = "Engine Mount" Else Lang = "Catálogo de Soportes" End if
			
			' PART SEARCH RESULT
			Case "num_pieza"
				If language = "E" Then Lang = "PART NUMBER" Else Lang = "NUM. PIEZA" End if
			Case "reemplazo"
				If language = "E" Then Lang = "REPLACEMENT" Else Lang = "REEMPLAZO" End if
			Case "descripcion"
				If language = "E" Then Lang = "DESCRIPTION" Else Lang = "DESCRIPCION" End if
			Case "precio"
				If language = "E" Then Lang = "PRICE" Else Lang = "PRECIO" End if
			Case "disponible"
				If language = "E" Then Lang = "AVL." Else Lang = "DISP." End if
			Case "cant_ordenar"
				If language = "E" Then Lang = "AMT. TO<br>ORDER" Else Lang = "ORDENADO" End if
			Case "categoria"
				If language = "E" Then Lang = "CATEGORY" Else Lang = "CATEGORIA" End if

			'PAGE NAVIGATOR
			Case "primera_pagina"
				If language = "E" Then Lang = "First Page" Else Lang = "Primera Pagina" End if
			Case "pagina_anterior"
				If language = "E" Then Lang = "Previous Page" Else Lang = "Pagina Anterior" End if
			Case "proxima_pagina"
				If language = "E" Then Lang = "Next Page" Else Lang = "Proxima Pagina" End if
			Case "ultima_pagina"
				If language = "E" Then Lang = "Last Page" Else Lang = "Ultima Pagina" End if

			'NAVIGATOR STATUS
			Case "desplegando"
				If language = "E" Then Lang = "Record" Else Lang = "Record" End if
			Case "de"
				If language = "E" Then Lang = "of" Else Lang = "de" End if

			Case "print"
				If language = "E" Then Lang = "Print" Else Lang = "Imprimir" End if

			Case "total_piezas"
				If language = "E" Then Lang = "Total Parts" Else Lang = "Total de Piezas" End if
			Case "pagina"
				If language = "E" Then Lang = "Page" Else Lang = "Pagina" End if
			
			'CART
			Case "listado_piezas_seleccionadas"
				If language = "E" Then Lang = "Selected Parts List" Else Lang = "Piezas Seleccionadas" End if
			Case "pieza"
				If language = "E" Then Lang = "PART" Else Lang = "PIEZA" End if
			Case "ordenado"
				If language = "E" Then Lang = "ORDERED" Else Lang = "ORDENADO" End if
			Case "no_existen_piezas"
				If language = "E" Then Lang = "NO PARTS" Else Lang = "NO EXISTEN PIEZAS" End if
			Case "num_order"
				If language = "E" Then Lang = "ORDER NUM." Else Lang = "NUM. ORDEN" End if
			Case "orden_enviada"
			
			'ACCOUNT STATEMENT DISPLAY
				If language = "E" Then Lang = "Your Order Have Been Submited" Else Lang = "Su Orden a Sido Enviada" End if
			Case "order_through_our_web"
				If language = "E" Then Lang = "Order Procesed Through Our Website" Else Lang = "Orden procesada por internet" End if
			Case "no_existe_record"
				If language = "E" Then Lang = "No Record Exist" Else Lang = "No Existe Record" End if
			Case "num_factura"
				If language = "E" Then Lang = "INVOICE NUM." Else Lang = "NUM. FACTURA" End if
			Case "cant_fact_fecha"
				If language = "E" Then Lang = "INVOICE AMT./DATE" Else Lang = "NUM. FACTURA/FECHA" End if
			Case "pagos_recibos_fecha"
				If language = "E" Then Lang = "PAYMENT RECEIPT/DATE" Else Lang = "RECIBO PAGO/FECHA" End if
			Case "descuentos"
				If language = "E" Then Lang = "DISCOUNTS" Else Lang = "DESCUENTOS" End if
			Case "balance"
				If language = "E" Then Lang = "BALANCE" Else Lang = "BALANCE" End if
			Case "sub_total"
				If language = "E" Then Lang = "SUB TOTAL" Else Lang = "SUB TOTAL" End if
			Case "statement_days"
				If language = "E" Then Lang = "DAYS" Else Lang = "DIAS" End if
			Case "termino_pago"
				If language = "E" Then Lang = "Our payment terms is NET in 30 days." Else Lang = "Nuestro termino de pago es neto 30 dias." End if
			
			'CART SUBMIT iFrame
			Case "verificar_cart"
				If language = "E" Then Lang = "Please check part number and quantity before submitting the order." Else Lang = "Antes de someter la orden verifíque las piezas y sus cantidades." End if
			Case "orden_procesada"
				If language = "E" Then Lang = "The order will be procesed the next labor day if submited after 4:00 pm !" Else Lang = "Ordenes recibidas después de las 4:00 pm serán procesadas el próximo día laborable!" End if
			Case "Instrucciones"
				If language = "E" Then Lang = "Instructions:" Else Lang = "Instrucciones:" End if	
			Case "metodo_envio"
				If language = "E" Then Lang = "Delivery Method" Else Lang = "Método de Entrega" End if
			Case "entrega"
				If language = "E" Then Lang = "Delivery" Else Lang = "Entrega" End if
			Case "recoger"
				If language = "E" Then Lang = "Pickup" Else Lang = "Recoger" End if
			Case "someter_orden"
				If language = "E" Then Lang = "Submit Order" Else Lang = "Enviar Orden" End if
			Case "enviada_correctamente"
				If language = "E" Then Lang = "Order Sucessfully Submited!" Else Lang = "Orden Enviada Exitosamente" End if
			

			'STATUS
			Case "status_pending"
				If language = "E" Then Lang = "Pending" Else Lang = "Pendiente" End if
			Case "status_processed"
				If language = "E" Then Lang = "Processed" Else Lang = "Procesado" End if


			'ORDER DETAIL
			Case "detalles_orden"
				If language = "E" Then Lang = "Order Number: " Else Lang = "Orden Número:" End if
			Case "fecha_y_hora"
				If language = "E" Then Lang = "Date and Hour" Else Lang = "Fecha y Hora" End if
			Case "estatus"
				If language = "E" Then Lang = "Status" Else Lang = "Estatus" End if
			Case "detalles_piezas"
				If language = "E" Then Lang = "Parts Detail" Else Lang = "Detalle de piezas" End if
			
			'CHANGE PASSWORD
			Case "cambiar_contrasena"
				If language = "E" Then Lang = "CHANGE PASSWORD" Else Lang = "CAMBIAR CONTRASENA" End if
			Case "indique_nueva_contrasena"
				If language = "E" Then Lang = "Enter your new password." Else Lang = "Indique la nueva contraseña." End if
			Case "contrasena"
				If language = "E" Then Lang = "PASSWORD" Else Lang = "CONTRASENA" End if
			Case "confirmacion"
				If language = "E" Then Lang = "CONFIRMATION" Else Lang = "CONFIRMACION" End if
			
		End Select

END FUNCTION

SUB ResetUserSession()
	'RESET SESSION VARIABLES
	Session("MM_Username") = ""
	Session("MM_UserAuthorization") = ""
	Session("user_level") = ""
	Response.Redirect("index.asp")
END SUB
 
FUNCTION FamilyCode(fam_desc)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	if Session("lang") = "S" then
		strSQL = "SELECT dbo.familicat.field_2 FROM dbo.familicat WHERE dbo.familicat.field_1 = '" + fam_desc + "' ;"
	else
		strSQL = "SELECT dbo.familicat.field_2 FROM dbo.familicat WHERE dbo.familicat.english_desc = '" + fam_desc + "' ;"
	end if	
	oRS.Open strSQL, MM_overseaspr_STRING
	If Not oRS.EOF Then
		FamilyCode = CStr(oRS.Fields.Item("field_2"))
	End if
END FUNCTION
 
Sub DisplayParts(ItemsPerPage)	
	Dim RCount
	Dim nItemsPerPage
	Dim nPageCount
	Dim sqlScript
	
	sqlScript = Session("SQLSearch")
	nItemsPerPage = ItemsPerPage
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.CursorLocation = 3 'adUseClient
	if len(sqlScript) > 0 then
		'Replace sqlScript, "'", "''"
		RS.Open sqlScript, MM_overseaspr_STRING
		'Response.Write sqlScript
		if Not RS.EOF Then
			RCount = CStr(RecordsetCount(RS))
			RS.PageSize = nItemsPerPage
			nPageCount = RS.PageCount
			nPage = CLng(Request.QueryString("Page"))
			If nPage < 1 Or nPage > nPageCount Then
				nPage = 1
			End If
			Response.Write "<br><table cellpadding=0 cellspacing=0 width='96%'><tr><td align='left' >"
			NavigationStatus RS, nPage, nPageCount,RCount,nItemsPerPage
			Response.Write "<br></td></tr><tr><td align=""left"">"
			NavigationControls nPage,nPageCount,"part_search.asp"
			Response.Write "<br></td></tr></table>"
			RS.AbsolutePage = nPage
			' nPageCount, RCount, nItemsPerPage
			DisplayResult RS,nPage,nPageCount, RCount, nItemsPerPage
			NavigationControls nPage,nPageCount,"part_search.asp"
			Response.Write "<br><br>"
		end if
		RS.Close
		
	end if
	
	Set RS = Nothing
End Sub

Sub DisplayResult(RecordsetObject,nPage,nPageCount, RCount, nItemsPerPage)
	'-------- AVAILABLE DATA IN RECORDSET

	' PARTS COUNTER, VALIDATE TO BE USED ON N PAGES.
	If nPage = 1 Then counter = 1 Else counter = (nItemsPerpage * nPage) - nItemsPerpage + 1
	
	Response.Write "<table cellpadding=""0"" cellspacing=""0""  border=""0"" width=""700"" class=""font10"" >"
	Response.Write "<tr><td colspan=""9"" style=""background-color:#CCCCCC;""><hr style=""color:black;vertical-align:top"" size=""1""/></td></tr>"
	Response.Write "<tr style=""background-color:#CCCCCC;"">"
	Response.Write "<td width=""88px"" align=""left"">&nbsp;&nbsp;<font color=black></font></td>"
	Response.Write "<td width=""88px"" align=""left""><font color=black><strong>&nbsp;" + Lang("num_pieza") + "</strong></font></td>"
	Response.Write "<td width=""88px"" align=""left""><font color=black><strong>&nbsp;" + Lang("reemplazo") + "</strong></font></td>"
	Response.Write "<td width=""88px"" align=""left""><font color=black><strong>&nbsp;" + Lang("descripcion") + "</strong></font></td>"
	Response.Write "<td width=""88px"" align=""left""><font color=black><strong>" + Lang("precio") + "</strong></font></td>"
	Response.Write "<td width=""88px"" align=""left""><font color=black><strong>&nbsp;" + Lang("disponible") + "</strong></font></td>"
	Response.Write "<td width=""88px"" align=""left""><font color=black><strong>" + Lang("cant_ordenar") + "</strong></font></td>"
	Response.Write "<td width=""88px"" align=""left""><font color=black><strong>" + Lang("categoria") + "</strong></font></td>"
	Response.Write "<td width=""88px"" align=""left""><font color=black><strong></strong></font></td>"	
	
	Response.Write "<tr><td colspan=""9"" style=""background-color:#CCCCCC;""><hr style=""color:black;vertical-align:bottom"" size=""1""/></td></tr>"
	Response.Write "</tr>" '</table>"
	'Response.Write "<table cellpadding=0 cellspacing=0  border=0 width=""700"">"
	Do While Not (RecordsetObject.EOF or RecordsetObject.AbsolutePage <> nPage)
		
		Response.Write "<tr>"

		Response.Write "<td width='25px' align='center' >" + CStr(counter) + "</td>"
		
		Response.Write "<td width='75px' align=""left"">" + PartPhoto(Trim(RecordsetObject("field_1").Value)) + "</td>" ' NUM. PIEZA
		
		Response.Write "<td width='75px' align=""left"">"
		
		If len(Trim(RecordsetObject("replacement").Value)) > 0 Then
			Response.Write Trim(RecordsetObject("replacement").Value) ' REEMPLAZO
		Else
			Response.Write "<center>N/A</center>"
		End if
		
		Response.Write "</td>"
		Response.Write "<td width='400px' align=""left"">" 
		
		If Session("lang") = "E" Then
			If len(trim(RecordsetObject("english_version").Value)) > 0 Then
				Response.Write RecordsetObject("english_version").Value + "</td>" 'DESCRIPCION INGLES
			Else
				Response.Write RecordsetObject("field_2").Value + "</td>" 'DESCRIPCION ESPANOL
			End if	
		Else	
			Response.Write RecordsetObject("field_2").Value + "</td>" 'DESCRIPCION ESPANOL
		End if
		
		Response.Write "<td width='40px' align=""left"">" + FinalPrice(RecordsetObject("field_1").Value) + "</td>" 'PRECIO para otro usuario
		
		if RecordsetObject("field_5").Value > 50 then
			Response.Write "<td width='25px' align=""left"">50+</td>" 'DISPONIBILIDAD
		else	
			Response.Write "<td width='25px' align=""left"">" + CStr(RecordsetObject("field_5").Value) + "</td>" 'DISPONIBILIDAD
		end if	
		Response.Write "<td width='25px' align=""center""><input id='cant_ordenar_" + CStr(counter) + "' class='search_box' size='4' value='1' onFocus='this.select()'/>  </td>" 'CANT. A ORDENAR
		
		if Session("lang") = "S" then
			Response.Write "<td width='75px' align=""center"">" + RecordsetObject("familia_descripcion").Value +  "</td>" 'FAMILIA / CATEGORIA
		else
			Response.Write "<td width='75px' align=""center"">" + RecordsetObject("family_description").Value +  "</td>" 'FAMILIA / CATEGORIA
		end if	
		Response.Write "<td width='25px' align=""right"">&nbsp;&nbsp;&nbsp;"
		
		if PartAvailability(RecordsetObject("field_1").Value) > 0 Then
			if CheckPartInCart(RecordsetObject("field_1").Value) then
				Response.Write "<div style='margin-left:10px' width='36px' id='add_cart_" + CStr(counter) + "'><a style='width:100%' href='#' onclick='ValidateCart(""" + RecordsetObject("field_1").Value + """,""cant_ordenar_" + CStr(counter) + """);'><img src='images/add-to-cart.gif' width='26' height='18' border='0' onclick=""document.getElementById('add_cart_" + CStr(counter) + "').style.visibility='hidden'"" /></a></div>"
			end if	
		end if
		
		Response.Write "</tr>"
		Response.Write "<tr>"
        Response.Write "<td colspan=""9""><hr style=""color:green"" size=""1""/></td>"
        Response.Write "</tr>"
		
		RecordsetObject.MoveNext
		counter = counter + 1
	Loop
	Response.Write "</table>"
End Sub

Sub NavigationControls(nPage,nPageCount,nPageName)
	Response.Write "<table cellpadding=0 cellspacing=0 width=""700px""><tr>"
	Response.Write "<td align=""center""><a href=" + nPageName + "?Page=1 >&nbsp;" + Lang("primera_pagina") + "&nbsp;</a>&nbsp;&nbsp;</td>"
	Response.Write "<td align=""center""><a href=" + nPageName + "?Page=" + CStr(nPage - 1) + " >&nbsp;" + Lang("pagina_anterior") + "&nbsp;</a>&nbsp;&nbsp;</td>"
	Response.Write "<td align=""center""><a href=" + nPageName + "?Page=" + CStr(nPage + 1) + " >&nbsp;" + Lang("proxima_pagina") + "&nbsp;</a>&nbsp;&nbsp;</td>"
	Response.Write "<td align=""center""><a href=" + nPageName + "?Page=" + CStr(nPageCount) + " >&nbsp;" + Lang("ultima_pagina") + "&nbsp;</a></td>"
	Response.Write "<td align=""center"">&nbsp;</td>"
	Response.Write "</tr></table>"

End Sub

Sub NavigationStatus(RecordSetObject, nPage, nPageCount,RCount,nItemsPerPage)
	Response.Write "<table cellpadding=0 cellspacing=0 width=""100%"" border=0 class=""search_titles""><tr><td>"
	Response.Write Lang("desplegando") + ": <strong>" + CStr(RecordsPerPageStatus (RecordsetObject, nPage,nPageCount,nItemsPerPage,RCount)) + "</strong> " + Lang("de") + " <strong>" + CStr(RCount) + "</strong></td><td>"
	Response.Write Lang("total_piezas") + ": <strong style=""color:black"">" + CStr(RCount) + "</strong></td><td align=""right"">"
	Response.Write Lang("pagina") + ": <strong>" + CStr(nPage) + "</strong> " + Lang("de") + " <strong>" + CStr(nPageCount) + "</strong></td></tr></table>"
End Sub

Function RecordsetCount(RecordsetObject)
	Do While Not RecordsetObject.EOF
		RCount = RCount + 1
		RecordsetObject.MoveNext
	Loop
	RecordsetObject.MoveFirst
	RecordsetCount = RCount
End Function

Function RecordsPerPageStatus(RecordsetObject,nPage,nPageCount,nItemsPerPage,RCount)
	Do While Not (RecordsetObject.EOF)
		total_rec = total_rec + 1
		RecordsetObject.MoveNext
	Loop
	final_rec = (nPage * nItemsPerPage)
	if nPage = nPageCount then
		final_rec = RCount
	end if
	RecordsetObject.MoveFirst
	RecordsPerPageStatus = final_rec
End Function

SUB CartDisplay(ShowDescription)

if ShowDescription = "" then ShowDescription = "Y"
If PartsOnCart() Then
	'Response.Write "[" + CStr(ShowDescription) + "]"
	Set oRS = Server.CreateObject("ADODB.Recordset")
		
	strSQL = "SELECT dbo.clients_cart.shop_auto_id, dbo.clients_cart.shop_part_number, dbo.clients_cart.shop_part_price, dbo.clients_cart.shop_auto_id, dbo.clients_cart.shop_quantity FROM dbo.clients_cart WHERE dbo.clients_cart.shop_client_user = '" + Session("MM_Username") + "' ;"
	
	oRS.Open strSQL, MM_overseaspr_STRING
	
	Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='tablas'>"
	Response.Write "<tr height='16' >"
	Response.Write "<td  align='center' bgcolor='#F2F2F2'><strong>" + Lang("pieza") + "</strong></td>"
	if ShowDescription = "Y" then
	 Response.Write "<td width='350'  align='center' bgcolor='#F2F2F2'><strong>" + Lang("descripcion") + "</strong></td>"
	else
	 Response.Write "<td  align='center' bgcolor='#F2F2F2'><strong>&nbsp;</strong></td>"
	end if
	Response.Write "<td  align='center' bgcolor='#F2F2F2'><strong>&nbsp;" + Lang("ordenado") + "&nbsp;</strong></td>"
	Response.Write "<td  align='center' bgcolor='#F2F2F2'><strong>&nbsp;" + Lang("precio") + "&nbsp;</strong></td>"


	Response.Write "<td  align='center' bgcolor='#F2F2F2'>&nbsp;&nbsp;&nbsp;&nbsp;<strong>Subtotal</strong></td>"

	Response.Write "<td  width='16' bgcolor='#F2F2F2'>&nbsp;</td>"
		
	Response.Write "</tr>"
	
'	Response.Write "  <tr height='10px'>"
'	Response.Write "     <td align='center'>&nbsp;</td>"
'	Response.Write "  </tr>"

	' ALTERNATE ROW COLOR VARIABLES
	color1 = "#CCCCCC"
	color2 = "#FFFFFF"
	current_color = color1
	rc = 1

	'TOTAL EN EL CART
	cart_total = 0
	Do While Not oRS.EOF
	
		if ShowDescription = "Y" then
			' Record set to get description in english or spanish for part item ...
			Set oRS2 = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT dbo.partmst1_distinct.field_2, dbo.partmst1_distinct.english_version FROM dbo.partmst1_distinct WHERE dbo.partmst1_distinct.field_1 = '" + CStr(oRS.Fields.Item("shop_part_number"))  + "' ;"
			oRS2.Open strSQL, MM_overseaspr_STRING
			ldescription = ""
			If Session("lang") = "E" Then
				if Trim(oRS2.Fields.Item("english_version")) > "" then
					ldescription = Trim(oRS2.Fields.Item("english_version"))
				else
					ldescription = Trim(oRS2.Fields.Item("field_2"))
				end if		
			else
				ldescription = Trim(oRS2.Fields.Item("field_2"))
			End if
			'  mod ends here .
		end if
		
		Response.Write "  <tr>"
		Response.Write "    <td style='font-size:9px'  align='center' bgcolor='" +  current_color + "' >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + Trim(oRS.Fields.Item("shop_part_number")) + "</td>"
		Response.Write "    <td align='left' bgcolor='" +  current_color + "' height='10px'>&nbsp;&nbsp;&nbsp;" + ldescription + "</td>"		
		
		' allow update quantity order in shopping cart
		if ShowDescription = "Y" then

			Response.Write "    <td align='center' bgcolor='" +  current_color + "'><input style='font-size:12px' valign='top' maxlength='4' size='2' name='shop_quantity" & oRS.Fields("shop_auto_id").Value & "' type='text' id='shop_quantity" & oRS.Fields("shop_auto_id").Value & "' value='" + CStr(oRS.Fields.Item("shop_quantity")) + "'>"

		response.write "&nbsp;<a id='UpdCart' href='#' onclick='UpdCart(""" + CStr(oRS.Fields.Item("shop_auto_id")) + """);'><img valign='baseline' src='images/20x20-save.png' width='16' height='16' border='0' style='margin:0px;padding:0px' />&nbsp;</a></td>"

		else

			Response.Write "    <td align='center' bgcolor='" +  current_color + "'>" + CStr(oRS.Fields.Item("shop_quantity")) + "</td>"
			
		end if				

		
'UPDATE QUANTITY AJAX

		' Show that part has a special price or discounted ...
		if PartOnSale( CStr(oRS.Fields.Item("shop_part_number")) , oRS.Fields.Item("shop_part_price") ) then
	
 		 Response.Write "    <td align='center' color='RED' bgcolor='" +  current_color + "'><strong><font color='red'>" + CStr(CurrencyConvert(oRS.Fields.Item("shop_part_price"))) + "</font></td>"
		else
		 Response.Write "    <td align='center' bgcolor='" +  current_color + "'>" + CStr(CurrencyConvert(oRS.Fields.Item("shop_part_price"))) + "</td>"
		end if

		cart_total = (oRS.Fields.Item("shop_part_price") * oRS.Fields.Item("shop_quantity")) + cart_total
		cart_subtotal = (oRS.Fields.Item("shop_part_price") * oRS.Fields.Item("shop_quantity"))
		
		if ShowDescription = "Y" then

		 Response.Write "    <td width='30px' align='right' bgcolor='" +  current_color + "'>" + CStr(CurrencyConvert(cart_subtotal)) + "</td>"

		Response.Write "    <td  align='right' bgcolor='" + current_color + "'>&nbsp;</td>"

		end if

		Response.Write "    <td align='right' bgcolor='" + current_color + "'><a  id='DelFromCart' href='#' onclick='DelFromCart(""" + CStr(oRS.Fields.Item("shop_auto_id")) + """);'><img valign='baseline' src='images/del.gif' width='16' height='16' border='0' style='margin:0px;padding:0px' />&nbsp;</a></td>"


		
		Response.Write "  </tr>"
		oRS.MoveNext
		
		if ShowDescription = "Y" then
			'  close oRS2 ...
			oRS2.Close
			Set oRS2 = Nothing
		end if
				
		If current_color = color2 then current_color = color1 else current_color = color2
		rc = rc + 1
    
	Loop
	
	Response.Write "<tr><td colspan='5'><hr></td></tr>"
	Response.Write "<tr valign='top'><td align='right' colspan='4' ><strong>Total:&nbsp;" + CurrencyConvert(cart_total) + "</strong></td><td>&nbsp;</td></tr>"
	
'	If PartsOnCart() Then 
	If Request.ServerVariables("SCRIPT_NAME") = "part_search.asp"  or Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/part_search.asp" Then
		Response.Write " <tr><td colspan='5'><h3><center><a style='width:100%' href='cart.asp'>" + Lang("completar_orden") + "</a></center></h3></td></tr>"
	End if
'	End if
	
	Response.Write "</table>"
	
         


Else
	Response.Write "<h5>" + Lang("no_existen_piezas") + "</h5> "
End if
END SUB

// 
FUNCTION PartOnSale(pieza,precio)
	
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.prespecials.especial FROM dbo.prespecials WHERE dbo.prespecials.specials = '" + pieza + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
	
	PartOnSale = false		
	If Not oRS.EOF Then
		if CStr(oRS.Fields.Item("especial")) = CStr(precio) then
			PartOnSale = true
		else
			PartOnSale = false
		end if	
	End if
	oRS.Close
	Set oRS = Nothing

END FUNCTION


SUB AddToCart(pieza,cant)

	client_code = UserClientCode(Session("MM_Username"))
	user_id		= Session("MM_Username")
	available	= PartAvailability(pieza)
	cant		= CStr(cant)
	precio		= CurrencyConvert(PricePerUser(pieza))
	
	Set objConn5 = Server.CreateObject("ADODB.Connection")
		objConn5.Open MM_overseaspr_STRING
	
sql="INSERT INTO clients_cart (shop_client_code,shop_client_user,shop_product_id,shop_quantity,shop_part_price,shop_part_number,available,shop_order_date,shop_type) "
sql=sql & " VALUES "
sql=sql & "('" & client_code & "',"
sql=sql & "'" & user_id & "',"
sql=sql & "'" & pieza & "',"
sql=sql & "'" & cant & "',"
sql=sql & "'" & precio & "',"
sql=sql & "'" & pieza & "',"
sql=sql & "'" & available & "',"
sql=sql & "'" & CStr(Now()) & "',"
sql=sql & "'W' )  ;" 	  
						   
	objConn5.Execute sql
	sError = err.description
	objConn5.Close 
	
	If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error añadiendo esta pieza al cart: <br><br><strong>" & sError & "</strong")
	End if

END SUB

SUB DelFromCart(auto_id)
	Set objConn5 = Server.CreateObject("ADODB.Connection")
	objConn5.Open MM_overseaspr_STRING
	sql = "DELETE FROM clients_cart WHERE shop_auto_id = " + CStr(auto_id) + " AND shop_client_user = '" + Session("MM_Username") + "' ;"
	objConn5.Execute sql
	'response.Write sql
	sError = err.description
	
	 If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error eliminando esta pieza al cart: <br><br><strong>" & sError & "</strong")
	End if
	
END SUB

SUB UpdateCart( auto_id, qty )

	'response.write("LLEGO")
	
	Set objConn5 = Server.CreateObject("ADODB.Connection")
	objConn5.Open MM_overseaspr_STRING
	sql = "UPDATE clients_cart SET shop_quantity = " & qty & "WHERE shop_auto_id = " + CStr(auto_id) + " AND shop_client_user = '" + Session("MM_Username") + "' ;"
	objConn5.Execute sql
	'response.Write sql
	sError = err.description
	
	If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error eliminando esta pieza al cart: <br><br><strong>" & sError & "</strong")
	End if
	
END SUB

FUNCTION UserClientCode(user_name)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.users.user_client_code FROM dbo.users WHERE dbo.users.user_auto_id = '" + Session("MM_Username") + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
	If Not oRS.EOF Then UserClientCode = CStr(oRS.Fields.Item("user_client_code"))
END FUNCTION

FUNCTION PartAvailability(part)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.partmst1_distinct.field_5 FROM dbo.partmst1_distinct WHERE dbo.partmst1_distinct.field_1 = '" + part + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
	If Not oRS.EOF Then PartAvailability = CStr(oRS.Fields.Item("field_5"))
END FUNCTION

FUNCTION PricePerUser(pieza)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.prespecials.especial FROM dbo.prespecials WHERE dbo.prespecials.specials = '" + pieza + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
		
	If Not oRS.EOF Then
		PricePerUser = CStr(oRS.Fields.Item("especial"))
	Else
		oRS.Close
		' LUIS ESTO LO CAMBIE  user_level '3' = precio mas barato que es field_4  !!!!!   GGG
		If Session("user_level") = "3" Then
			strSQL = "SELECT field_4 FROM partmst1_distinct WHERE field_1 = '" + pieza + "' ;"
			oRS.Open strSQL
			PricePerUser = CStr(oRS.Fields.Item("field_4"))
		Else
			strSQL = "SELECT field_3 FROM partmst1_distinct WHERE field_1 = '" + pieza + "' ;"
			oRS.Open strSQL
			PricePerUser = CStr(oRS.Fields.Item("field_3"))
		End if
		oRS.Close
		Set oRS = Nothing
	End if
END FUNCTION

FUNCTION CurrencyConvert(precio)
	CurrencyConvert = FormatCurrency(precio, 2)
END FUNCTION

FUNCTION CheckPartInCart(part)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT Count(dbo.clients_cart.shop_part_number) as countrecs FROM dbo.clients_cart WHERE dbo.clients_cart.shop_part_number = '" + part + "' AND dbo.clients_cart.shop_client_user = '" + Session("MM_Username") + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
	if Not oRS.EOF Then
		if oRS.Fields.Item("countrecs") > 0 then
			CheckPartInCart = false
		Else
			CheckPartInCart = true
		end if	
	else
		CheckPartInCart = true
	end if
END FUNCTION

FUNCTION MultipleModelKeywords(search_field,keyword)
	a = Split(keyword)
	like_sql = ""
	counter = 0
	MyKeyword = ""
	
	
	for each x in a
		MyKeyword = ModelKeywordShortcut(x)
		
		If MyKeyword = x Then
		
			if counter = 0 then
				like_sql = like_sql + search_field + " LIKE '%" + MyKeyword + "%' "
			Else
				like_sql = like_sql + " AND (" + search_field + " LIKE '%" + MyKeyword + "%' ) "
			End if
		Else
		
			if counter = 0 then
				like_sql = like_sql + search_field + " LIKE '%" + MyKeyword + "%' or " + search_field + " LIKE '%" + x + "%' "
			Else
				like_sql = like_sql + " AND ( " + search_field + " LIKE '%" + MyKeyword + "%' ) or (" + search_field + " LIKE '%" + x + "%' ) "
			End if
			
		End if
		counter = counter + 1
	next
	
	MultipleModelKeywords = like_sql
	
END FUNCTION

FUNCTION ModelKeywordShortcut(keyword)

	' VALIDATE IF USER ADD ' (coma) IN SEARCH CRITERIA TO CALL AddComa
	If InStr(keyword,"'") > 0 Then keyword = AddComa(keyword)


	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.mykeywords.replacedby FROM dbo.mykeywords WHERE Upper(dbo.mykeywords.keyword) = '" + Ucase(keyword) + "' ;"    
	oRS.Open strSQL, MM_overseaspr_STRING
	
	If Not oRS.EOF Then
		ModelKeywordShortcut = AddComa(Trim(oRS.Fields.Item("replacedby")))
	Else
		ModelKeywordShortcut = keyword
	End if

END FUNCTION

FUNCTION AddComa(keyword)
		pos1 = InStr(keyword,"'")	
		If pos1 > 0 Then
			AddComa = Mid(keyword,1,pos1) + "'" + Mid(keyword,pos1 + 1,len(keyword))
		Else
			AddComa = keyword
		End if
END FUNCTION

SUB OrdersDisplay()

	Set oRS = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "SELECT DISTINCT (dbo.clients_orders.order_number) FROM dbo.clients_orders WHERE dbo.clients_orders.order_user = '" + Session("MM_Username") + "' order by dbo.clients_orders.order_number desc ;"
	
	oRS.Open strSQL, MM_overseaspr_STRING
	
	Response.Write "<table margin='0' border='0' cellpadding='0' cellspacing='0' >"
	Response.Write "  <tr>"
	Response.Write "    <td width='25%' height='20px' align='center' bgcolor='#F2F2F2'><strong>" + Lang("num_order") + "</strong></td>"
	Response.Write "    <td width='50%' align='center' bgcolor='#F2F2F2'><strong>" + Lang("fecha") + "</strong></td>"
	Response.Write "    <td width='25%' align='center' bgcolor='#F2F2F2'><strong>TOTAL</strong></td>"
	Response.Write "  </tr>"
'	Response.Write "  <tr>"
'	Response.Write "    <td align='center'>&nbsp;</td>"
'	Response.Write "    <td align='center'>&nbsp;</td>"
'	Response.Write "    <td align='center'>&nbsp;</td>"
'	Response.Write "  </tr>"

	' ALTERNATE ROW COLOR VARIABLES
	color1 = "#F2F2F2"
	color2 = "#FFFFFF"
	current_color = color1
	rc = 1
	
	Do While Not oRS.EOF
	
		Response.Write "  <tr height='20px'>"
		Response.Write "    <td align='right' height='20px' bgcolor='" +  current_color + "'><a style='width:100%' href='#' OnClick='location.href=""order_detail.asp?o=" + CStr(oRS.Fields.Item("order_number")) + """;'>" + CStr(oRS.Fields.Item("order_number")) + "</a></td>"
		Response.Write "    <td align='center' bgcolor='" +  current_color + "'>&nbsp;" + FechaOrden(CStr(oRS.Fields.Item("order_number"))) + "</td>"
		Response.Write "    <td align='right' bgcolor='" +  current_color + "'>" + OrderTotal(CStr(oRS.Fields.Item("order_number"))) + "&nbsp;&nbsp;&nbsp;</td>"
		Response.Write "  </tr>"
		
		oRS.MoveNext
		
		If current_color = color2 then current_color = color1 else current_color = color2
		rc = rc + 1
    
	Loop
	
'	Response.Write "  <tr>"
'	Response.Write "    <td align='center'>&nbsp;</td>"
'	Response.Write "    <td align='center'>&nbsp;</td>"
'	Response.Write "    <td align='center'>&nbsp;</td>"
'	Response.Write "  </tr>"
	
	Response.Write "</table>"
	
END SUB

FUNCTION FechaOrden(num_order)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "SELECT dbo.clients_orders.order_date FROM dbo.clients_orders WHERE dbo.clients_orders.order_number = '" + num_order + "' ;"
	
	oRS.Open strSQL, MM_overseaspr_STRING
	
	FechaOrden = CStr(FormatDateTime(oRS.Fields.Item("order_date"),2))
	
END FUNCTION

FUNCTION OrderTotal(num_order)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "SELECT dbo.clients_orders.item_price * dbo.clients_orders.order_qty as item_price FROM dbo.clients_orders WHERE dbo.clients_orders.order_number = '" + num_order + "' ;"
	
	oRS.Open strSQL, MM_overseaspr_STRING
	
	order_total = 0
	
	If Not oRS.EOF Then
		Do While Not oRS.EOF
			order_total = order_total + oRS.Fields.Item("item_price")
			oRS.MoveNext
		Loop
	End if
	
	OrderTotal = CurrencyConvert(order_total)
	
END FUNCTION

FUNCTION PartsOnCart()
	Set oRS = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "SELECT dbo.clients_cart.shop_auto_id  FROM dbo.clients_cart WHERE dbo.clients_cart.shop_client_user = '" + Session("MM_username") + "' ;"
	
	oRS.Open strSQL, MM_overseaspr_STRING
	
	If Not oRS.EOF Then 
		PartsOnCart = True
	Else
		PartsOnCart = False
	End if
	
END FUNCTION

SUB SleepConnection()
	Set conn = CreateObject("ADODB.Connection")
	conn.Open MM_overseaspr_STRING
	'conn.commandTimeout = 10 + 5
	sql = "WAITFOR DELAY '00:00:03:00'"
	'Response.Write now & "<p>" & sql
	conn.Execute sql,,129
	'Response.Write now & "<p>" & sql
	conn.close
	Set conn = Nothing
END SUB

FUNCTION AvailableOrderNumber()
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.next_order_number.in_use FROM dbo.next_order_number"
	oRS.Open strSQL, MM_overseaspr_STRING
	If Trim(oRS.Fields.Item("in_use")) = "N" Then
		AvailableOrderNumber = True
	Else
		SleepConnection()
		AvailableOrderNumber = False
	End if
END FUNCTION

FUNCTION NextOrderNumber()

	NewOrderNumber = 0
	CurrentOrderNumber = 0

	'UPDATE IN_USE FIELD TO 'Y'
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open MM_overseaspr_STRING
	sql="UPDATE dbo.next_order_number SET in_use = 'Y' ;"
	objConn.Execute sql
	objConn.Close
	
	'SELECT NEXT ORDER NUMBER
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "  SELECT dbo.next_order_number.order_number FROM dbo.next_order_number WHERE in_use = 'Y' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
	CurrentOrderNumber = oRS.Fields.Item("order_number")
	NewOrderNumber = CurrentOrderNumber + 1
	
	'UPDATE IN_USE_FIELD TO 'N'
	objConn.Open MM_overseaspr_STRING
	sql="UPDATE dbo.next_order_number SET order_number = " + CStr(NewOrderNumber) + ", in_use = 'N'  ;"
	objConn.Execute sql
	objConn.Close
	Set objConn = Nothing
	
	'RETURN NEXT ORDER NUMBER
	NextOrderNumber = NewOrderNumber
	
END FUNCTION

SUB SubstrackQOH(part,qty)
	'SELECT CURRENT QOH FOR PART
	parts_available = PartAvailability(part)
	
	'UPDATE QOH
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open MM_overseaspr_STRING
	sql="UPDATE dbo.partmst1_distinct SET field_5 = " + CStr(parts_available - qty) + "  WHERE dbo.partmst1_distinct.field_1 = '" + part + "' ;"
	objConn.Execute sql
	objConn.Close
	Set objConn = Nothing
	
END SUB

SUB ResetCart(user)
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open MM_overseaspr_STRING
	sql="DELETE FROM dbo.clients_cart WHERE dbo.clients_cart.shop_client_user = '" + user + "' ;"
	objConn.Execute sql
	objConn.Close
	Set objConn = Nothing
END SUB

SUB OrderSubmit(DeliveryType,instructions)

	Session("Last_Order_Number") = 0
	
	next_order_num = 0
	available_order_number = AvailableOrderNumber()
	
	If available_order_number Then
	
		AddOrder DeliveryType,instructions
		Response.Write "<br><br><hr><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + Lang("enviada_correctamente")
		
		if Session("Last_Order_Number") > 0 then
			' iframe use target = _parent
			Response.Write ( "<hr><br>&nbsp;<a style='width:100%'  href='/order_detail.asp?o=" + Session("Last_Order_Number") + "' target='_parent'>" + Lang("VerOrden") + "</a>" )

		End if
			
	Else
		counter = 0
		Do While NOT available_order_number = True
			
			If available_order_number = True Then
				
				' ADD ORDER
				'Response.Write "Number Available"
				AddOrder DeliveryType,instructions
				
				Response.Write Lang("enviada_correctamente")
		
			
			End if
			available_order_number = AvailableOrderNumber()
			counter = counter + 1
		Loop
	End if
		
END SUB

SUB AddOrder(DeliveryType,instructions)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.clients_cart.shop_auto_id, dbo.clients_cart.shop_client_code, dbo.clients_cart.shop_client_user, dbo.clients_cart.shop_product_id,  dbo.clients_cart.shop_quantity, dbo.clients_cart.shop_part_price, dbo.clients_cart.shop_part_number FROM dbo.clients_cart WHERE dbo.clients_cart.shop_product_id not in ( select dbo.prespecials.specials from dbo.prespecials ) and  dbo.clients_cart.shop_client_user = '" + Session("MM_Username") + "' ;"
	
	oRS.Open strSQL, MM_overseaspr_STRING
	
	client_code 		=	UserClientCode(Session("MM_Username"))
	client_user 		=	Session("MM_Username")
	//next_order_number	=	CStr(NextOrderNumber())
	loop_counter        =   0
	
		
	Do While Not oRS.EOF
	
		loop_counter = loop_counter + 1

		if loop_counter = 1 then
			next_order_number	=	CStr(NextOrderNumber())
			
			// New code for comments ...
			if trim(instructions) > "" then
				// Create ZZZNOF
					
				// Get next order number reset counter from 0 to 1
					
				part_id 			=	"ZZZNOF"
				quantity			=	0
				part_price			=	0
					
				'INSERT COMMENT ...
				Set objConn = Server.CreateObject("ADODB.Connection")
				objConn.Open MM_overseaspr_STRING
				sql="INSERT INTO dbo.clients_orders"
				 sql=sql & "( order_number,   "
				 sql=sql & "order_part,   "
				 sql=sql & "order_client,  " 
				 sql=sql & "order_user,   "
				 sql=sql & "order_qty,   "
				 sql=sql & "order_part_id,"   
				 sql=sql & "order_status,  " 
				 sql=sql & "order_date,   "
				 sql=sql & "order_type,   "
				 sql=sql & "item_price,   "
				 sql=sql & "comments,  "
				 sql=sql & "deliv_type )  "
				 sql=sql & "VALUES ( '" & next_order_number & "', "  
				 sql=sql & "'" & part_id & "',   "
				 sql=sql & "'" & client_code & "',   "
				 sql=sql & "'" & Session("MM_Username") & "',   "
				 sql=sql & CStr(quantity) & ", "  
				 sql=sql & "'" & part_id & "',   "
				 sql=sql & "'O',   "
				 sql=sql & "'" & CStr(Now()) & "',  "
				 sql=sql & "'" & DeliveryType & "',  "
				 sql=sql & CStr(part_price) & ",  "
				 sql=sql & "'" & instructions & "',  "
				 sql=sql & "'" & DeliveryType & "')  ;"
				
				objConn.Execute sql
				objConn.Close
				Set objConn = Nothing
				
			end if
			//

		End if	
		if loop_counter = 22 then
			next_order_number	=	CStr(NextOrderNumber())
			loop_counter        =   1
		End if	
				 
		part_availability	=	PartAvailability(oRS.Fields.Item("shop_part_number"))
		part_id 			=	oRS.Fields.Item("shop_product_id")
		quantity			=	CStr(oRS.Fields.Item("shop_quantity"))
		part_price			=	CStr(oRS.Fields.Item("shop_part_price"))
			
			'INSERT PART INTO CLIENTS ORDER
			Set objConn = Server.CreateObject("ADODB.Connection")
			objConn.Open MM_overseaspr_STRING
			sql="INSERT INTO dbo.clients_orders"
			 sql=sql & "( order_number,   "
			 sql=sql & "order_part,   "
			 sql=sql & "order_client,  " 
			 sql=sql & "order_user,   "
			 sql=sql & "order_qty,   "
			 sql=sql & "order_part_id,"   
			 sql=sql & "order_status,  " 
			 sql=sql & "order_date,   "
			 sql=sql & "order_type,   "
			 sql=sql & "item_price,   "
			 sql=sql & "comments,  "
			 sql=sql & "deliv_type )  "
			 sql=sql & "VALUES ( '" & next_order_number & "', "  
			 sql=sql & "'" & part_id & "',   "
			 sql=sql & "'" & client_code & "',   "
			 sql=sql & "'" & Session("MM_Username") & "',   "
			 sql=sql & CStr(quantity) & ", "  
			 sql=sql & "'" & part_id & "',   "
			 sql=sql & "'O',   "
			 sql=sql & "'" & CStr(Now()) & "',  "
			 sql=sql & "'" & DeliveryType & "',  "
			 sql=sql & CStr(part_price) & ",  "
			 
			 'NO QOH MESSAGE IN COMMENT
			 If part_availability < 1 Then
		 		sql=sql & "'No QOH available for this part at order post.',  "
			 Else
 			 	if instructions > "" then
					sql=sql & "'" & instructions & "',  "
				else
					sql=sql & "'',  "
				end if	
			 End if
			 
			 sql=sql & "'" & DeliveryType & "')  ;"
			
			'Response.Write sql + "<br><br>" 
			objConn.Execute sql
			objConn.Close
			Set objConn = Nothing
			
			If part_availability > 0 Then
				SubstrackQOH part_id,quantity
			End if
		
		oRS.MoveNext
		
	Loop
	
	oRS.Close
	Set oRS = Nothing

	Set oRS_specials = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.clients_cart.shop_auto_id, dbo.clients_cart.shop_client_code, dbo.clients_cart.shop_client_user, dbo.clients_cart.shop_product_id,  dbo.clients_cart.shop_quantity, dbo.clients_cart.shop_part_price, dbo.clients_cart.shop_part_number FROM dbo.clients_cart WHERE dbo.clients_cart.shop_product_id in ( select dbo.prespecials.specials from dbo.prespecials ) and  dbo.clients_cart.shop_client_user = '" + Session("MM_Username") + "' ;"
	
	oRS_specials.Open strSQL, MM_overseaspr_STRING
	
	client_code 		=	UserClientCode(Session("MM_Username"))
	client_user 		=	Session("MM_Username")
	
	// In case order has comment 
	//next_order_number	=	CStr(NextOrderNumber())
	loop_counter        =   0
	
	Do While Not oRS_specials.EOF
	
		loop_counter = loop_counter + 1
		
		if loop_counter = 1 then
			next_order_number	=	CStr(NextOrderNumber())
			
			// New code for comments ...
			if trim(instructions) > "" then
				// Create ZZZNOF
					
				// Get next order number reset counter from 0 to 1
					
				part_id 			=	"ZZZNOF"
				quantity			=	0
				part_price			=	0
					
				'INSERT COMMENT ...
				Set objConn = Server.CreateObject("ADODB.Connection")
				objConn.Open MM_overseaspr_STRING
				sql="INSERT INTO dbo.clients_orders"
				 sql=sql & "( order_number,   "
				 sql=sql & "order_part,   "
				 sql=sql & "order_client,  " 
				 sql=sql & "order_user,   "
				 sql=sql & "order_qty,   "
				 sql=sql & "order_part_id,"   
				 sql=sql & "order_status,  " 
				 sql=sql & "order_date,   "
				 sql=sql & "order_type,   "
				 sql=sql & "item_price,   "
				 sql=sql & "comments,  "
				 sql=sql & "deliv_type )  "
				 sql=sql & "VALUES ( '" & next_order_number & "', "  
				 sql=sql & "'" & part_id & "',   "
				 sql=sql & "'" & client_code & "',   "
				 sql=sql & "'" & Session("MM_Username") & "',   "
				 sql=sql & CStr(quantity) & ", "  
				 sql=sql & "'" & part_id & "',   "
				 sql=sql & "'O',   "
				 sql=sql & "'" & CStr(Now()) & "',  "
				 sql=sql & "'" & DeliveryType & "',  "
				 sql=sql & CStr(part_price) & ",  "
				 sql=sql & "'" & instructions & "',  "
				 sql=sql & "'" & DeliveryType & "')  ;"
				
				objConn.Execute sql
				objConn.Close
				Set objConn = Nothing
				
			end if
			//

			
		End if
		if loop_counter = 22 then
			next_order_number	=	CStr(NextOrderNumber())
			loop_counter        =   1
		End if	
		 
		part_availability	=	PartAvailability(oRS_specials.Fields.Item("shop_part_number"))
		part_id 			=	oRS_specials.Fields.Item("shop_product_id")
		quantity			=	CStr(oRS_specials.Fields.Item("shop_quantity"))
		part_price			=	CStr(oRS_specials.Fields.Item("shop_part_price"))
			
		'INSERT PART INTO CLIENTS ORDER
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open MM_overseaspr_STRING
		sql="INSERT INTO dbo.clients_orders"
		 sql=sql & "( order_number,   "
		 sql=sql & "order_part,   "
		 sql=sql & "order_client,  " 
		 sql=sql & "order_user,   "
		 sql=sql & "order_qty,   "
		 sql=sql & "order_part_id,"   
		 sql=sql & "order_status,  " 
		 sql=sql & "order_date,   "
		 sql=sql & "order_type,   "
		 sql=sql & "item_price,   "
		 sql=sql & "comments,  "
		 sql=sql & "deliv_type )  "
		 sql=sql & "VALUES ( '" & next_order_number & "', "  
		 sql=sql & "'" & part_id & "',   "
		 sql=sql & "'" & client_code & "',   "
		 sql=sql & "'" & Session("MM_Username") & "',   "
		 sql=sql & CStr(quantity) & ", "  
		 sql=sql & "'" & part_id & "',   "
		 sql=sql & "'O',   "
		 sql=sql & "'" & CStr(Now()) & "',  "
		 sql=sql & "'" & DeliveryType & "',  "
		 sql=sql & CStr(part_price) & ",  "
		 
		 'NO QOH MESSAGE IN COMMENT
		 If part_availability < 1 Then
			sql=sql & "'No QOH available for this part at order post.',  "
		 Else
			sql=sql & "'',  "
		 End if
		 
		 sql=sql & "'" & DeliveryType & "')  ;"
		
		'Response.Write sql + "<br><br>" 
		objConn.Execute sql
		objConn.Close
		Set objConn = Nothing
		
		If part_availability > 0 Then
			SubstrackQOH part_id,quantity
		End if
		
		oRS_specials.MoveNext
		
	Loop
	oRS_specials.Close
	Set oRS_specials = Nothing
	
	'RESET ROWS ON CLIENTS_CART FOR THIS USER.
	ResetCart(client_user)
	
	Session("Last_Order_Number") = next_order_number
	
END SUB

FUNCTION FinalPrice(part)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.prespecials.especial FROM dbo.prespecials WHERE dbo.prespecials.specials = '" + part + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
		
	If Not oRS.EOF Then
		FinalPrice = "<strong style='color:red'>" + CurrencyConvert(oRS.Fields.Item("especial")) + "</strong>"
	Else
		oRS.Close
		' LUIS ESTO LO CAMBIE  user_level '3' = precio mas barato que es field_4  !!!!!   GGG
'		Response.Write Session("user_level")
		FinalPrice = 0
		If Session("user_level") = "3" Then
			strSQL = "SELECT field_3,field_4 FROM partmst1_distinct WHERE field_1 = '" + part + "' ;"
			oRS.Open strSQL
			if CStr(oRS.Fields.Item("field_4")) > "" then
				FinalPrice = CurrencyConvert(oRS.Fields.Item("field_4"))
			elseif CStr(oRS.Fields.Item("field_3")) > "" then
				FinalPrice = CurrencyConvert(oRS.Fields.Item("field_3"))			
			else
				FinalPrice = 0
			end if	
		Else
			strSQL = "SELECT field_3,field_4 FROM partmst1_distinct WHERE field_1 = '" + part + "' ;"
			oRS.Open strSQL
			if CStr(oRS.Fields.Item("field_3")) > "" then
				FinalPrice = CurrencyConvert(oRS.Fields.Item("field_3"))
			elseif CStr(oRS.Fields.Item("field_4")) > "" then
				FinalPrice = CurrencyConvert(oRS.Fields.Item("field_4"))			
			else
				FinalPrice = 0
			end if	
		End if
		oRS.Close
		Set oRS = Nothing
	End if
		
END FUNCTION

SUB ChangePassword(user,pwd)
	If len(user) > 0 and len(pwd) > 0 Then
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open MM_overseaspr_STRING
		' LO CAMBIE USABA user  comillas daban problemas
		sql="UPDATE dbo.users SET user_pwd = '" + pwd + "' WHERE dbo.users.user_auto_id = '" + Session("MM_Username2") + "'  ; "
		objConn.Execute sql
	End if
END SUB

SUB LanguageText(header,language)

END SUB

SUB AccountStatementDisplay()

	Set oRS = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "SELECT dbo.invoice.cust_inv,dbo.invoice.invamt,dbo.invoice.inv_date,dbo.invoice.payment_amt,dbo.invoice.payment_date,dbo.invoice.disc1,dbo.invoice.disc2,dbo.invoice.cust,dbo.invoice.inv,dbo.invoice.idcol,dbo.users.user_name,dbo.users.address,dbo.users.address2,dbo.users.city,dbo.users.tel,dbo.users.fax FROM dbo.invoice,dbo.users WHERE dbo.invoice.cust = dbo.users.user_auto_id and dbo.invoice.cust = '" + Session("MM_Username") + "' ORDER BY dbo.invoice.inv_date ASC,dbo.invoice.cust_inv ASC ;"

	oRS.Open strSQL, MM_overseaspr_STRING
	
	Response.Write "<table class = 'tablas' width='1000px' border='0' cellpadding='0' cellspacing='0'  >"
	Response.Write "  <tr>"
	Response.Write "    <td colspan='7' align='left'><img src='images/oiclogo2.gif' border='0'> </td>"
	Response.Write "  </tr>"
	Response.Write "  <tr>"
	Response.Write "    <td colspan='7' align='center'><H2>" + Lang("estado_de_cuenta_actual") + "</H2></td>"
	Response.Write "  </tr>"
	Response.Write "  <tr>"
	Response.Write "    <td colspan='7' align='center'>" + Lang("order_through_our_web") + ": http://www.overseaspr.com</td>"
	Response.Write "  </tr>"
	Response.Write "  <tr>"
	Response.Write "    <td colspan='7' align='center'>PO Box 364951 San Juan PR 00936-4951</td>"
	Response.Write "  </tr>"
	Response.Write "  <tr>"
	Response.Write "    <td colspan='7' align='center'>GANGES #9 URB. EL PARAISO, RIO PIEDRAS PR 00926</td>"
	Response.Write "  </tr>"
	Response.Write "  <tr>"
	Response.Write "    <td colspan='7' align='center'>TEL (787) 751-4036     Fax (787) 765-6735</td>"
	Response.Write "  </tr>"

	Response.Write "  <tr height='50px' align='center'><td colspan='7' ><strong>PRINT ACCOUNT STATEMENT HERE -> :&nbsp;&nbsp;&nbsp;</strong><a style='width:100%'  href='JavaScript:window.print();'><img src='printer_icon.gif' border='0' width='17' height='17' alt='Print Version'/></a>&nbsp;</td></tr>"


	' ALTERNATE ROW COLOR VARIABLES
	color1 = "#CCCCCC"
	color2 = "#FFFFFF"
	current_color = color1
	rc = 1

	'TOTAL EN EL CART
	cart_total = 0
	line_cum = 0
	current_cum = 0	
	if oRS.EOF then
		Response.Write "<h5>" + Lang("no_existe_record") + "</h5> "
	else

		Response.Write "  <tr>"
		Response.Write "    <td width='100%' align='center' colspan='7'><h1>" + CStr(oRS.Fields.Item("user_name")) + "</h1></td>"
		Response.Write "  </tr>"
		if oRS.Fields.Item("address") > "" then	
			Response.Write "  <tr>"
			Response.Write "    <td width='100%' align='center' colspan='7'>" + CStr(oRS.Fields.Item("address")) + "</td>"
			Response.Write "  </tr>"
		end if
		if oRS.Fields.Item("address2") > "" then	
			Response.Write "  <tr>"
			Response.Write "    <td width='100%' align='center' colspan='7'>" + CStr(oRS.Fields.Item("address2")) + "</td>"
			Response.Write "  </tr>"
		end if
		if oRS.Fields.Item("city") > "" then	
			Response.Write "  <tr>"
			Response.Write "    <td width='100%' align='center' colspan='7'>" + CStr(oRS.Fields.Item("city")) + "</td>"
			Response.Write "  </tr>"
		end if
		Response.Write "  <tr><td colspan='7'>&nbsp;</td></tr>"				
		Response.Write "  <tr><td colspan='7'><hr></td></tr>"
		Response.Write "  <tr>"
		Response.Write "    <td width='15%' height='20px' align='center' bgcolor='#F2F2F2'><strong>" + Lang("num_factura") + "</strong></td>"
		Response.Write "    <td width='20%' align='right' bgcolor='#F2F2F2'><strong>" + Lang("cant_fact_fecha") + "</strong></td>"
		Response.Write "    <td width='20%' align='right' bgcolor='#F2F2F2'><strong>" + Lang("pagos_recibos_fecha") + "</strong></td>"
		Response.Write "    <td width='15%' align='right' bgcolor='#F2F2F2'><strong>" + Lang("descuentos") + "</strong></td>"
		Response.Write "    <td width='15%' align='right' bgcolor='#F2F2F2'><strong>" + Lang("balance") + "</strong></td>"
		Response.Write "    <td width='10%' align='right' bgcolor='#F2F2F2'><strong>" + Lang("sub_total") + "</strong></td>"
		Response.Write "    <td width='5%' align='center' bgcolor='#F2F2F2'><strong>" + Lang("statement_days") + "</strong></td>"
		Response.Write "  </tr>"
	
		Do While Not oRS.EOF

			order_date = DateValue( CStr(FormatDateTime(oRS.Fields.Item("inv_date"),2)))
			DaysOld = DateDiff("d", order_date, Now)

			if Not IsNull(oRS.Fields.Item("invamt")) and len(oRS.Fields.Item("invamt")) > 0 then
				invamt = oRS.Fields.Item("invamt")
			else
				invamt = 0
			end if
			if Not IsNull(oRS.Fields.Item("payment_amt")) and len(oRS.Fields.Item("payment_amt")) > 0 then
				
				payment_amt = oRS.Fields.Item("payment_amt")
				' format for cents with 1 decimal ...
				pos = InStr(payment_amt, ".")
				if pos > 0 then
				cents = mid( payment_amt, pos + 1 ,2 )
				if len(cents) = 1 then cents = cents + "0"
				payment_amt = FormatNumber( mid( payment_amt, 1 , pos ) + cents ,2 )
				end if
				
			else
				payment_amt = 0	
			end if
			if Not IsNull(oRS.Fields.Item("disc1")) and len(oRS.Fields.Item("disc1")) > 0 then
				disc1 = oRS.Fields.Item("disc1")
			else
				disc1 = 0	
			end if
			if Not IsNull(oRS.Fields.Item("disc2")) and len(oRS.Fields.Item("disc2")) > 0 then
				disc2 = oRS.Fields.Item("disc2")
			else
				disc2 = 0	
			end if
			
			if trim(oRS.Fields.Item("payment_date")) <> "0" and Not IsNull(oRS.Fields.Item("payment_date")) and len(oRS.Fields.Item("payment_date")) > 0 then
				payment_date = CStr(oRS.Fields.Item("payment_date"))
			else
				payment_date = "&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;&nbsp;&nbsp;"	
			end if
	
			Response.Write "  <tr>"
			Response.Write "    <td align='center' bgcolor='" + current_color + "' height='25'>" + CStr(oRS.Fields.Item("inv")) + "</td>"
			Response.Write "    <td align='right' bgcolor='" + current_color + "'><strong>" + CStr(CurrencyConvert(oRS.Fields.Item("invamt"))) + "&nbsp;&nbsp;" + CStr(oRS.Fields.Item("inv_date")) + "</strong></td>"
				Response.Write "    <td align='right' style='color:green' bgcolor='" +  current_color + "'><strong>" + CStr(payment_amt) + "&nbsp;&nbsp;" + payment_date + "</strong></td>"
			if disc1 > 0 then
				if disc2 > 0 then
					Response.Write "    <td align='right' style='color:purple' bgcolor='"+current_color + "'><strong>" + CStr(CurrencyConvert(oRS.Fields.Item("disc1"))) + CStr(CurrencyConvert(oRS.Fields.Item("disc2"))) + "</strong></td>"			
				else
					Response.Write "    <td align='right' style='color:purple' bgcolor='"+current_color + "'><strong>" + CStr(CurrencyConvert(oRS.Fields.Item("disc1"))) + "</strong></td>"
				end if	
			elseif disc2 > 0 then
				Response.Write "    <td align='right' style='color:purple' bgcolor='" +  current_color + "'><strong>" + CStr(CurrencyConvert(oRS.Fields.Item("disc2"))) + "</strong></td>"			
			else	
				Response.Write "    <td align='right' style='color:purple' bgcolor='" +  current_color + "'>&nbsp;</td>"
			end if	

			Response.Write "    <td align='right' bgcolor='" +  current_color + "'><strong>" + CStr(CurrencyConvert( invamt - payment_amt - disc1 - disc2)) + "</strong></td>"
			cart_total = cart_total + invamt - payment_amt - disc1 - disc2			
			Response.Write "    <td align='right' bgcolor='" +  current_color + "'>" + CStr(CurrencyConvert( cart_total)) + "</td>"			
			Response.Write "    <td align='center' bgcolor='" +  current_color + "'>" + CStr(DaysOld) + "</td>"						
			Response.Write "  </tr>"

			if DaysOld >= 0 and DaysOld <= 30 then
				current_cum = current_cum + invamt - payment_amt - disc1 - disc2			
			elseif DaysOld > 30 and DaysOld <= 60 then
				over_30 = over_30 + invamt - payment_amt - disc1 - disc2			
'				past_due = past_due + over_30				
			elseif DaysOld > 60 and DaysOld <= 90 then
				over_60 = over_60 + invamt - payment_amt - disc1 - disc2			
'				past_due = past_due + over_60				
			elseif DaysOld > 90 and DaysOld <= 120 then
				over_90 = over_90 + invamt - payment_amt - disc1 - disc2			
'				past_due = past_due + over_90
			elseif DaysOld > 120 then
				over_120 = over_120 + invamt - payment_amt - disc1 - disc2			
'				past_due = past_due + invamt - payment_amt - disc1 - disc2			
			end if

			oRS.MoveNext
			
			If current_color = color2 then current_color = color1 else current_color = color2
			rc = rc + 1
		
		Loop
		past_due = over_30 + over_60 + over_90 + over_120
			
		Response.Write "    <td align='center'>&nbsp;</td>"														
		Response.Write "<tr><td colspan='6'><hr></td></tr>"
		Response.Write "    <td align='center'>&nbsp;</td>"												
		Response.Write "</table><br><br>"

		Response.Write "<table class = 'tablas' align='center' width='600px' border='0' cellpadding='0' cellspacing='0'  >"						
		Response.Write "  <tr>"
		Response.Write "    <td width='100px' align='center'><strong>CURRENT</strong></td>"
		Response.Write "    <td width='100px' align='center'><strong>PAST DUE</strong></td>"
		Response.Write "    <td width='100px' align='center'><strong>OVER 30</strong></td>"
		Response.Write "    <td width='100px' align='center'><strong>OVER 60</strong></td>"
		Response.Write "    <td width='100px' align='center'><strong>OVER 90</strong></td>"
		Response.Write "    <td width='100px' align='center'><strong>OVER 120</strong></td>"
		Response.Write "  </tr>"

		Response.Write "  <tr>"
		Response.Write "    <td width='100px' align='center' height='25'><h2>" + CStr(CurrencyConvert(current_cum)) + "</h2></td>"
		Response.Write "    <td width='100px' align='center'><h2>" + CStr(CurrencyConvert(past_due)) + "</h2></td>"
		Response.Write "    <td width='100px' align='center'><h2>" + CStr(CurrencyConvert(over_30)) + "</h2></td>"
		Response.Write "    <td width='100px' align='center'><h2>" + CStr(CurrencyConvert(over_60)) + "</h2></td>"
		Response.Write "    <td width='100px' align='center'><h2>" + CStr(CurrencyConvert(over_90)) + "</h2></td>"
		Response.Write "    <td width='100px' align='center'><h2>" + CStr(CurrencyConvert(over_120)) + "</h2></td>"
		Response.Write "  </tr>"

'		Response.Write "    <td align='center'>&nbsp;</td>"														
		Response.Write "<tr><td colspan='6'><hr></td></tr>"
'		Response.Write "    <td align='center'>&nbsp;</td>"												
'		Response.Write "  <tr>"
		
		Response.Write "    <td colspan='6' align='center'>Nuestro término de pago es Neto 30 Dias</td>"
		Response.Write "  </tr>"
		Response.Write "    <td align='center'>&nbsp;</td>"												
		Response.Write "</table>"
	
	End if
END SUB
' Function to get user name ( Uses MM_UserName as input parm )
FUNCTION GetUserName(userid)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.users.user_name FROM dbo.users WHERE dbo.users.user_auto_id = '" + userid + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
	GetUserName = CStr(oRS.Fields.Item("user_name"))
END FUNCTION

' Function to get user name ( Uses MM_UserName as input parm )
FUNCTION GetCurrentSales(userid)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	StringDate 				= CStr(Date())
	
	'IF WANT TO DISPLAY ORDERS BETWEEN YESTERDAY AND TODAY ADD THIS DATE INTO WHERE.
	StringDateLessOneDate	= CStr(DateAdd("d", -1, Now()))
	
	'GGG old SQL
	'SELECT     MIN(order_number) AS minorder, MAX(order_number) AS maxorder
	'FROM         clients_orders
	'WHERE     (order_date >= '2011-02-27')

	'LRO current SQL 2/26/2011
	strSQL = "SELECT MIN(dbo.clients_orders.order_number)as minorder,MAX(dbo.clients_orders.order_number)as maxorder FROM dbo.clients_orders WHERE dbo.clients_orders.order_date >= '" + StringDate + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
	'LRO  2/26/2011
	'CHANGE FROM EOF TO VALIDATE IF COUNT IS 0 BECAUSE ALLWAYS WILL RETURN ONE ROW WITH A TOTAL.
	If oRS.EOF then
	  GetCurrentSales = StringDate
	else
	  If oRS.Fields.Item("minorder") > "" AND oRS.Fields.Item("maxorder") > "" Then
		GetCurrentSales = "Ordenes: " + CStr( oRS.Fields.Item("maxorder") - oRS.Fields.Item("minorder") + 1) 
	  Else
		GetCurrentSales = StringDate
	  End if
	End if
END FUNCTION

FUNCTION PartPhoto(PartNum)
	
	photo_exist = false
	
	If PartNum = ""  Then
		EXIT FUNCTION
	End If
	
	PartNumJPG = Server.MapPath("./parts_images/" & PartNum & ".jpg")
	PartNumPNG = Server.MapPath("./parts_images/" & PartNum & ".png")
	PartNumGIF = Server.MapPath("./parts_images/" & PartNum & ".gif")

	'response.Write ToRootedVirtual(PartNumJPG)

	Dim objFSO
	
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	If objFSO.FileExists( PartNumJPG ) Then
		photo_exist = true
		PartPhoto = "<a href='" + ToRootedVirtual(PartNumJPG) + "/parts_images/" + PartNum + ".jpg' target='_new'>" + PartNum + "</a>"
	End if
	
	If objFSO.FileExists( PartNumPNG ) Then
		photo_exist = true
		PartPhoto = "<a href='" + ToRootedVirtual(PartNumPNG) + "/parts_images/" + PartNum + ".png' target='_new'>" + PartNum + "</a>"
	End if
	
	If objFSO.FileExists( PartNumGIF ) Then
		photo_exist = true
		PartPhoto = "<a href='<a href='" + ToRootedVirtual(PartNumGIF) + "/parts_images/" + PartNum + ".gif' target='_new'>" + PartNum + "</a>"
	End if
	
	If photo_exist = false Then
		PartPhoto = PartNum 
	End if
	
	Set objFSO = Nothing   
END FUNCTION

Function ToRootedVirtual(relativePath)
    Dim applicationMetaPath : applicationMetaPath = Request.ServerVariables("APPL_MD_PATH")
    Dim instanceMetaPath : instanceMetaPath = Request.ServerVariables("INSTANCE_META_PATH")
    Dim rootPath : rootPath = Mid(applicationMetaPath, Len(instanceMetaPath) + Len("/ROOT/"))
    ToRootedVirtual = rootPath
	'Response.Write( ToRootedVirtual )
End Function

' Code for SQL Injection attack ...
FUNCTION GetSecureVal(param)
	If IsEmpty(param) or param = "" then
		GetSecureVal = param
		Exit Function
	End If
		
	If IsNumeric(param) Then
		GetSecureVal = CLng(param)
	Else
		GetSecureVal = Replace( CStr(param),"'","''")
	End If	
END FUNCTION

SUB Image_exist(part)
	
	if Not IsEmpty(part) then
		'UPDATE QOH
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open MM_overseaspr_STRING
		sql="UPDATE dbo.partmst1_distinct SET image_exist = '1' WHERE dbo.partmst1_distinct.field_1 = '" + part + "' ;"
		objConn.Execute sql
		objConn.Close
		Set objConn = Nothing
	end if
			
END SUB

%>