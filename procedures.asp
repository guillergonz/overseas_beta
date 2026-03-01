<%
SUB CheckCredentials()

	
	' ONLY AFTER SUCESSFUL LOGIN, TO CREATE SESSION CREDENTIALS
	If Len(Session("MM_Username")) > 0 Then
		Set oRS = Server.CreateObject("ADODB.Recordset")
		' IF USER ACCESS IS VALIDATED IN TABLE, PLEASE ADD COLUMN IN WHERE CONDITION
		strSQL = "SELECT users.user_name, users.user_level, users.user_pwd, users.lang, users.discount, statetax, citytax FROM dbo.users WHERE user_auto_id = '" + Session("MM_Username") + "' ;"
	
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
				' GGG 06/2015
				if Session("user_level") = "9" then
					Session("user_level") = "1"
				end if	
				
				' GGG 8/8/2013
				If LEN(oRS.Fields.Item("discount")) > 0 And IsNumeric( trim(oRS.Fields.Item("discount"))) Then 
					Session("MM_discount") = FormatNumber( trim(oRS.Fields.Item("discount")),0)
				Else
					Session("MM_discount") = 0
				End If	

				Session("MM_CityTax") = "N"
				If IsNull(oRS.Fields.Item("citytax")) or oRS.Fields.Item("citytax") = "" Then
					Session("MM_CityTax") = "N"
				ElseIf LEN(oRS.Fields.Item("citytax")) = 1 Then
					Session("MM_CityTax") = oRS.Fields.Item("citytax")
				End If		

				Session("MM_StateTax") = "N"
				If IsNull(oRS.Fields.Item("statetax")) or oRS.Fields.Item("statetax") = "" Then
					Session("MM_StateTax") = "N"
				ElseIf LEN(oRS.Fields.Item("statetax")) = 1 Then
					Session("MM_StateTax") = oRS.Fields.Item("statetax")
				End If		
				
				Session("MM_CatID") = 1	
				Session("MM_CatalogID") = 1
				Session("MM_Titulo") = ""
				Session("MM_auto") = false
				Session("SQLSearch") = ""
									
				Response.Redirect("part_search.asp")
				'Response.Redirect("oic_news.asp")
				
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
				If language = "E" Then Lang = "LIQUIDATION" Else Lang = "LIQUIDACION" End If

			Case "Specials"
				If language = "E" Then Lang = "SPECIALS" Else Lang = "ESPECIALES" End If
			'INDEX
			Case "password"
				If language = "E" Then Lang = "Verificar Credenciales de Acceso!" Else Lang = "Invalid user or password!" End If
			' TEMPLATE HEADERS
			Case "fecha"
				If language = "E" Then Lang = "Date" Else Lang = "Fecha" End If
			Case "completar_orden"
				If language = "E" Then Lang ="Check Out" Else Lang = "Terminar Orden" End if
			Case "ACCOUNTSTATEMENT"
				If language = "E" Then Lang ="Account Statement" Else Lang = "Estado de Cuenta" End if	
			Case "salir"
				If language = "E" Then Lang ="Log Out" Else Lang = "Salir del Sistema" End if
			Case "preparado"
				If language = "E" Then Lang = "CART" Else Lang = "PREPARADO" End if
			Case "entre_10_piezas"
				If language = "E" Then Lang = "Enter up to 10 items" Else Lang = "Entre hasta 10 piezas" End if

			Case "entre_5_piezas"
				If language = "E" Then Lang = "Enter up to 5 items" Else Lang = "Entre hasta 5 piezas" End if

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
				If language = "E" Then Lang = "PART NUMBER" Else Lang = "PIEZA" End if
			Case "reemplazo"
				If language = "E" Then Lang = "REPLACEMENT" Else Lang = "REEMPLAZO" End if
			Case "descripcion"
				If language = "E" Then Lang = "Description" Else Lang = "Descripción" End if
			Case "precio"
				If language = "E" Then Lang = "Price" Else Lang = "Precio" End if
			Case "disponible"
				If language = "E" Then Lang = "AVL." Else Lang = "DISPONIBLE" End if
			Case "cant_ordenar"
				If language = "E" Then Lang = "Qty" Else Lang = "Cant" End if
			Case "categoria"
				If language = "E" Then Lang = "CATEGORY" Else Lang = "CATEGORIA" End if

			'PAGE NAVIGATOR
			Case "primera_pagina"
				If language = "E" Then Lang = "First Page" Else Lang = "Primera Página" End if
			Case "pagina_anterior"
				If language = "E" Then Lang = "Previous Page" Else Lang = "Página Anterior" End if
			Case "proxima_pagina"
				If language = "E" Then Lang = "Next Page" Else Lang = "Próxima Página" End if
			Case "ultima_pagina"
				If language = "E" Then Lang = "Last Page" Else Lang = "Ultima Página" End if

			'NAVIGATOR STATUS
			Case "desplegando"
				If language = "E" Then Lang = "Record" Else Lang = "Récord" End if
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
				If language = "E" Then Lang = "Part" Else Lang = "Pieza" End if
			Case "ordenado"
				If language = "E" Then Lang = "Quantity" Else Lang = "Cantidad" End if
			Case "no_existen_piezas"
				If language = "E" Then Lang = "NO PARTS" Else Lang = "NO EXISTEN PIEZAS" End if
			Case "num_order"
				If language = "E" Then Lang = "ORDER NUM." Else Lang = "NUM. ORDEN" End if
				
			Case "item_added"
				If language = "E" Then Lang = "Item added to cart ..." Else Lang = "Pieza añadida (cart) ..." End if
					
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
				If language = "E" Then
				 Lang = "The order will be procesed the next labor day if submited after 4:00 pm !" 
				Else
				
					'If Session("MM_Username") = "Z099" then 
						Lang = "<p><span style='color:red'>SAME DAY:</span> <span style='color:black'>Ordenes recibidas en o antes 9:50 AM serán entregadas el mismo día.(Ver Mapa)</span></p>" + "<p><span style='color:red'>NEXT DAY:</span> <span style='color:black'>Ordenes recibidas en o antes 4:00 PM serán entregadas el próximo día de entrega.</span></p>" + "<p>Ordenes recibidas sabado o domingo serán entregadas el próximo día laborable.</p>" + "<p>ORDENES <span style='color:red'>SAME DAY - NEXT DAY</span> <span style='color:black'>Entrega gratis Valor final $76.00 o más" + "<p>ORDENES <span style='color:red'>SAME DAY - NEXT DAY</span> <span style='color:black'>Ordenes entre $30.00 - $75.00 Costo de entrega $3.00"
				 	'Else
				 		'Lang = "Ordenes recibidas después de las 4:00 pm serán procesadas el próximo día laborable!"
			     	'End if
					
				End If
			Case "Instrucciones"
				If language = "E" Then Lang = "Instructions:" Else Lang = "Instrucciones" End if	
			
			Case "metodo_envio"
				If language = "E" Then Lang = "Delivery Method" Else Lang = "Método de Entrega" End if
			
			Case "entrega"
				If language = "E" Then Lang = "Delivery next delivery working day (Next Day)" Else Lang = "Entrega próximo día de entrega (Next Day)" End if
			
			Case "entregamismodia"
				If language = "E" Then Lang = "Delivery Same Day (Until 9:50 AM) (Same Day)" Else Lang = "Entrega mismo día (Hasta las 9:50 AM) (Same Day)" End if
				
			Case "entregamismodiaSD"
				If language = "E" Then Lang = "Delivery Same Day (Next labor day)" Else Lang = "Entrega mismo día (Próximo día laborable)" End if	

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

			Case "addtocart"
				If language = "E" Then Lang = "Add to cart" Else Lang = "Añadir a la cesta" End if

			
			
			
		End Select

END FUNCTION

FUNCTION SameDayFlag(user_auto_id)
	SameDayFlag = "N"  ' Default return value
	
	If Len(user_auto_id) = 0 Then
		Exit Function
	End If
	
	Dim oRS
	Set oRS = Server.CreateObject("ADODB.Recordset")
	
	Dim strSQL
	strSQL = "SELECT samedayflag FROM dbo.users WHERE user_auto_id = '" & Replace(user_auto_id, "'", "''") & "'"
	
	On Error Resume Next
	oRS.Open strSQL, MM_overseaspr_STRING
	
	If Err.Number = 0 Then
		If Not oRS.EOF Then
			If Not IsNull(oRS.Fields.Item("samedayflag")) Then
				Dim tempFlag
				tempFlag = Trim(CStr(oRS.Fields.Item("samedayflag")))
				If tempFlag = "Y" Or tempFlag = "N" Then
					SameDayFlag = tempFlag
				End If
			End If
		End If
		oRS.Close
	End If
	On Error Goto 0
	
	Set oRS = Nothing
END FUNCTION

SUB ResetUserSession()
	'RESET SESSION VARIABLES
	Session("MM_UserId_Multi") = ""
	
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
	
	' response.write "SQL script " + sqlScript
	
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
			'Response.Write "<table class='table'><tr><td>"
			NavigationStatus RS, nPage, nPageCount,RCount,nItemsPerPage
			'Response.Write "</td></tr></table>"
			
			if nPageCount > 1 then
				NavigationControls nPage,nPageCount,"part_search.asp"
			end if
			
			RS.AbsolutePage = nPage
			' nPageCount, RCount, nItemsPerPage
			DisplayResult RS,nPage,nPageCount, RCount, nItemsPerPage
			
			if nPageCount > 1 then
				NavigationControls nPage,nPageCount,"part_search.asp"
			end if
			
		end if
		RS.Close
		
	end if
	
	Set RS = Nothing
End Sub

Sub DisplayResult(RecordsetObject,nPage,nPageCount, RCount, nItemsPerPage)
	' PARTS COUNTER, VALIDATE TO BE USED ON N PAGES.
	If nPage = 1 Then counter = 1 Else counter = (nItemsPerpage * nPage) - nItemsPerpage + 1
	Response.Write "<table class=""table table-striped"" >"
	Response.Write "<tr>"
	Response.Write "<td  align=""left"">&nbsp;</td>"
	Response.Write "<td  align=""left"">&nbsp;" + ucase(Lang("num_pieza")) + "</td>"
	Response.Write "<td  align=""left"">&nbsp;" + ucase(Lang("reemplazo")) + "</td>"
	Response.Write "<td align=""center"">" + ucase(Lang("descripcion")) + "</td>"
	Response.Write "<td  >&nbsp;" + ucase(Lang("precio")) + "</td>"
	Response.Write "<td  >&nbsp;" + ucase(Lang("disponible")) + "</td>"
	Response.Write "<td  align=""center"">" + ucase(Lang("cant_ordenar")) + "</td>"
	'" + ucase(Lang("categoria")) + "
	Response.Write "<td></td>"
	Response.Write "<td align=""center""></td>"	

	Do While Not (RecordsetObject.EOF or RecordsetObject.AbsolutePage <> nPage)
		
		Response.Write "<tr>"
		Response.Write "<td  align='center' >" + CStr(counter) + "</td>" 
		Response.Write "<td  align='center'>" + PartPhoto(Trim(RecordsetObject("field_1").Value))  + "</td>"
		Response.Write "<td  align=""center"">"
		If len(Trim(RecordsetObject("replacement").Value)) > 0 Then
			Response.Write ("<label>" + Trim(RecordsetObject("replacement").Value) + "</label>") 
		Else
			Response.Write "N/A"
		End if
		
		Response.Write "</td>"
		' GGG add </span> Response.Write "<td style=""max-width:200px"" align=""center""><span style=""max-width:200px"">" 
		Response.Write "<td class=""wordwrap"" align=""center"">" 
		If Session("lang") = "E" Then
			If len(trim(RecordsetObject("english_version").Value)) > 0 Then
				Response.Write RecordsetObject("english_version").Value + "</td>" 
			Else
				Response.Write RecordsetObject("field_2").Value + "</td>" 
			End if	
		Else	
			Response.Write RecordsetObject("field_2").Value + "</td>" 
		End if

 		Session("MM_price_type") = ""
 		Session("MM_Discounted") = 0
		
		'if trim(RecordsetObject("field_1").Value)= "S1140005TJ" then
'			stop
'		end if	
		
		' FINALPRICE CALCULATES Session("MM_Discounted")
		Dim fp
		Response.Write "<td  align=""right"">"
		fp = CStr(FinalPrice(RecordsetObject("field_1").Value))
		
		
		
		if Session("MM_discount") > 0 then 
		
			if LEN(Session("MM_Multi_Username")) > 0 AND Session("MM_Multi_Username") = "COUNTER" then
				Response.write "<label title='Descuento' >&nbsp;</label>"
            else
				if Session("MM_price_type") = "E" then
					Response.write "<label style='color:red' title='Descuento' >*" + CStr(Session("MM_Discounted")) + "</label>"
				else
					Response.write "<label title='Descuento' >*" + CStr(Session("MM_Discounted")) + "</label>"
				end if	
            
	    	end if
            
        else
			Response.Write trim(fp)
		end if
		
		Response.write "</td>" 
		
		if RecordsetObject("field_5").Value > 50 then
			Response.Write "<td  align='center'>50+</td>"
		else	
			Response.Write "<td  align='center'>" + trim(RecordsetObject("field_5").Value) + "</td>"
		end if	
		Response.Write "<td  align=""center""><input id='cant_ordenar_" + CStr(counter) + "' class='search_box' size='4' value='1' onFocus='this.select()'/>  </td>"
		
		if Session("lang") = "S" then
			Response.Write "<td style='font-size:10px' align=""center"">" + RecordsetObject("familia_descripcion").Value +  "</td>" 
		else
			Response.Write "<td style='font-size:10px' align=""center"">" + RecordsetObject("family_description").Value +  "</td>" 
		end if	
		Response.Write "<td  align='center'>"
		
		if fp <> "$0.00" then
			if PartAvailability(RecordsetObject("field_1").Value) > 0 Then
				if CheckPartInCart(RecordsetObject("field_1").Value) then
					' GGG href='#' NO VA 1/16/2016 -> part_search.asp# 
					
					Response.Write "<div id='add_cart_" + CStr(counter) + "' style='vertical-align:top' ><a title='" + lang("addtocart") + "' class='btn btn-default'  onclick='ValidateCart(""" + RecordsetObject("field_1").Value + """,""cant_ordenar_" + CStr(counter) + """);'><span class='glyphicon glyphicon-shopping-cart'  onclick=""javascript:document.getElementById('add_cart_" + CStr(counter) + "') "" ></span></a></div>"
					
					'Response.Write "<div id='add_cart_" + CStr(counter) + "' style='vertical-align:top' ><a class='btn btn-default'  onclick='ValidateCart(""" + RecordsetObject("field_1").Value + """,""cant_ordenar_" + CStr(counter) + """);'><img src='images/add-to-cart.gif' width='15'  height='15' onclick=""javascript:document.getElementById('add_cart_" + CStr(counter) + "') "" /></a></div>"
					
					
				end if	
			end if
		end if
		
		Response.Write "</tr>"
		
		RecordsetObject.MoveNext
		counter = counter + 1
	Loop
	Response.Write "</table>"
End Sub

Sub NavigationControls(nPage,nPageCount,nPageName)
	Response.Write "<table class=""table table-condensed"" width=""100%""><tr>"
	Response.Write "<td width=""25%"" align=""center""><a class=""btn btn-default btn-sm"" href=""/" + nPageName + "?Page=1"" >&nbsp;" + Lang("primera_pagina") + "&nbsp;</a>&nbsp;&nbsp;</td>"
	If nPage - 1 > 0 Then
		Response.Write "<td width=""25%"" align=""center""><a class=""btn btn-default btn-sm"" href=""/" + nPageName + "?Page=" + CStr(nPage - 1) + """ >&nbsp;" + Lang("pagina_anterior") + "&nbsp;</a>&nbsp;&nbsp;</td>"
	Else
		Response.Write "<td width=""25%"" align=""center"">&nbsp;</td>"
	End If
	If nPage + 1 < nPageCount Then
		Response.Write "<td width=""25%"" align=""center""><a class=""btn btn-default btn-sm"" href=""/" + nPageName + "?Page=" + CStr(nPage + 1) + """ >&nbsp;" + Lang("proxima_pagina") + "&nbsp;</a>&nbsp;&nbsp;</td>"
	Else
		Response.Write "<td width=""25%"" align=""center"">&nbsp;</td>"
	End If
	Response.Write "<td width=""25%"" align=""center""><a class=""btn btn-default btn-sm"" href=""/" + nPageName + "?Page=" + CStr(nPageCount) + """ >&nbsp;" + Lang("ultima_pagina") + "&nbsp;</a></td>"
	Response.Write "</tr></table>"
End Sub

Sub NavigationStatus(RecordSetObject, nPage, nPageCount,RCount,nItemsPerPage)
	Response.Write "<table class='table table-striped' width='100%' ><tr><td>"
	Response.Write "&nbsp;&nbsp;&nbsp;" + Lang("desplegando") + ": <strong>" + CStr(RecordsPerPageStatus (RecordsetObject, nPage,nPageCount,nItemsPerPage,RCount)) + "</strong> " + Lang("de") + " <strong>" + CStr(RCount) + "</strong></td><td>"
	Response.Write Lang("total_piezas") + ": <strong style=""color:black"">" + CStr(RCount) + "</strong></td><td align=""right"">"
	Response.Write Lang("pagina") + ": <strong>" + CStr(nPage) + "</strong> " + Lang("de") + " <strong>" + CStr(nPageCount) + "</strong>&nbsp;&nbsp;&nbsp;</td></tr></table>"
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
if ShowDescription = "" then ShowDescription = "N"
If PartsOnCart() Then
	
	Set oRS = Server.CreateObject("ADODB.Recordset")
	if Session("MM_UserName") = "B001" AND LEN(Session("MM_UserId_Multi")) > 0 then	

		strSQL = "SELECT shop_auto_id, shop_part_number, shop_part_price, shop_auto_id,shop_quantity FROM dbo.clients_cart WHERE client_multi = '" +  Session("MM_UserId_Multi") + "' AND shop_client_user = '" + Session("MM_Username") + "' ORDER BY shop_order_date,shop_part_number ;"

	else

		strSQL = "SELECT shop_auto_id, shop_part_number, shop_part_price, shop_auto_id,shop_quantity FROM dbo.clients_cart WHERE shop_client_user = '" + Session("MM_Username") + "' ORDER BY shop_order_date,shop_part_number ;"

	end if
	oRS.Open strSQL, MM_overseaspr_STRING
	
	Response.Write "<table class='table table-condensed' >"
	Response.Write "<tr  >"
	Response.Write "<td>" + Lang("pieza") + "</td>"
	if ShowDescription = "Y" then
	 Response.Write "<td>" + Lang("descripcion") + "</td>"
	end if
	Response.Write "<td  align='center' width='80px' >" + Lang("cant_ordenar") + "</td>"
	Response.Write "<td  align='center' >" + Lang("precio") + "</td>"
	if ShowDescription = "Y" then
		Response.Write "<td align='center' >Total</td>"
	end if	
	Response.Write "<td align='center' >&nbsp;</td>"
	Response.Write "</tr>"
	
	' ALTERNATE ROW COLOR VARIABLES
	'color1 = "#CCCCCC"
	'color2 = "#FFFFFF"
	'current_color = color1
	rc = 1

	'TOTAL EN EL CART
	cart_total = 0
	
	
	
	Do While Not oRS.EOF
	
		if ShowDescription = "Y" then
			' Record set to get description in english or spanish for part item ...
			
			shop_part_number  = Trim(oRS.Fields.Item("shop_part_number"))
			
			Set oRS2 = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT field_2, english_version FROM dbo.partmst1_distinct WHERE field_1 = '" + shop_part_number  + "' ;"
			oRS2.Open strSQL, MM_overseaspr_STRING
			ldescription = ""
			If Not IsNull(oRS2.Fields.Item("english_version")) AND Session("lang") = "E" Then
				if Trim(oRS2.Fields.Item("english_version")) > "" then
					ldescription = Trim(oRS2.Fields.Item("english_version"))
				else
					ldescription = (oRS2.Fields.Item("field_2"))
				end if		
			else
				if oRS2.EOF then
					ldescription = "M i s s i n g"
				else
					ldescription = (oRS2.Fields.Item("field_2"))
				end if	
			End if
			oRS2.Close
			Set oRS2 = Nothing
			
			'  mod ends here .
		end if
		
		
		
		Response.Write "  <tr>"
		If ShowDescription = "Y" then
			' bgcolor='" +  current_color + "'
			Response.Write "    <td  style='font-size:12px' >" + Trim(oRS.Fields.Item("shop_part_number")) + "</td>"
			Response.Write "    <td style='font-size:12px'  >" + ldescription + "</td>"
		Else
			Response.Write "    <td  style='font-size:10px'>" + Trim(oRS.Fields.Item("shop_part_number")) + "</td>"
		End If
					
		' allow update quantity order in shopping cart
		if ShowDescription = "Y" then

			'<a title='Cambiar cantidad ordenada' id='UpdCart' href='#' onclick='UpdCart(""" + CStr(oRS.Fields.Item("shop_auto_id")) + """);'><img valign='baseline' src='images/20x20-save.png' border='0'  /></a>

			Response.Write "<td  align='center' ><input maxlength='3' onchange='UpdCart(""" + CStr(oRS.Fields.Item("shop_auto_id")) + """);'  style='margin-left:5px;margin-right:5px;text-align:center;width:40px' name='shop_quantity" & oRS.Fields("shop_auto_id").Value & "' type='text' id='shop_quantity" & oRS.Fields("shop_auto_id").Value & "' value='" + CStr(oRS.Fields.Item("shop_quantity")) + "'></td>"
		else
			Response.Write "<td style='font-size:10px' align='center' width='40px' >" + trim(oRS.Fields.Item("shop_quantity")) + "</td>"
		end if
		
		
		if LEN(Session("MM_Multi_Username")) > 0 AND ucase(Session("MM_Multi_Username")) = "COUNTER" then
			' do not show price
			Response.Write "<td>&nbsp;</td>"
			
		else
				if PartOnSale( CStr(oRS.Fields.Item("shop_part_number")) , oRS.Fields.Item("shop_part_price") ) then
				
					if ShowDescription = "Y" then			
						Response.Write "    <td style='font-size:12px' align='right' color='RED'><font color='red'>" + CStr(CurrencyConvert(oRS.Fields.Item("shop_part_price"))) + "</font></td>"
					else
						Response.Write "    <td style='font-size:10px' align='right' color='RED'><font color='red'>" + CStr(CurrencyConvert(oRS.Fields.Item("shop_part_price"))) + "</font></td>"
					end if
												
				else
				
					if ShowDescription = "Y" then			
						Response.Write "    <td style='font-size:12px' align='right' >" + CStr(CurrencyConvert(oRS.Fields.Item("shop_part_price"))) + "</td>"
					else
						Response.Write "    <td style='font-size:10px' align='right' >" + CStr(CurrencyConvert(oRS.Fields.Item("shop_part_price"))) + "</td>"			
					end if
					
				end if
		end if		
		cart_total = (oRS.Fields.Item("shop_part_price") * oRS.Fields.Item("shop_quantity")) + cart_total
		cart_subtotal = (oRS.Fields.Item("shop_part_price") * oRS.Fields.Item("shop_quantity"))
		
		
		
		if LEN(Session("MM_Multi_Username")) > 0 AND ucase(Session("MM_Multi_Username")) = "COUNTER" then
			if ShowDescription = "Y" then
				Response.Write "<td>&nbsp;</td>"
			end if
		else	
			if ShowDescription = "Y" then
				Response.Write "<td style='font-size:12px;padding-right:10px' align='right'>" + CStr(CurrencyConvert(cart_subtotal)) + "</td>"
			end if
		end if
		
		if ShowDescription = "Y" then
			Response.Write "<td><a title='Eliminar' id='DelFromCartbutton' onclick='DelFromCart(""" + CStr(oRS.Fields.Item("shop_auto_id")) + """);'  ><img src='images/del.gif' alt='delete' width='15' height='15' style='vertical-align:top' ></a></td></tr>"
		else
			Response.Write "<td><a title='Eliminar' id='DelFromCartbutton' onclick='DelFromCart(""" + CStr(oRS.Fields.Item("shop_auto_id")) + """);'   ><img src='images/del.gif' alt='delete' width='15' height='15' style='vertical-align:top' ></a></td></tr>"
		end if	
		oRS.MoveNext
		
				
		'If current_color = color2 then current_color = color1 else current_color = color2
		rc = rc + 1
    
	Loop

	
	' <hr>
	Response.Write "</table>"

	Response.Write "<table class='table table-condensed' width='100%'>"

	if ShowDescription = "Y" then
		
		if ucase(Session("MM_Multi_Username")) = "COUNTER" then
			Response.Write "<tr valign='top'><td width='80%'></td><td align='right'>&nbsp;</td><td width='5%'></td>"
			Response.Write "<td align='right' >&nbsp;</td><td width='5%'></td></tr>"
		else
			Response.Write "<tr valign='top'><td width='80%'></td><td align='right'>Subtotal</td><td width='5%'></td>"
			Response.Write "<td align='right' >" + CurrencyConvert(cart_total) + "</td><td width='5%'></td></tr>"
		end if
	Else
		
		if ucase(Session("MM_Multi_Username")) = "COUNTER" then
			Response.Write "<tr valign='top'><td width='55%'></td><td align='right'>&nbsp;</td><td></td>"
			Response.Write "<td align='right' >&nbsp;</td><td width='5%'></td></tr>"
		else
			Response.Write "<tr valign='top'><td width='55%'></td><td align='right'>Subtotal</td><td></td>"
			Response.Write "<td align='right' >" + CurrencyConvert(cart_total) + "</td><td width='20%'></td></tr>"
		end if

	End If
	
	If Session("MM_CityTax") = "Y" AND Session("MM_Statetax") = "Y" Then
	 vCityTax  = round( (cart_total * .01) ,2)
	 'vStateTax = round( (cart_total * .06) ,2)
	 vStateTax = round( (cart_total * .105) ,2)
	 cart_total = cart_total + vCityTax + vStateTax
	End IF
	
	if ShowDescription = "Y" then
		
		If Session("MM_CityTax") = "Y" AND ucase(Session("MM_Multi_Username")) <> "COUNTER" Then	
			Response.Write "<tr valign='top'><td colspan='2' align='right'>Impuesto Municipal</td><td width='5%'></td><td align='right'>" + CurrencyConvert(vCityTax) + "</td><td>&nbsp;</td></tr>"
		End if
		'procedure consider both for calc and visibility of tax lines in cart
		If Session("MM_CityTax") = "Y" AND Session("MM_StateTax") = "Y" AND ucase(Session("MM_Multi_Username")) <> "COUNTER" Then	
			Response.Write "<tr valign='top'><td colspan='2' align='right'>Impuesto Estatal</td><td width='5%'></td><td align='right'>" + CurrencyConvert(vStateTax) + "</td><td>&nbsp;</td></tr>"
		End if
	
		If ucase(Session("MM_Multi_Username")) <> "COUNTER" Then
			Response.Write "<tr valign='top'><td colspan='2' align='right'>TOTAL</td><td width='5%'>&nbsp;</td><td align='right' >" + CurrencyConvert(cart_total) + "</td><td></td></tr>"
		end if
		
	end if
			

	If Request.ServerVariables("SCRIPT_NAME") = "part_search.asp"  or Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/part_search.asp" Then
		Response.Write " <tr><td colspan='5'><h3><center><a style='width:100%' href='cart.asp'>" + Lang("completar_orden") + "</a></center></h3></td></tr>"
	End if

	
	Response.Write "</table>"
	
	
	
    oRS.Close
	Set oRS = Nothing


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

FUNCTION GetCatalog(idcatalog)
Set oRS = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT catalogname FROM dbo.OIC_Catalog WHERE idcatalog = " + CStr(idcatalog) 
oRS.Open strSQL, MM_overseaspr_STRING
If Not oRS.EOF AND Not IsNull(oRS.Fields.Item("catalogname")) then
	GetCatalog = trim(oRS.Fields.Item("catalogname"))
else
	GetCatalog = "N/A"
End if
oRS.Close
Set oRS = Nothing
END FUNCTION

FUNCTION GetSubCategoryName(categoryid,subcategoryid)
Set oRS2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT Category FROM dbo.OIC_Subcategory WHERE CategoryId = " + CStr(categoryid) + " AND SubCategoryId = " + CStr(subcategoryid)
oRS2.Open strSQL, MM_overseaspr_STRING
If Not oRS2.EOF AND Not IsNull(oRS2.Fields.Item("Category")) then
	GetSubCategoryName = trim(oRS2.Fields.Item("Category"))
else
	GetSubCategoryName = "N/A"
End if
oRS2.Close
Set oRS2 = Nothing
END FUNCTION

FUNCTION GetCategoryName(categoryid)
Set oRS2 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT Category FROM dbo.OIC_Category WHERE CategoryId = " + CStr(categoryid) 
oRS2.Open strSQL, MM_overseaspr_STRING
If Not oRS2.EOF AND Not IsNull(oRS2.Fields.Item("Category")) then
	GetCategoryName = trim(oRS2.Fields.Item("Category"))
else
	GetCategoryName = "N/A"
End if
oRS2.Close
Set oRS2 = Nothing
END FUNCTION

SUB AddPartToCategory (catid,categoryid,subcategoryid,partno,titulo,nota)
If LEN(catid) > 0 AND LEN(categoryid) > 0 AND LEN(partno) > 0 Then
	Set objConn5 = Server.CreateObject("ADODB.Connection")
	objConn5.Open MM_overseaspr_STRING
	
	sql="INSERT INTO dbo.OIC_PartsPerCategory(idcatalog,categoryid,subcategoryid,image,link,partno,titulo,nota) "
	sql=sql & " VALUES "
	sql=sql & "(" & catid & ","
	sql=sql & categoryid & ","
	sql=sql & subcategoryid & ","
	sql=sql & "'" & partno & "',"
	sql=sql & "'" & partno & "',"
	sql=sql & "'" & partno & "',"
	sql=sql & "'" & titulo & "',"
	sql=sql & "'" & nota & "' )  ;" 	  
						   
	objConn5.Execute sql
	sError = err.description
	objConn5.Close 
	Set objConn5 = Nothing
	
	If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error añadiendo esta pieza al catálogo: <br><br>" & sError )
	End if
End If
END SUB

SUB DelPartFromCategory(uid)
If uid > 0 Then
	Set objConn5 = Server.CreateObject("ADODB.Connection")
	objConn5.Open MM_overseaspr_STRING
	
	sql="DELETE FROM dbo.OIC_PartsPerCategory WHERE uniqueidcol = " + CStr(uid)
	objConn5.Execute sql
	sError = err.description
	objConn5.Close 
	Set objConn5 = Nothing
	
	If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error eliminando esta pieza del catálogo: <br><br>" & sError )
	End if
End If
END SUB

SUB AddCatalog(catname)
If LEN(catname) > 0 Then
	Set objConn5 = Server.CreateObject("ADODB.Connection")
	objConn5.Open MM_overseaspr_STRING
	
	sql="INSERT INTO dbo.OIC_Catalog(CatalogName,catalogpdfname,orden) "
	sql=sql & " VALUES ('" & catname & "',"
	sql=sql & "'" & catname & "',"
	sql=sql & 100 & " )  ;" 	  
						   
	objConn5.Execute sql
	sError = err.description
	objConn5.Close 
	Set objConn5 = Nothing
	
	If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error añadiendo este catálogo: <br><br>" & sError )
	End if
End If
END SUB

SUB AddCat(category)
If LEN(catid) > 0 AND LEN(categoryid) > 0 AND LEN(titulo) > 0 Then
	Set objConn5 = Server.CreateObject("ADODB.Connection")
	objConn5.Open MM_overseaspr_STRING
	
	sql="INSERT INTO dbo.OIC_Category(category) VALUES ('" & category & "');" 
	objConn5.Execute sql
	sError = err.description
	objConn5.Close 
	Set objConn5 = Nothing
	
	If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error añadiendo esta categoria: <br><br>" & sError )
	End if
End If
END SUB

SUB AddCategory(catid,categoryid,titulo)
If LEN(catid) > 0 AND LEN(categoryid) > 0 AND LEN(titulo) > 0 Then
	Set objConn5 = Server.CreateObject("ADODB.Connection")
	objConn5.Open MM_overseaspr_STRING
	
	sql="INSERT INTO dbo.OIC_PartsPerCategory(idcatalog,categoryid,titulo,partno,image,link,nota) "
	sql=sql & " VALUES "
	sql=sql & "(" & catid & ","
	sql=sql & categoryid & ","
	sql=sql & "'" & titulo & "','','','','');" 	  
						   
	objConn5.Execute sql
	sError = err.description
	objConn5.Close 
	Set objConn5 = Nothing
	
	If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error añadiendo esta categoria o sección: <br><br>" & sError )
	End if
End If
END SUB

SUB AddToCart(pieza,cant)

	dim counter
	counter = 0

	Set oRS2 = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT COUNT(*) AS counter FROM dbo.clients_cart WHERE shop_client_code = '" + client_code + "' AND shop_product_id = '" + CStr(pieza) + "' " 
	oRS2.Open strSQL, MM_overseaspr_STRING
	If Not oRS2.EOF AND Not IsNull(oRS2.Fields.Item("counter")) then
		counter = trim(oRS2.Fields.Item("counter"))
	else
		counter = 0
	End if
	oRS2.Close
	Set oRS2 = Nothing

	if counter = 0 then

		
		'client_multi new code for B&V
		if Session("MM_UserName") = "B001" AND LEN(Session("MM_UserId_Multi")) > 0 then
			client_multi = Session("MM_UserId_Multi")
		else
			client_multi = Session("MM_UserName")
		end if
		
		client_code = UserClientCode(Session("MM_Username"))
		user_id	 = Session("MM_Username")
		available   = PartAvailability(pieza)
		cant		= CStr(cant)
		precio	  = CurrencyConvert(PricePerUser(pieza))
		
		Set objConn5 = Server.CreateObject("ADODB.Connection")
		objConn5.Open MM_overseaspr_STRING
		
		sql="INSERT INTO clients_cart (shop_client_code,shop_client_user,shop_product_id,shop_quantity,shop_part_price,shop_part_number,available,shop_order_date,client_multi,shop_type) "
		sql=sql & " VALUES "
		sql=sql & "('" & client_code & "',"
		sql=sql & "'" & user_id & "',"
		sql=sql & "'" & pieza & "',"
		sql=sql & "'" & cant & "',"
		sql=sql & "'" & precio & "',"
		sql=sql & "'" & pieza & "',"
		sql=sql & "'" & available & "',"
		sql=sql & "'" & CStr(Now()) & "',"
		sql=sql & "'" & client_multi & "',"
		sql=sql & "'W' )  ;" 	  
							   
		objConn5.Execute sql
		sError = err.description
		'objConn5.Close 
		'Set objConn5 = Nothing
		
		If len(sError) > 0 Then
			Response.Write("Ocurrió el siguiente error añadiendo esta pieza al cart: <br><br>" & sError )
		End if
		
	End If
	
END SUB

SUB DelFromCart(auto_id)
	Set objConn5 = Server.CreateObject("ADODB.Connection")
	objConn5.Open MM_overseaspr_STRING
	
	if Session("MM_UserName") = "B001" AND LEN(Session("MM_UserId_Multi")) > 0 then
		client_multi = Session("MM_UserId_Multi")
	else
		client_multi = Session("MM_UserName")
	end if
	
	if Session("MM_UserName") = "B001" AND LEN(Session("MM_UserId_Multi")) > 0 then
		sql = "DELETE FROM clients_cart WHERE client_multi = '" + client_multi + "' AND shop_auto_id = " + CStr(auto_id) + " AND shop_client_user = '" + Session("MM_Username") + "' ;"
	else	
		' pk shop_auto_id + " AND shop_client_user = '" + Session("MM_Username") + "' ;"
		sql = "DELETE FROM clients_cart WHERE shop_auto_id = " + CStr(auto_id) 
	end if
	
	objConn5.Execute sql
	'response.Write sql
	sError = err.description
	'objConn5.Close 
	'Set objConn5 = Nothing
	
	If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error eliminando esta pieza al cart: <br><br><strong>" & sError & "</strong")
	End if
	
END SUB

SUB UpdateCart( auto_id, qty )

	'response.write("LLEGO")
	
	if Session("MM_UserName") = "B001" AND LEN(Session("MM_UserId_Multi")) > 0 then
		client_multi = Session("MM_UserId_Multi")
	else
		client_multi = Session("MM_UserName")
	end if
	
	Set objConn5 = Server.CreateObject("ADODB.Connection")
	objConn5.Open MM_overseaspr_STRING
	
	if Session("MM_UserName") = "B001" AND LEN(Session("MM_UserId_Multi")) > 0 then
		sql = "UPDATE clients_cart SET shop_quantity = " & qty & "WHERE client_multi = '" + client_multi + "' AND shop_auto_id = " + CStr(auto_id) + " AND shop_client_user = '" + Session("MM_Username") + "' ;"
	else
		sql = "UPDATE clients_cart SET shop_quantity = " & qty & "WHERE shop_auto_id = " + CStr(auto_id) + " AND shop_client_user = '" + Session("MM_Username") + "' ;"
	end if
	objConn5.Execute sql
	'response.Write sql
	sError = err.description
	
	'objConn5.Close 
	'Set objConn5 = Nothing
	
	If len(sError) > 0 Then
		Response.Write("Ocurrió el siguiente error eliminando esta pieza al cart: <br><br><strong>" & sError & "</strong")
	End if
	
END SUB

FUNCTION UserClientCode(user_name)
	Set oRecSet = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT user_client_code FROM dbo.users WHERE user_auto_id = '" + Session("MM_Username") + "' ;"
	oRecSet.Open strSQL, MM_overseaspr_STRING
	If Not oRecSet.EOF Then UserClientCode = CStr(oRecSet.Fields.Item("user_client_code"))
	oRecSet.Close
	Set oRecSet = Nothing
END FUNCTION

FUNCTION PartAvailability(part)

	Set oRecSet = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.partmst1_distinct.field_5 FROM dbo.partmst1_distinct WHERE dbo.partmst1_distinct.field_1 = '" + part + "' ;"
	oRecSet.Open strSQL, MM_overseaspr_STRING
	If Not oRecSet.EOF Then
	 PartAvailability = CInt(oRecSet.Fields.Item("field_5"))
	Else
	 PartAvailability = 0
	End If 
	oRecSet.Close
	Set oRecSet = Nothing
	
END FUNCTION

FUNCTION PricePerUser(pieza)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT dbo.prespecials.especial FROM dbo.prespecials WHERE dbo.prespecials.specials = '" + pieza + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
		
	If Not oRS.EOF Then
		PricePerUser = CStr(oRS.Fields.Item("especial"))
	Else
	
		Set oRS = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT price, regular FROM dbo.prod_liqui WHERE field_1 = '" + part + "' ;"
		oRS.Open strSQL, MM_overseaspr_STRING
		If Not oRS.EOF Then
			PricePerUser = CStr(oRS.Fields.Item("price"))
		Else
	
			oRS.Close
			' user_level '3' = precio mas barato que es field_4  !!!!!   GGG
			' 2015 user_level = 1 3 4 6
			'  field_3, field_4, field_7, field_8
			
			If Session("user_level") = "3" Then
				strSQL = "SELECT field_4 FROM partmst1_distinct WHERE field_1 = '" + pieza + "' ;"
				oRS.Open strSQL, MM_overseaspr_STRING
				if Not oRS.EOF then
					PricePerUser = CStr(oRS.Fields.Item("field_4"))
				Else
					PricePerUser = "0"
				End If	
			
			Elseif Session("user_level") = "1" Then
				strSQL = "SELECT field_3 FROM partmst1_distinct WHERE field_1 = '" + pieza + "' ;"
				oRS.Open strSQL, MM_overseaspr_STRING
				if Not oRS.EOF then
					PricePerUser = CStr(oRS.Fields.Item("field_3"))
				Else
					PricePerUser = "0"
				End If	
				
			Elseif Session("user_level") = "4" Then
				strSQL = "SELECT field_7 FROM partmst1_distinct WHERE field_1 = '" + pieza + "' ;"
				oRS.Open strSQL, MM_overseaspr_STRING
				if Not oRS.EOF then
					PricePerUser = CStr(oRS.Fields.Item("field_7"))
				Else
					PricePerUser = "0"
				End If	
				
			Elseif Session("user_level") = "6" Then
				strSQL = "SELECT field_8 FROM partmst1_distinct WHERE field_1 = '" + pieza + "' ;"
				oRS.Open strSQL, MM_overseaspr_STRING
				if Not oRS.EOF then
					PricePerUser = CStr(oRS.Fields.Item("field_8"))
				Else
					PricePerUser = "0"
				End If	
				
						
				
			End if
			
						
			'If Session("user_level") = "3" Then
'				strSQL = "SELECT field_4 FROM partmst1_distinct WHERE field_1 = '" + pieza + "' ;"
'				oRS.Open strSQL, MM_overseaspr_STRING
'				if Not oRS.EOF then
'					PricePerUser = CStr(oRS.Fields.Item("field_4"))
'				Else
'					PricePerUser = "0"
'				End If	
'			Else
'				strSQL = "SELECT field_3 FROM partmst1_distinct WHERE field_1 = '" + pieza + "' ;"
'				oRS.Open strSQL, MM_overseaspr_STRING
'				if Not oRS.EOF then
'					PricePerUser = CStr(oRS.Fields.Item("field_3"))
'				Else
'					PricePerUser = "0"
'				End If	
'			End if
			
			
			
			' Session("MM_UserName") = "Z099"
			If Session("MM_discount") > 0 AND FormatNumber(PricePerUser,2) > 0 then
				
				PricePerUser = CStr(FormatNumber( PricePerUser * (1 - (Session("MM_discount"))/100) ,2))
				
			Else
			
				PricePerUser = CStr(FormatNumber( PricePerUser,2 ))
				
			End If	
			
		End If	

	End if
END FUNCTION

FUNCTION CurrencyConvert(precio)
	
	if len(precio) > 0 then
		CurrencyConvert = FormatCurrency(precio, 2)
	end if
		
END FUNCTION

FUNCTION CheckPartInCart(part)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT Count(shop_part_number) as countrecs FROM dbo.clients_cart WHERE shop_part_number = '" + part + "' AND shop_client_user = '" + Session("MM_Username") + "' ;"
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
	
	strSQL = "SELECT order_number,order_date FROM dbo.clients_orders WHERE order_user = '" + Session("MM_Username") + "' GROUP BY order_number, order_date ORDER BY order_date desc, order_number desc ;"
	
	oRS.Open strSQL, MM_overseaspr_STRING
	
	Response.Write "<table class='table table-striped'  >"
	Response.Write "  <tr valign='top'>"
	Response.Write "    <td style='font-size:9px' width='25%' height='20px' align='center' ><strong>" + Lang("num_order") + "</strong></td>"
	<!--bgcolor='#F2F2F2'-->
	Response.Write "    <td style='font-size:9px' width='50%' align='center' ><strong>" + Lang("fecha") + "</strong></td>"
	Response.Write "    <td style='font-size:9px' width='25%' align='center' ><strong>TOTAL</strong></td>"
	Response.Write "  </tr>"

	' ALTERNATE ROW COLOR VARIABLES
	'color1 = "#F2F2F2"
	'color2 = "#FFFFFF"
	'current_color = color1
	rc = 1
	
	LASTordernumber = ""
	Do While Not oRS.EOF
	
		if trim(oRS.Fields.Item("order_number")) <> LASTordernumber then
			Response.Write "  <tr height='20px'>"
			
			Response.Write "    <td ><a style='font-size:12px' href='#' OnClick='location.href=""order_detail.asp?o=" + CStr(oRS.Fields.Item("order_number")) + """;'>" + CStr(oRS.Fields.Item("order_number")) + "</a></td>"
			
			Response.Write "    <td style='font-size:9px'  >" + FechaOrden(CStr(oRS.Fields.Item("order_number"))) + "</td>"
			Response.Write "    <td style='font-size:9px'  >" + OrderTotal(CStr(oRS.Fields.Item("order_number"))) + "</td>"
			Response.Write "  </tr>"
		end if
		
		LASTordernumber = trim(oRS.Fields.Item("order_number"))
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
	if not oRS.eof AND Not IsNull(oRS("order_date")) then
		FechaOrden = trim(oRS("order_date"))	
	else
		FechaOrden = ""
	end if	
	'FechaOrden = CStr(FormatDateTime(oRS.Fields.Item("order_date"),2))
	
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
	
	strSQL = "SELECT shop_auto_id  FROM dbo.clients_cart WHERE shop_client_user = '" + Session("MM_username") + "' ;"
	
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
    Dim cmd, nextNum
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = MM_overseaspr_STRING
    cmd.CommandText = "GetNextOrderNumber"
    cmd.CommandType = 4 'adCmdStoredProc
    
    ' Create output parameter
    cmd.Parameters.Append cmd.CreateParameter("@NextOrderNum", 3, 2, 4) 'adInteger, adParamOutput
    
    cmd.Execute
    NextOrderNumber = cmd.Parameters("@NextOrderNum").Value
    
    Set cmd = Nothing
END FUNCTION

SUB SubstrackQOH(part,qty)
	'SELECT CURRENT QOH FOR PART
	parts_available = PartAvailability(part)
	
	'UPDATE QOH
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open MM_overseaspr_STRING
	sql="UPDATE dbo.partmst1_distinct SET field_5 = " + CStr(CInt(parts_available) - CInt(qty)) + "  WHERE dbo.partmst1_distinct.field_1 = '" + part + "' ;"
	objConn.Execute sql
	objConn.Close
	Set objConn = Nothing
	
END SUB

SUB ResetCart(user)
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open MM_overseaspr_STRING
	
	if user = "B001" then
		
		sql="DELETE FROM dbo.clients_cart WHERE client_multi = '" + Session("MM_UserId_Multi") + "' ;"
		
	else	
		
		sql="DELETE FROM dbo.clients_cart WHERE shop_client_user = '" + user + "' ;"
	
	end if
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
		counter = 1
		Do While NOT available_order_number
			
			available_order_number = AvailableOrderNumber()
			' Required to get out of the infine loop
			If Not available_order_number AND counter > 0 Then
				' Force update InUse = N is case of loop due to record lock
				NextOrderNumber()
				available_order_number = AvailableOrderNumber()
			end if
				
			'Check counter variable > 10
			If available_order_number Then
				
				AddOrder DeliveryType,instructions
				Response.Write "<br><br><hr><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + Lang("enviada_correctamente")
				
				if Session("Last_Order_Number") > 0 then
					' iframe use target = _parent
					Response.Write ( "<hr><br>&nbsp;<a style='width:100%'  href='/order_detail.asp?o=" + Session("Last_Order_Number") + "' target='_parent'>" + Lang("VerOrden") + "</a>" )
		
				End if
		
			End if			
			counter = counter + 1
		Loop
	End if
		
END SUB

SUB AddOrder(DeliveryType,instructions)
	
	Set oRS = Server.CreateObject("ADODB.Recordset")
	
	if ucase(Session("MM_Username")) = "B001" then
		strSQL = "SELECT shop_auto_id, shop_client_code, shop_client_user, shop_product_id,  shop_quantity, shop_part_price, shop_part_number FROM dbo.clients_cart WHERE shop_product_id not in ( select dbo.prespecials.specials from dbo.prespecials ) and  client_multi = '" + Session("MM_UserId_Multi") + "' ;"	
	else
		strSQL = "SELECT shop_auto_id, shop_client_code, shop_client_user, shop_product_id,  shop_quantity, shop_part_price, shop_part_number FROM dbo.clients_cart WHERE shop_product_id not in ( select dbo.prespecials.specials from dbo.prespecials ) and  shop_client_user = '" + Session("MM_Username") + "' ;"
	end if
	
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
				Dim sinst
				
				sinst = trim(instructions)
				
				while LEN(sinst) > 0	
				
						// Get next order number reset counter from 0 to 1
							
						part_id 	= "ZZZNOF"
						quantity   = 0
						part_price = 0
						
						
							
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
						 
						 sql=sql & "zcustomer,  "
						 
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
						 sql=sql & "'" & mid(sinst,1,25) & "',  "
						 
						 sql=sql & "'" & Session("MM_UserId_Multi") & "',   "
						 
						 sql=sql & "'" & DeliveryType & "')  ;"
						
						objConn.Execute sql
						objConn.Close
						Set objConn = Nothing
				
					
					sinst = mid(sinst,26)
					
					
				wend
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
			 
			 sql=sql & "zcustomer,  "
			  
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
		 		sql=sql & "'No QOH available',  "
			 Else
 			 	'if instructions > "" then
				'	sql=sql & "'" & instructions & "',  "
				'else
					sql=sql & "'',  "
				'end if	
			 End if
			 
			 sql=sql & "'" & Session("MM_UserId_Multi") & "',   "
			 
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
	if ucase(Session("MM_Username")) = "B001" then
		strSQL = "SELECT shop_auto_id, shop_client_code, shop_client_user, shop_product_id,  shop_quantity, shop_part_price, shop_part_number FROM dbo.clients_cart WHERE shop_product_id in ( select dbo.prespecials.specials from dbo.prespecials ) and client_multi = '" + Session("MM_UserId_Multi") + "' ;"
		
	else
		strSQL = "SELECT shop_auto_id, shop_client_code, shop_client_user, shop_product_id,  shop_quantity, shop_part_price, shop_part_number FROM dbo.clients_cart WHERE shop_product_id in ( select dbo.prespecials.specials from dbo.prespecials ) and  shop_client_user = '" + Session("MM_Username") + "' ;"
	end if
	
	
	
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
				 
				 sql=sql & "zcustomer,  "
				 
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
				 
				 sql=sql & "'" & Session("MM_UserId_Multi") & "',   "
				 
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
		 
		 sql=sql & "zcustomer,  "
		 
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
			sql=sql & "'No QOH available',  "
		 Else
			sql=sql & "'',  "
		 End if
		 
		 sql=sql & "'" & Session("MM_UserId_Multi") & "',   "
		 
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

'	Falta prod_liqui
'   Añadir MM_price_type = 'L' Liquidacion / 'E' Especial / 'P' Precio asignado a usuario
'
	Session("MM_price_type") = ""
				
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT price, regular FROM dbo.prod_liqui WHERE field_1 = '" + part + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
	If Not oRS.EOF Then
	
		Session("MM_price_type") = "L"		
		FinalPrice = "<strong style='color:red'>&nbsp;L&nbsp;" + CurrencyConvert(oRS.Fields.Item("price").value) + "</strong>"
	Else
		
	 	Set oRS = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT dbo.prespecials.especial FROM dbo.prespecials WHERE dbo.prespecials.specials = '" + part + "' ;"
		oRS.Open strSQL, MM_overseaspr_STRING
		If Not oRS.EOF Then
		
			
			
			Session("MM_price_type") = "E"		
			FinalPrice = "<strong style='color:red'>&nbsp;E&nbsp;" + CurrencyConvert(oRS.Fields.Item("especial").value) + "</strong>"
			' MUST calculate Session("MM_Discounted") here !
			Session("MM_Discounted") = FormatNumber(oRS.Fields.Item("especial").value,2)  
			
		Else
			
			' 2015 user_level = 1 3 4 6
			'  field_3, field_4, field_7, field_8
			
			Set oRS = Server.CreateObject("ADODB.Recordset")							
			
			'if trim(part)= "S1140005TJ" then
'				stop
'			end if	
			
			If Session("user_level") = "3" Then
				strSQL = "SELECT field_4 FROM partmst1_distinct WHERE field_1 = '" + part + "' ;"
				oRS.Open strSQL, MM_overseaspr_STRING
				if Not oRS.EOF then
					FinalPrice = CurrencyConvert(oRS.Fields.Item("field_4").value)
				Else
					FinalPrice = "0"
				End If	
			
			Elseif Session("user_level") = "1" Then
				strSQL = "SELECT field_3 FROM partmst1_distinct WHERE field_1 = '" + part + "' ;"
				oRS.Open strSQL, MM_overseaspr_STRING
				if Not oRS.EOF then
					FinalPrice = CurrencyConvert(oRS.Fields.Item("field_3").value)
				Else
					FinalPrice = "0"
				End If	
				
			Elseif Session("user_level") = "4" Then
				strSQL = "SELECT field_7 FROM partmst1_distinct WHERE field_1 = '" + part + "' ;"
				oRS.Open strSQL, MM_overseaspr_STRING
				if Not oRS.EOF then
					FinalPrice = CurrencyConvert(oRS.Fields.Item("field_7").value)
				Else
					FinalPrice = "0"
				End If	
				
			Elseif Session("user_level") = "6" Then
				strSQL = "SELECT field_8 FROM partmst1_distinct WHERE field_1 = '" + part + "' ;"
				oRS.Open strSQL, MM_overseaspr_STRING
				if Not oRS.EOF then
					FinalPrice = CurrencyConvert(oRS.Fields.Item("field_8").value)
				Else
					FinalPrice = "0"
				End If	
						
			End if					
		
			'if part = "S1140005TJ" then
'				stop
'			end if	
		
			If Session("MM_discount") > "" AND FinalPrice > 0 then
				
				Session("MM_Discounted") = CurrencyConvert( FinalPrice * (1 - CurrencyConvert(Session("MM_discount"))/100) )
			
			Else
			
				Session("MM_Discounted") = 0
				Session("MM_price_type") = ""
				
			End If
			
			'oRS.Close
			'Set oRS = Nothing
		End if
		
	End If
		
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
	
	strSQL = "SELECT dbo.invoice.cust_inv,dbo.invoice.invamt,dbo.invoice.inv_date,dbo.invoice.payment_amt,dbo.invoice.payment_date,dbo.invoice.disc1,dbo.invoice.disc2,dbo.invoice.cust,dbo.invoice.inv,dbo.users.user_name,dbo.users.address,dbo.users.address2,dbo.users.city,dbo.users.tel,dbo.users.fax FROM dbo.invoice,dbo.users WHERE dbo.invoice.cust = dbo.users.user_auto_id and dbo.invoice.cust = '" + Session("MM_Username") + "' ORDER BY dbo.invoice.inv_date ASC,dbo.invoice.cust_inv ASC ;"

	oRS.Open strSQL, MM_overseaspr_STRING
	
	Response.Write "<table class = 'tablas' width='1000px' border='0' cellpadding='0' cellspacing='0'  >"
	Response.Write "  <tr>"
	Response.Write "    <td colspan='7' align='center'><img src='images/oiclogo2.gif' border='0'> </td>"
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
		Response.Write "    <td width='15%' height='20px' align='center' ><strong>" + Lang("num_factura") + "</strong></td>"
		Response.Write "    <td width='20%' align='right' ><strong>" + Lang("cant_fact_fecha") + "</strong></td>"
		Response.Write "    <td width='20%' align='right' ><strong>" + Lang("pagos_recibos_fecha") + "</strong></td>"
		Response.Write "    <td width='15%' align='right' ><strong>" + Lang("descuentos") + "</strong></td>"
		Response.Write "    <td width='15%' align='right' ><strong>" + Lang("balance") + "</strong></td>"
		Response.Write "    <td width='10%' align='right' ><strong>" + Lang("sub_total") + "</strong></td>"
		Response.Write "    <td width='5%' align='center' ><strong>" + Lang("statement_days") + "</strong></td>"
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
			Response.Write "    <td align='center'  height='25'>" + CStr(oRS.Fields.Item("inv")) + "</td>"
			Response.Write "    <td align='right' ><strong>" + CStr(CurrencyConvert(oRS.Fields.Item("invamt"))) + "&nbsp;&nbsp;" + CStr(oRS.Fields.Item("inv_date")) + "</strong></td>"
				Response.Write "    <td align='right' style='color:green' ><strong>" + CStr(payment_amt) + "&nbsp;&nbsp;" + payment_date + "</strong></td>"
			if disc1 > 0 then
				if disc2 > 0 then
					Response.Write "    <td align='right' style='color:purple' bgcolor='"+current_color + "'><strong>" + CStr(CurrencyConvert(oRS.Fields.Item("disc1"))) + CStr(CurrencyConvert(oRS.Fields.Item("disc2"))) + "</strong></td>"			
				else
					Response.Write "    <td align='right' style='color:purple' bgcolor='"+current_color + "'><strong>" + CStr(CurrencyConvert(oRS.Fields.Item("disc1"))) + "</strong></td>"
				end if	
			elseif disc2 > 0 then
				Response.Write "    <td align='right' style='color:purple' ><strong>" + CStr(CurrencyConvert(oRS.Fields.Item("disc2"))) + "</strong></td>"			
			else	
				Response.Write "    <td align='right' style='color:purple' >&nbsp;</td>"
			end if	

			Response.Write "    <td align='right' ><strong>" + CStr(CurrencyConvert( invamt - payment_amt - disc1 - disc2)) + "</strong></td>"
			cart_total = cart_total + invamt - payment_amt - disc1 - disc2			
			Response.Write "    <td align='right' >" + CStr(CurrencyConvert( cart_total)) + "</td>"			
			Response.Write "    <td align='center' >" + CStr(DaysOld) + "</td>"						
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
		Response.Write "    <td width='100px' align='center' height='25'><h3>" + CStr(CurrencyConvert(current_cum)) + "</h3></td>"
		Response.Write "    <td width='100px' align='center'><h3>" + CStr(CurrencyConvert(past_due)) + "</h3></td>"
		Response.Write "    <td width='100px' align='center'><h3>" + CStr(CurrencyConvert(over_30)) + "</h3></td>"
		Response.Write "    <td width='100px' align='center'><h3>" + CStr(CurrencyConvert(over_60)) + "</h3></td>"
		Response.Write "    <td width='100px' align='center'><h3>" + CStr(CurrencyConvert(over_90)) + "</h3></td>"
		Response.Write "    <td width='100px' align='center'><h3>" + CStr(CurrencyConvert(over_120)) + "</h3></td>"
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
	strSQL = "SELECT user_name FROM dbo.users WHERE user_auto_id = '" + userid + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
	If Not oRS.EOF then
		if Not IsNull(oRS.Fields.Item("user_name").value) then
			GetUserName = oRS.Fields.Item("user_name")
		end if
	Else
		GetUserName = "N/A"
	End If	
END FUNCTION

' Function to get user name ( Uses MM_UserName as input parm )
FUNCTION GetCurrentSales(userid)
	Set oRSSales = Server.CreateObject("ADODB.Recordset")
	StringDate 				= CStr(Date())
		
	
	
	'GGG old SQL
	'SELECT     MIN(order_number) AS minorder, MAX(order_number) AS maxorder
	'FROM         clients_orders
	'WHERE     (order_date >= '2011-02-27')

	'error using order_number char(10) cannot add use order_id
	strSQL = "SELECT MIN(dbo.clients_orders.order_id)as minorderid,MAX(dbo.clients_orders.order_id)as maxorderid FROM dbo.clients_orders WHERE dbo.clients_orders.order_date >= '" + StringDate + "' ;"
	oRSSales.Open strSQL, MM_overseaspr_STRING
	
	If oRSSales.EOF then
	  
	  GetCurrentSales = 0
	  
	else
	
		
		
		'min = (oRSSales.Fields.Item("minorderid").value)
		if IsEmpty(oRSSales.Fields.Item("minorderid").value) or IsNull(oRSSales.Fields.Item("minorderid").value) then 
			min = 0
		else
			min = CDbl(oRSSales.Fields.Item("minorderid").value)
		end if
		if IsEmpty(oRSSales.Fields.Item("maxorderid").value) or IsNull(oRSSales.Fields.Item("maxorderid").value) then  	
			max = 0
		else	
			max = CDbl(oRSSales.Fields.Item("maxorderid").value)
		end if
		oRSSales.close
		Set oRSSales = Nothing 
	
		if min = 0 or max = 0 then
			GetCurrentSales = 0
		else
			
			Set oRSSales = Server.CreateObject("ADODB.Recordset")
			strSQL = "select count( distinct order_number) as counter from dbo.clients_orders where order_id >= " + trim(min) 
			oRSSales.Open strSQL, MM_overseaspr_STRING
			If Not oRSSales.EOF AND LEN(oRSSales("counter")) > 0 then
				GetCurrentSales =  CDbl(oRSSales.Fields.Item("counter").value) 
			Else
				GetCurrentSales = 0
			End If
			oRSSales.close
			Set oRSSales = Nothing   
	
		End If	
		
	End if
	
	

END FUNCTION

FUNCTION GetLastPurchase(userid)
Set oRS5 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT max(order_date) as maxorderD FROM dbo.clients_orders WHERE order_client = '" + trim(userid) + "' "
oRS5.Open strSQL, MM_overseaspr_STRING
If Not IsNull(oRS5("maxorderD")) and Len(oRS5.Fields.Item("maxorderD").value) > 0 then
	GetLastPurchase = CDate(oRS5.Fields.Item("maxorderD")) 
else
	GetLastPurchase = "N/A"
End if
oRS5.close
Set oRS5 = Nothing   
END FUNCTION


FUNCTION GetLastForcedate()
Set oRS5 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT forcedatetime , forceinventoryupd  FROM dbo.counters "
oRS5.Open strSQL, MM_overseaspr_STRING


If Len(oRS5.Fields.Item("forcedatetime").value) > 0 then
	GetLastForcedate = CStr(oRS5.Fields.Item("forcedatetime")) + " - " + CStr( oRS5.Fields.Item("forceinventoryupd"))
else
	GetLastForcedate = "N/A"
End if
oRS5.close
Set oRS5 = Nothing   
END FUNCTION


FUNCTION GetBELastForcedate()
Set oRS5 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT BEforcedatetime FROM dbo.counters "
oRS5.Open strSQL, MM_overseaspr_STRING

If oRS5.EOF or IsNull(oRS5("BEforcedatetime")) then
	GetBELastForcedate = "N/A"
Else
	GetBELastForcedate = CStr(oRS5.Fields.Item("BEforcedatetime"))
End if
oRS5.close
Set oRS5 = Nothing   
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
		PartPhoto = "<a style='width:100%' href='" + ToRootedVirtual(PartNumJPG) + "/parts_images/" + PartNum + ".jpg' target='_new'  >" + PartNum + "</a>"
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
	
	if ( LEN(part) > 0) then
		'UPDATE QOH
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open MM_overseaspr_STRING
		sql="UPDATE dbo.partmst1_distinct SET image_exist = '1' WHERE dbo.partmst1_distinct.field_1 = '" + part + "' ;"
		objConn.Execute sql
		objConn.Close
		Set objConn = Nothing
	end if
			
END SUB

SUB AddToBanner(part)
	
	if Not IsEmpty(part) then

		'UPDATE QOH
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open MM_overseaspr_STRING
		sql="UPDATE dbo.partmst1_distinct SET image_exist = '2' WHERE field_1 = '" + part + "' ;"
		objConn.Execute sql
		objConn.Close
		Set objConn = Nothing
		
		response.redirect("catalog_beta.asp")
		
	end if
			
END SUB

FUNCTION GetUserTax(userid)
Set oRS5 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT citytax,statetax FROM dbo.users where user_auto_id = '" + userid + "' "
oRS5.Open strSQL, MM_overseaspr_STRING

If oRS5.EOF or IsNull(oRS5("citytax")) then
	GetUserTax = "N"
Else
	
	If IsEmpty(oRS5.Fields.Item("citytax")) or trim(oRS5.Fields.Item("citytax")) = "" or IsNull(oRS5.Fields.Item("citytax")) Then
		GetUserTax = "N"
	Else
		GetUserTax = oRS5.Fields.Item("citytax")
	End If	
	
End if
oRS5.close
Set oRS5 = Nothing   
END FUNCTION

FUNCTION GetNombreCatalogo(IdCatalog)
Set oRS5 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT CatalogName FROM dbo.OIC_Catalog where IdCatalog = " + CStr(IdCatalog) 
oRS5.Open strSQL, MM_overseaspr_STRING

If oRS5.EOF then
	GetNombreCatalogo = "N/A"
Else
	
	If IsNull(oRS5.Fields.Item("CatalogName")) Then
		GetNombreCatalogo = "N/A"
	Else
		GetNombreCatalogo = oRS5.Fields.Item("CatalogName")
	End If	
	
End if
oRS5.close
Set oRS5 = Nothing   
END FUNCTION

FUNCTION GetNombreCategoria(CategoryId)
Set oRS5 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT Category FROM dbo.OIC_Category where CategoryId = " + CStr(CategoryId) 
oRS5.Open strSQL, MM_overseaspr_STRING

If oRS5.EOF then
	GetNombreCategoria = "N/A"
Else
	
	If IsNull(oRS5.Fields.Item("Category")) Then
		GetNombreCategoria = "N/A"
	Else
		GetNombreCategoria = oRS5.Fields.Item("Category")
	End If	
	
End if
oRS5.close
Set oRS5 = Nothing   
END FUNCTION

FUNCTION GetNombreSubCategoria(CategoryId, SubCategoryId)

if not IsEmpty(CategoryId) AND Not IsEmpty(SubCategoryId) then

	Set oRS5 = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT Category FROM dbo.OIC_SubCategory where CategoryId = " + CStr(CategoryId) + " AND SubCategoryId = " + CStr(SubCategoryId) 
	oRS5.Open strSQL, MM_overseaspr_STRING
	
	If oRS5.EOF then
		GetNombreSubCategoria = "N/A"
	Else
		
		If IsNull(oRS5.Fields.Item("Category")) Then
			GetNombreSubCategoria = "N/A"
		Else
			GetNombreSubCategoria = oRS5.Fields.Item("Category")
		End If	
		
	End if
	oRS5.close
	Set oRS5 = Nothing   
else
	GetNombreSubCategoria = ""
end if
END FUNCTION


FUNCTION GetCategoriadeNombre(nombrecategoria)
Set oRS5 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT CategoryId FROM dbo.OIC_Category where Category = '" + nombrecategoria + "' ;" 
oRS5.Open strSQL, MM_overseaspr_STRING

If oRS5.EOF then
	GetCategoriadeNombre = 0
Else
	GetCategoriadeNombre = Clng(oRS5.Fields.Item("CategoryId"))
End if
oRS5.close
Set oRS5 = Nothing   
END FUNCTION

FUNCTION GetCountCategoria(categoryid)
Set oRS5 = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT count(*) as counter FROM dbo.OIC_Category where CategoryId = " & CStr(categoryid)  
oRS5.Open strSQL, MM_overseaspr_STRING

If oRS5.EOF then
	GetCountCategoria = 0
Else
	GetCountCategoria = Clng(oRS5.Fields.Item("counter"))
End if
oRS5.close
Set oRS5 = Nothing   
END FUNCTION

FUNCTION availabletags()

	
	Set keywords = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT distinct field_1 FROM dbo.partmst1_distinct WHERE image_exist = '1' Order by field_1 ;"
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
	If Len(vchoices) > 3 Then
		vchoices = mid(vchoices,1,Len(vchoices)-2)
		vchoices = "["""&vchoices &"]"
	Else
		vchoices = "'N/A'"
	End If
	response.write vchoices
	
END FUNCTION	

FUNCTION GetPartNoDefault(catalogid,categoryid,titulo)
If catalogid > 0 AND categoryid > 0 Then
	Set oRS5 = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT partno FROM dbo.OIC_PartsPerCategory WHERE uniqueidcol = (select min(uniqueidcol) from dbo.OIC_PartsPerCategory WHERE titulo = '" + trim(titulo) + "' AND idcatalog = " + CStr(catalogid) + " AND categoryid = " + CStr(categoryid) + ")"
	oRS5.Open strSQL, MM_overseaspr_STRING
	
	If oRS5.EOF then
		GetPartNoDefault = ""
	Else
		
		If IsNull(oRS5.Fields.Item("partno")) Then
			GetPartNoDefault = ""
		Else
			GetPartNoDefault = oRS5.Fields.Item("partno")
		End If	
		
	End if
End If
END FUNCTION	

FUNCTION GetTituloDefault(catalogid,categoryid)
If catalogid > 0 AND categoryid > 0 Then
	Set oRS5 = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT titulo FROM dbo.OIC_PartsPerCategory WHERE uniqueidcol = (select min(uniqueidcol) from dbo.OIC_PartsPerCategory WHERE idcatalog = " + CStr(catalogid) + " AND categoryid = " + CStr(categoryid) + ")"
	oRS5.Open strSQL, MM_overseaspr_STRING
	
	If oRS5.EOF then
		GetTituloDefault = ""
	Else
		
		If IsNull(oRS5.Fields.Item("titulo")) Then
			GetTituloDefault = ""
		Else
			GetTituloDefault = oRS5.Fields.Item("titulo")
		End If	
		
	End if
End If
END FUNCTION

FUNCTION PartInInventory(part)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT count(*) as counter FROM dbo.partmst1_distinct WHERE field_1 = '" + part + "' ;"
	oRS.Open strSQL, MM_overseaspr_STRING
	If Not oRS.EOF Then
	 PartInInventory = CInt(oRS.Fields.Item("counter"))
	Else
	 PartInInventory = 0
	End If  
END FUNCTION	

FUNCTION CheckTime()
		
tlimit = FormatDateTime("9:30", vbShortTime)
tnow = FormatDateTime(Now, vbShortTime)
tlimit2 = FormatDateTime("9:51", vbShortTime)			

if (tnow > tlimit AND tnow < tlimit2) then
	CheckTime = "<p class='btn btn-danger' style='height:50px'>30 min (Same Day)</p>"
else
	CheckTime = ""	
end if
			
END FUNCTION	

%>