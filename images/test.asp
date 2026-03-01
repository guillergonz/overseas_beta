<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<SCRIPT Language=JavaScript>
function dwMine2_ButtonClicked(row, objName) {
	if (objName == "b_update_cart") {
		// DELETE THIS COMMENT FOR PRODUCTION, ONLY FOR REFERENCE
		// trn = 10 = (Insertar)
		// trn = 20 = (Eliminar)
		// trn = 30 = (Editar"
		ls_qty = dwMine2.GetItem(row,'quantity');
		ls_part = dwMine2.GetItem(row,'field_1');
		alert("Cantidad \n" + ls_qty + "\nParte\n" + ls_part );
		location.href="update_cart.asp?prt=" + ls_part + "&qty=" + ls_qty + "&trn=10"
	}
}
</SCRIPT>


<%
dim resultado,prt1,prt2,prt3,prt4,prt5,prt6,prt7,prt8,prt9,prt10
dim qty1,qty2,qty3,qty4,qty5,qty6,qty7,qty8,qty9,qty10
dim dw_context, dw_action
dim user_id

user_id = Session("MM_Username")

prt1 = UCase(Request.Form("part1"))

if len(prt1) = 0 then
	prt1 = Request.Form("APart1")
end if

qty1 = UCase(Request.Form("part1_qty"))

'set zorrilla = server.CreateObject("zorrilla.framework")
'resultado = zorrilla.f_retrieve2(user_id,client_code,dw_action,dw_context,prt1,qty1)
'response.Write(resultado)
'set zorrilla = nothing


   '/* Create instance of the COM object */
   set dwMine2 = Server.CreateObject("PowerBuilder.HTMLDataWindow")

   '/* set datawindow object */
   retVal = dwMine2.SetDWObject ("C:\\apps\\zorrilla\\zorrilla\\pb_workspace\\com_framework.pbl", "dw_zorrilla2")

 if retVal = 1 Then 
     '/* set control name, browser */
     retVal = dwMine2.SetHTMLObjectName ("dwMine2")
     browser = Request.ServerVariables ("HTTP_USER_AGENT")
     dwMine2.SetBrowser(browser)

     '/* allow page navigation, and support for methods that cause page reloads */
     selfLink = Request.ServerVariables ("SCRIPT_NAME")
     '//var selfLinkArgs = "name='\"" + dwMine2.name + "\"'";
     retVal = dwMine2.SetSelfLink (selfLink, "APart1=''")

     '/* display 5 rows of data per page */
     'dwMine2.SetPageSize(1)

     '/* set transaction properties */
     connStr = "ConnectString='DSN=zorrilla;UID=dba;PWD=sql',ConnectOption='SQL_DRIVER_CONNECT,SQL_DRIVER_NOPROMPT'"
     dwMine2.setTrans "ODBC", connStr, "", "", "", "", ""

     '/* retrieve the data */
     retVal = dwMine2.retrieveex ("D577")
     if retVal < 0  Then
        Response.Write ("<H1>Retrieve Error: " + retVal + dwMine2.GetLastErrorString()+ "</H1>")
     end if

     '/* Check if page parameters (passed from client) indicate that an action needs to be performed */
     dwMine2_action = Request.Form ("dwMine2_action")
     dwMine2_context = Request.Form ("dwMine2_context")
     if dwMine2_action <> "undefined" then
       '/* perform the action on the server data */
       retVal = dwMine2.SetAction (dwMine2_action, dwMine2_context)
       if retVal < 0  then
  		Response.Write ("<H1>Error on SetAction(): " + retVal + dwMine2.GetLastErrorString() + "</H1>")
       end if
     end if

     '/* generate the HTML DataWindow with data */
     Response.Write dwMine2.Generate()
 
 else 

    Response.Write ("<H1>Error on SetDWObject() = " + retVal + dwMine2.GetLastErrorString() + "</H1>")

 End if
Set dwMine2 = nothing 

%>
<body>

</body>
</html>
