<% @LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->
<%
Response.Expires = -1
MM_Logout = GetSecureVal(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_UserID")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "index.asp"
	' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Not(Request.QueryString = "") Then
    MM_newQS = "?"
    For Each Item In (Request.QueryString)
      If Not (Item = "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(GetSecureVal(Request.QueryString(Item)))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="../index.asp"
MM_grantAccess=false
If Not (Session("MM_Username") = "") Then
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
 if Session("MM_UserName") <> "Z099" then 
 	MM_authFailedURL="index.asp"
    Response.Redirect(MM_authFailedURL)
 End if
%>
<%
if Len(CStr(Request.Form("fecha"))) > 0 then
	Parm_date = FormatDateTime( CStr(Request.Form("fecha")) )
else
	Parm_date = FormatDateTime(Date(), 0)
end if	
'Response.Write( CStr(Parm_date) + "<br>")
'Parm_date = DateAdd("d",-1,Parm_date) 
'Response.Write( CStr(Parm_date) + "<br>")
%>

<%



Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_overseaspr_STRING
'Recordset1_cmd.CommandText = "SELECT  user_auto_id,user_name,order_id,order_number,order_part,order_client,order_qty,order_part_id,order_status,order_date,item_price,comments,citytax,statetax,order_type FROM dbo.clients_orders,dbo.users WHERE dbo.users.user_auto_id = dbo.clients_orders.order_client AND order_date >= '" & Parm_date & "' ORDER BY order_id desc, order_date ;"
'WHERE order_date >= '" & Parm_date & "' 

Recordset1_cmd.CommandText = "SELECT  user_auto_id,user_name,order_id,order_number,order_part,order_client,order_qty,order_part_id,order_status,order_date,item_price,comments,citytax,statetax,order_type FROM dbo.clients_orders,dbo.users WHERE dbo.users.user_auto_id = dbo.clients_orders.order_client AND order_date >= '" & Parm_date & "' ORDER BY order_id desc;"


Recordset1_cmd.Prepared = true
Set recordset1 = Recordset1_cmd.Execute
%>

<!doctype html>
<head>
<meta charset="utf-8"> 
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" >
<link rel="icon" href="images/favicon.ico" type="image/x-icon" >
<title>Overseas Import Corporation</title>

<!-- Bootstrap -->
<link rel="stylesheet" type="text/css" href="bootstrap-3.3.6-dist/css/bootstrap.min.css">

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script type="text/javascript" charset="utf-8" src="bootstrap-3.3.6-dist/jquery.min.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->



<link href="overseas.css" rel="stylesheet" type="text/css" >

<link rel="stylesheet" href="jquery-ui/jquery-ui-1.8.16-uidarkness.custom/development-bundle/themes/base/jquery.ui.all.css">
<link href="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/css/ui-darkness/jquery-ui-1.10.4.custom.css" rel="stylesheet">
<!--<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-1.10.2.js"></script>-->
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/js/jquery-ui-1.10.4.custom.js"></script>

<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/jquery-1.10.2.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.core.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.widget.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.button.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.menu.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.position.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.tooltip.min.js"></script>
<script src="jquery-ui-1.10.4.custom/jquery-ui-1.10.4.custom/development-bundle/ui/minified/jquery.ui.datepicker.min.js"></script>

<style type="text/css" title="currentStyle">
	@import "DataTables-1.9.4/media/css/demo_page.css";
	@import "DataTables-1.9.4/media/css/demo_table_jui.css";

	body {
	margin:0;
	background-color: #FFF;
	background-repeat: repeat;
	}
    .navbar {
        margin-bottom: 0px;
    }
	.input focus{
		border: #FC0;
	}
	body,td,th {
	font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
	font-size: 14px;
	color: #000;
	}

	#dt_example {
		color: #555555;
		font-size:14px;
		font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
	}
	.ui-widget-content {
		background-color: #555555;
	}
	
	.ui-widget-content  input{
		color: #555555;
	}
	
	/*.ui-widget-content a {
		color: white;
	}*/
	
	button, input, textarea, select {
		color:black;
	}
	
	#example_info {
		color:white;
	}
	
	.dataTables, .example_paginate, .example_first, .example_previous  {
		color:white;
	}
    .printbuttonfloatright {
		float: right;
		margin-top: 0px;
		margin-right: 50px;
		margin-bottom: 0px;
		margin-left: 0px;
	}
	
	#example {
		display:none;
	}
	
/*body,td,th {
	color: #000;
	font-size: 12px;
}*/
</style>

<script type="text/javascript" language="javascript" src="DataTables-1.9.4/media/js/jquery.js"></script>
<script type="text/javascript" language="javascript" src="DataTables-1.9.4/media/js/jquery.dataTables.js"></script>

        




<script type="text/javascript" charset="utf-8">
var jq = jQuery.noConflict();

	$(document).ready(function() {

		//$('#updateinv').button();
//		$('#reprocesar').button();
//		$('#print').button();
//		$('#back').button();
//		
//		$('#usuarios').button();
//		$('#catalog1').button();
//		$('#catalog').button();
		
		jq('#example').dataTable({
			"bDestroy": true,
			"bJQueryUI": true,
			"bDeferRender": false,
			"bAutoWidth": true,
			"aLengthMenu":[[5,10,50,100,-1], [5,10,50,100, "All"]],
			"iDisplayLength":-1,
			"sPaginationType": "full_numbers",
			"bLengthChange":true,
			"bInfo":true,
			"bSort":false
		});	

		$('#fecha').datepicker({inline: true});
		
		$('#example').show();
		
	});
</script>
        
</head>


                    
                    
<body id="dt_example" >


<div class="container">

	<nav class="navbar navbar-inverse navbar-static-top" role="navigation">
        <div class="container">
            <div class="navbar-header">
                <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">
                    <span class="sr-only">Toggle navigation</span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                </button>
                
            </div>
    
            <!-- Collect the nav links, forms, and other content for toggling -->
            <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
                <ul class="nav navbar-nav">
                     <li ><a href="part_search.asp" title="<%= Lang("buscar") %>" target="_self"><%= Lang("buscar") %></a></li>
                
                     <li><a id="completar" href="cart.asp"  title="<%= Lang("completar_orden")%>" target="_self" ><%= Lang("completar_orden") %></a></li>
                     <li><a href="catalog_maint_CAT.asp" title="<%= Lang("catalogo") %> Catalog" target="new"><%= Lang("catalogo") %></a></li>
                    
                    
                    <% If LEN(Session("MM_Multi_Username")) = 0 or Session("MM_Multi_Username") = "CARLE BETANCOURT" then %>
                    
                    <li><a href="account_statement_iframe.asp" title="<%= Lang("estado_de_cuenta_actual") %>" target="_self"><%= Lang("estado_de_cuenta_actual") %></a></li>
                        
                    <% End IF %>
            
            
                    
                    <% if ucase(Session("MM_UserName")) = "Z099" then %>
                    <li><a href="uploadDataDaily.asp" title="Import parts" target="_self">Import parts</a></li>
                    <li class="active"><a href="monitor_beta.asp" title="Monitor" target="_self">Monitor  <span class="badge"><%= GetCurrentSales(Session("MM_UserName")) %> </span></a></li>
                    <li><a href="usuarios.asp" title="Usuarios" target="_self">Usuarios</a></li>
                    <% End If %>
                </ul>
            </div>
        </div>
    </nav>
    

    <a class="navbar-brand" rel="home" href="#" title="Overseas Import Corporation">
        			<img style="max-width:256px; margin-top: -7px;" src="/images/oiclogo2.gif"></a>
                    
		<!--<span class="printbuttonfloatright">
		<script type="text/javascript" language="javascript">
            if (window.print) {
              document.write('<form><input align="center" class="btn btn-default" type=button name="print" id="print" value="IMPRIMIR" onClick="window.print()"></form>');	
            }
			  </script>
		
      	</span>-->
          
      
      <!-- <div class="css_left" style="width:90%"  >-->
        <!--<a class="btn btn-default" id="usuarios" name="usuarios" href="usuarios.asp">Usuarios</a>-->
      	<!--<a id="back" name="back" href="part_search.asp">Catálogo de Internet</a>
	   	<a id="catalog1" name="catalog1" href="catalog_maint.asp">Catálogo de Soportes</a>
        
    </div>-->
    
    
      	<br clear="all">
       
          	
			<h4>Catálogo de Internet - Monitor de Actividad</h4>
            
            <div class="row">            
			
            	<div class="col-xs-4">
                  <form method="post" id="form-date" >
                    <label>Seleccione por fecha:</label><br>
                    <input name="fecha" type="text" class="form-control" id="fecha" onChange="form.submit()" value="<%= Parm_date %>" style="width:140px" >
                    <div id="datepicker"></div>
                  </form>
              	</div>  
                
				
                  <form action="reprocess.asp" method="post" id="form-reprocess" >
    
                    <label class="col-xs-3">No. Orden
                      <input  class="form-control" name="orden" type="text"  id="orden" value="" style="width:120px"  >
                    </label>
                    
                    <label class="col-xs-2"><br>
                        <input  class="btn btn-primary" name="reprocesar" type="submit" id="reprocesar"  value="REPROCESAR" style="width:120px">
                    </label>
                        
                  </form>
				
            </div>    
            
   			<br><br>
    	
      
      
        	                
            <div id="demo" style="overflow:hidden" >
              <table class="table table-striped" id="example"  >
                <thead>
                <tr>
                <th width="148">No Orden</th>
                <th width="170">Nombre</th>
                <th width="170">Fecha</th>
                <th width="159">No Pieza</th>
                <th width="109">Cantidad</th>
                <th width="96">Precio</th>
                <th width="89">Status</th>
                <th width="194">Comentarios</th>
                </tr>
                </thead>
                
                <% 
                Dim vloop
                Dim vorder
                Dim vclass
                
            
                
                vloop = 0
                if Not recordset1.EOF then
                 vorder = recordset1("order_number")
                else
                 vorder = 0	
                end if
                vclass = "gradeA"
                While Not recordset1.Eof
                    vloop = vloop + 1
                    if ( vorder <> recordset1("order_number")) then 
                          vorder = recordset1("order_number") 	
                          if (vclass = "gradeA") then
                            vclass = "gradeC"
                            Response.Write( "<tr class='gradeC'>") 
                          elseif (vclass = "gradeC") then 	
                            vclass = "gradeU"
                            Response.Write( "<tr class='gradeU'>") 
                          elseif (vclass = "gradeU") then 	
                            vclass = "gradeX"
                            Response.Write( "<tr class='gradeX'>") 
                          else
                            vclass = "gradeA"
                            Response.Write( "<tr class='gradeA'>") 
                          end if
                    else
                          if (vclass = "gradeA") then
                            Response.Write( "<tr class='gradeA'>") 
                          elseif (vclass = "gradeC") then 	
                            Response.Write( "<tr class='gradeC'>") 
                          elseif (vclass = "gradeU") then 	
                            Response.Write( "<tr class='gradeU'>") 
                          elseif (vclass = "gradeX") then 	
                            Response.Write( "<tr class='gradeX'>") 
                          end if
                    end if
                %>
                
                <td align="center"><%= recordset1("order_number") %></td>
                <td><%= recordset1("user_auto_id")%><br><%= recordset1("user_name") %><br>&nbsp;&nbsp;<label title="Tax" style="color:#F00">Tax:<%= GetUserTax(recordset1("user_auto_id")) %></label></td>
                <td align="center" nowrap><label style="font-size:12px"><%= recordset1("order_date") %></label></td>
                <td align="center"><%= recordset1("order_part") %></td>
                <td align="center"><%= recordset1("order_qty") %></td>
                <td align="right"><%= FormatCurrency(recordset1("item_price")) %>&nbsp;&nbsp;</td>
                
                <td >
                  
                <% If (trim(recordset1("order_status")) = "P") Then
                  Response.Write("Processed") 
                 else
                  Response.Write("Pending") 
                 end if
                %>
                </td>
                
                <td width="400px">
				<% if recordset1("order_type") = "D" then %>
                <label style="color:red">&nbsp;Próximo Día&nbsp;</label>

				<% elseif recordset1("order_type") = "S" then %>
                <label style="color:red">&nbsp;Mismo Día&nbsp;</label>
                
				<% elseif recordset1("order_type") = "P" then %>
                <label style="color:red">&nbsp;Pasa a recoger&nbsp;</label>
                
				<% End If %>
                <br>
				 <label><%= trim(recordset1("comments")) %></label></td>
                
                </tr>
                
                    
              <%
                recordset1.MoveNext
                wend
              %>
                
              </table>
                
            </div>
        <br>
        <br>
        <label>Overseas Import Corporation TM 2016 All rights reserved</label>
        
</div>

</body>
</html>

<%
recordset1.Close()
Set recordset1 = Nothing
%>
