<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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


<style type="text/css">
body {
	margin:0;
	background-color: #FFF;
	background-repeat: repeat;
}
.input focus{
	border: #FC0;
}
body,td,th {
font-family: "Lucida Sans Unicode", "Lucida Grande", sans-serif;
font-size: 12px;
color: #000;
}
</style>
</head>
<body>


<%
Dim dd_category
Dim dd_category_cmd
Dim dd_category_numRows

Set dd_category_cmd = Server.CreateObject ("ADODB.Command")
dd_category_cmd.ActiveConnection = MM_overseaspr_STRING
' if English use column english_desc ggg
'LRO - done
If Session("lang") = "S" Then
	dd_category_cmd.CommandText = "SELECT field_1, field_2, english_desc FROM dbo.familicat ORDER BY field_1 ASC" 
Else
	dd_category_cmd.CommandText = "SELECT field_1, field_2, english_desc FROM dbo.familicat ORDER BY english_desc ASC" 
End if

dd_category_cmd.Prepared = true

Set dd_category = dd_category_cmd.Execute
dd_category_numRows = 0
%>
															      
<script type="text/javascript" charset="utf-8">
function ValidateCart(part,amt) {
	amount = document.getElementById(amt).value;
	AddToCart(part,amount);
};
</script>

<script type="text/javascript" charset="utf-8">
function searchSel() {
  var input=document.getElementById('fks').value.toUpperCase();
  var output=document.getElementById('fcs').options;

  for(var i=0;i<output.length;i++) {
	if(output[i].value.indexOf(input)==0){
	  output[i].selected=true;
	  };
	if(document.forms[0].fks.value==''){
	  output[0].selected=true;
	  };
  }
};
</script>

<script type="application/javascript" language="javascript" >
$(document).ready(function() {
	
	$("#loader").addClass("display_no")
  		
});
</script>




<div class="container">
	
    
	<form id="search_part2" name="search_part2"  action="part_search.asp" method="post" style="background-color:white">
	<br>

    <p>
    <img id="loader" name="loader" src="images/ajax-loader.gif" width="16px" height="16px" alt="loader" >
    </p>

    
  	<div class="row">
    
    	<div class="col-xs-6 col-lg-3">
        <a href="part_search.asp" target="_self"><img src="images/oiclogo2.gif" width="256" height="31" alt="logo"  ></a>
        </div>
        
        
                 
            
        <div class="col-xs-6 col-lg-9">          
            <ul class="nav nav-tabs">
                <li ><a href="part_search.asp" title="<%= Lang("buscar") %>" target="_self"><%= Lang("buscar") %></a></li>
                
                <li class="active"><a id="completar" href="cart.asp"  title="<%= Lang("completar_orden")%>" target="_self" ><%= Lang("completar_orden") %></a></li>
                
                <li><a href="catalog_maint_CAT.asp" title="<%= Lang("catalogo") %> Catalog" target="new"><%= Lang("catalogo") %></a></li>
                
                
                <% If LEN(Session("MM_Multi_Username")) = 0 or Session("MM_Multi_Username") = "CARLE BETANCOURT" then %>
                
                <li><a href="account_statement_iframe.asp" title="<%= Lang("estado_de_cuenta_actual") %>" target="_self"><%= Lang("estado_de_cuenta_actual") %></a></li>
                  	
        		<% End IF %>
        
        
                
				<% if ucase(Session("MM_UserName")) = "Z099" then %>
                <li><a href="uploadDataDaily.asp" title="Import parts" target="_self">Import parts</a></li>
                <li><a href="monitor_beta.asp" title="Monitor" target="_self">Monitor  <span class="badge"><%= GetCurrentSales(Session("MM_UserName")) %> </span></a></li>
                <% End If %>
                
                
            </ul>
        </div>
          
	</div>    

  	<br>
   	<div class="table-responsive pull-right col-xs-12 col-md-3" >    
              
		<% If Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/part_search.asp" or Request.ServerVariables("SCRIPT_NAME") = "/part_search.asp" Then %>
        <br>
        <div class="grayback">
            <%= Lang("preparado") %>
        </div>
        <br>
        <% CartDisplay("N") %>
        <% End if %>
        
        <br>
        
        <% If Request.ServerVariables("SCRIPT_NAME") = "/overseaspr/cart.asp" or Request.ServerVariables("SCRIPT_NAME") = "/cart.asp" or len(Request("o")) > 0 Then %>
        <div class="grayback">
            <%= Lang("ordenes") %>
        </div>
        <br>
        <div id="orders">
            <% OrdersDisplay() %>
        </div>
        <% End if %>
               
	</div>
  
    <div class="panel panel-default col-xs-9">
   		<a href="<%= MM_Logout %>" class="btn btn-default btn-sm pull-right"  ><%= Lang("salir") %></a>
   		<!--<a id="completar" href="cart.asp" class="btn btn-default btn-sm pull-right"><= Lang("completar_orden") %></a>-->
        <div class="panel-heading">
            <%= GetUserName(Session("MM_Username"))%>
            &nbsp;<%= lang("fecha")%>&nbsp;<%=Date()%> 
            <% if LEN(Session("MM_Multi_Username")) > 0 then %>
            <p><%= Session("MM_Multi_Username") %></p>
            <% End If %>               
        </div>
        <div class="panel-body">
            <div class="row">
            	<div class="col-xs-12 col-lg-2">    
            		<%= Lang("entre_5_piezas") %> 
        		</div>
                <div class="col-xs-12 col-lg-2">    
            		<input name="p1" type="text" class="form-control" id="p1" size="9" maxlength="15" > 
        		</div>
                <div class="col-xs-12 col-lg-2">    
            		<input name="p2" type="text" class="form-control" id="p2" size="9" maxlength="15" > 
        		</div>
                <div class="col-xs-12 col-lg-2">    
            		<input name="p3" type="text" class="form-control" id="p3" size="9" maxlength="15" > 
        		</div>
                <div class="col-xs-12 col-lg-2">    
            		<input name="p4" type="text" class="form-control" id="p4" size="9" maxlength="15" > 
        		</div>
                <div class="col-xs-12 col-lg-2">    
            		<input name="p5" type="text" class="form-control" id="p5" size="9" maxlength="15" > 
        		</div>
            </div>	
        </div> 
	</div>
    
    <div class="panel panel-default col-xs-9">      
        <div class="panel-heading">
        	Parameters
        </div>
        <div class="panel-body">
            <div class="row">
            	<div class="col-lg-3"> 
                    
                    <%= Lang("family_keyword") %>
                    <br>
                    <input class="form-control" type="text"  id="fks"  onKeyUp="searchSel()" >
                                   
                </div>
             	<div class="col-lg-3">
					<%= Lang("family_category") %><br>
                    <select class="form-control" name="fcs"  id="fcs" >
                    <option selected value="">&nbsp;</option> 
                    <% While (NOT dd_category.EOF) %>
                    <% If Session("lang") = "S" Then %>
                        <option value="<%=trim(dd_category.Fields.Item("field_1").Value)%>" ><%=trim(dd_category.Fields.Item("field_1").Value)%></option>
                    <% Else %>
                        <option value="<%=trim(dd_category.Fields.Item("english_desc").Value)%>" ><%=trim(dd_category.Fields.Item("english_desc").Value)%></option>
                    <% End if %>
                    <% dd_category.MoveNext()
                    Wend
                    If (dd_category.CursorType > 0) Then
                    dd_category.MoveFirst
                    Else
                    dd_category.Requery
                    End If %>
                    </select>
             	</div>
             	<div class="col-lg-3">
					<%= Lang("modelo") %>
                    <br><input class="form-control" name="model" type="text"  id="fs2" >
             	</div>
                <div class="col-lg-1">
					<%= Lang("Specials")%>
                    <br><input name="special_items" type="checkbox" class="form-control" id="special_items" >
                </div>
    			<div class="col-lg-2">
    				<br><input class="btn btn-primary btn-sm" name="submit" type="submit" id="submit" value="<%=Lang("buscar")%>" >
    			</div>
            </div>
        </div> 			
	</div>
 
 	<div class="row"> 
    	<div class="col-xs-8 pull-left" >
            <iframe src="cart_iframe.asp"  frameborder="0" width="100%" height="4980px"  style="overflow-x:hidden; overflow-y:hidden;background-color:transparent" allowtransparency="yes" >
            </iframe>
        </div>    
    </div>              
                  
      
   
    
	</form>            
  	<script type="text/javascript" src="../bootstrap-3.3.6-dist/js/bootstrap.min.js"></script>
    <!--div container-->
</div>

</body>
</html>
