<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" -->

<!-- Bootstrap -->
<link rel="stylesheet" type="text/css" href="bootstrap-3.3.6-dist/css/bootstrap.min.css">

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script type="text/javascript" charset="utf-8" src="bootstrap-3.3.6-dist/jquery.min.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->

<%
Response.Expires = -1
'--------------------------------------------------------------------------
' THE PURPOSE OF THIS PAGE IS:
'
' 1) ADD ITEM TO OIC_CATALOG_CATEGORY
'--------------------------------------------------------------------------




vcatid 	  = Session("MM_CatalogID")

vpartno	 = Request("partno")
vcategoryid = Session("MM_CatID")
action	  = Request("action")
vuid		= Request("uid")

vsubcatid     = Session("MM_SubCatID")


If action = "DEL" Then
	vtitulo = Session("MM_Titulo")
Else
	if IsNull(request("titulo")) or trim(request("titulo")) = "undefined" then
		vtitulo = Session("MM_Titulo")
	else	
		'  ""(read) and "ADD"
		vtitulo = request("titulo")
	end if
End If

If action = "ADD" Then
	If LEN(request("nota")) > 0 and Not IsEmpty(request("nota")) Then
		vnota = request("nota")
	Else
		vnota = " "
	End If
Else
	vnota = " "
End If



If LEN(vpartno) > 0 AND LEN(vnota) > 0 AND action = "ADD" Then
	
	AddPartToCategory vcatid,vcategoryid,vsubcatid,vpartno,vtitulo,vnota
	
	if vtitulo <> Session("MM_Titulo") Then
		Session("MM_Titulo") = vtitulo
		response.write("<script>location.href=""catmaint.asp""</script>")
	end if
	
ElseIF LEN(vuid) > 0 AND action = "DEL" Then
	
	DelPartFromCategory vuid
	Session("MM_Titulo") = GetTituloDefault(vcatid,vcategoryid)
	' Check if last item was deleted to remove the whole titulo
	Dim vcounter
	vcounter = GetCountCategoria(vcategoryid)
	If vcounter = 0 Then 
		response.write("<script>location.href=""catmaint.asp""</script>")
	End If	
Else
	Session("MM_Titulo") = vtitulo
End If



Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Dim vcatid,vcatalogid

If IsNull(Session("MM_CatID")) or IsEmpty(Session("MM_CatID")) or trim(Session("MM_CatID"))="" or trim(Session("MM_CatalogID"))="" or IsNull(Session("MM_CatalogID")) or IsEmpty(Session("MM_CatalogID")) Then
	response.redirect("index.asp")
else
	vcatid 		= Session("MM_CatID")
	vcatalogid 	= Session("MM_CatalogID")
End If	
If IsNull(Session("MM_PartNo")) or IsEmpty(Session("MM_PartNo")) Then
	response.redirect("index.asp")
End If



Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_overseaspr_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.OIC_PartsPerCategory WHERE titulo = '" + vtitulo + "' AND categoryid = " + CStr(vcatid) + " AND nota > '' " 
Recordset1_cmd.Prepared = true
Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0

%>	


<script type="text/javascript" charset="utf-8">
jq = jQuery.noConflict();

this.imagePreview = function(){	

	jq("#preview").remove();

	/* CONFIG */
	xOffset = 10;
	yOffset = 30;
	// these 2 variable determine popup's distance from the cursor
	// you might want to adjust to get the right result
	/* END CONFIG */
	jq("a.original").hover(function(e){
		this.t = this.title;
		this.title = "";	
		var c = (this.t != "") ? "<br/>" + this.t : "";
		
		jq("#preview").remove();
		jq("#imageplaceholder").show();
		
		jq("#imageplaceholder").append("<p id='preview'><img style='height:150px;width:150%;display:block' src='"+ this.href +"' alt='Image preview' />"+ c +"</p>");								 
		jq("#preview")
			.css("top",(e.pageY - xOffset) + "px")
			.css("left",(e.pageX + yOffset) + "px")
			.fadeIn("fast");						
	},
	function(){
		this.title = this.t;	
		jq("#preview").remove();
		jq("#imageplaceholder").hide();
	});	
	jq("a.original").mousemove(function(e){
		jq("#imageplaceholder").show();
		jq("#preview")
			.css("top",(e.pageY - xOffset) + "px")
			.css("left",(e.pageX + yOffset) + "px");
	});			
};



	
	
	
	imagePreview();		

</script>




    
    <table id="example" class="table table-striped" >
    <thead>
    <tr>
    <th ></th>
    <th ></th>
 
    </tr>
    </thead>
    <tbody>
    <% While  (NOT Recordset1.EOF) %>
    <tr >
      <td  align="left" style="font-size:12px"   >
        &nbsp;<%=mid(Recordset1.Fields.Item("nota").Value,1,7)%>
        <br>&nbsp;<%=mid(trim(Recordset1.Fields.Item("partno").Value),1,16)%>
        </td>
      
      <td  align="center" valign="middle"  >
        
        <div style="margin:0;padding:0; text-wrap:none">
        
           <div style="width:30px;display:inline;margin:0px;padding:0px; text-align:center">
           
           
           <% 
		   ' SIMILARES
		   
			Set oRS_similares = Server.CreateObject("ADODB.Recordset")
			oRS_similares.Open "SELECT max(partid) as maxpart FROM similar WHERE replacement = '" & Recordset1.Fields.Item("partno").Value & "' or REPLACE(replacement,' ','') = '" & trim(Recordset1.Fields.Item("partno").Value) & "' ;", MM_overseaspr_STRING
						
			If Not oRS_similares.EOF AND Not IsNull(oRS_similares("maxpart")) Then
			
				
				maxpart = trim(oRS_similares("maxpart"))
			end if	
            
			oRS_similares.Close()
			Set oRS_similares = Nothing
                
          	%>
                               
                            
           <% If LEN(maxpart) > 0 or PartInInventory(Recordset1.Fields.Item("partno").Value) > 0 Then %>
           
			   <% if LEN(maxpart) > 0 then %>
    				<label style="color:green;font-size:10px"><%= lang("reemplazo")%>
                    <a style="text-align:center" href="/add_to_shopping_cart_fromcatalog.asp?p=<%=(maxpart)%>&a=1" title="Añadir <%=(maxpart)%> por pieza <%= (Recordset1.Fields.Item("partno").Value)%> al carrito de compras ..." >
                    <img src="/images/addtocart.jpg" width="25"  height="25" alt="addtocart"  >
                    </a>
                    </label>
                
               <% else %>
               
                    <a style="text-align:center" href="/add_to_shopping_cart_fromcatalog.asp?p=<%=(Recordset1.Fields.Item("partno").Value)%>&a=1" title="Añadir <%=(Recordset1.Fields.Item("partno").Value)%> al carrito de compras ..." >
                    <img src="/images/addtocart.jpg" width="25"  height="25" alt="addtocart"  >
                    </a>
               
               <% end if %>
            
           	<% Else %>
            
                <label style="width=30px"  >
                n/a
                </label>
            
           <% End If %>
            
           </div>
           
            <%
            Dim PartNumJPG
            PartNumJPG = Server.MapPath("./parts_images/" & (Recordset1.Fields.Item("partno").Value) & ".jpg")
            
			Dim objFSO
            Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
            %>
            
            <div style="width:30px;display:inline;margin:0px;padding:0px; text-align:center">
            
			<% If objFSO.FileExists( PartNumJPG ) Then %>
             
           	
           
            <a style="text-align:center" id="original" class="original" href="/parts_images/<%=trim(Recordset1.Fields.Item("partno").Value)%>.jpg" title="Ver imágen <%=(Recordset1.Fields.Item("partno").Value)%>"  >
            <img src="/images/lupa.jpg" width="25"  height="25" alt="viewimage"  >
            </a>
            
              
		
            
			<% Else %>
            
            
       		<!--<a class="ui-state-disabled" style="color:white" title="Imágen %=(Recordset1.Fields.Item("partno").Value)%> no existe!"  >
            <img src="/images/lupa.jpg" width="25"  height="25"  >
            </a>-->
            
           
            
            
            <% End if %>
			</div>
			
			<%
            Set objFSO = Nothing
            %>


			<% if Session("MM_UserName") = "Z099" then %>
            <div style="float:right;margin-right:10px">
            <a href="#" id="delfromcart" onclick="ajaxdelpart(<%=(Recordset1.Fields.Item("uniqueidcol").Value)%>);" title="Eliminar <%=(Recordset1.Fields.Item("partno").Value)%> ..."  >
              <img src="/images/del.png" width="15"  height="15" alt="delpart"  >
            </a>
            </div>
			<% End If %>
        
                
	  </td>
      
    </tr>
    <% 
      Recordset1.MoveNext()
    Wend
    %>
    </tbody>
    </table>
    
    <a class="btn btn-default small" style="width:100%"  href="cart.asp" title="Go to shopping cart" target="_self">Go to shopping cart</a>
    
   

<%
Recordset1.Close()
Set Recordset1 = Nothing
%>