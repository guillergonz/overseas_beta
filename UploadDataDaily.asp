<!-- InstanceBegin template="/Templates/master_template_overseas_bootstrap.dwt" codeOutsideHTMLIsLocked="false" --><%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/overseaspr.asp" -->
<!--#include file="procedures.asp" --> 
<%

%>
<!DOCTYPE HTML>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Overseas Import Corporation</title>
<!-- Bootstrap -->
<link rel="stylesheet" type="text/css" href="bootstrap-3.3.6-dist/css/bootstrap.min.css">

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script src="bootstrap-3.3.6-dist/jquery.min.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->

<link href="overseas.css" rel="stylesheet" type="text/css" />

<!-- InstanceBeginEditable name="doctitle" -->
<!-- InstanceEndEditable -->


<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
	body {
	margin:0;
	background-color: #FFF;
	background-repeat: repeat;
}
.navbar {
    margin-bottom: 0px;
}
.red {
  color: #d14;
}
</style>

<%
Server.ScriptTimeout = 3600
Dim auto
auto = Request.QueryString("auto")
%>

<script language="javascript" type="text/javascript" charset="utf-8">
function getDateTime() {
    var now     = new Date(); 
    var year    = now.getFullYear();
    var month   = now.getMonth()+1; 
    var day     = now.getDate();
    var hour    = now.getHours();
    var minute  = now.getMinutes();
    var second  = now.getSeconds(); 
    if(month.toString().length == 1) {
        var month = '0'+month;
    }
    if(day.toString().length == 1) {
        var day = '0'+day;
    }   
    if(hour.toString().length == 1) {
        var hour = '0'+hour;
    }
    if(minute.toString().length == 1) {
        var minute = '0'+minute;
    }
    if(second.toString().length == 1) {
        var second = '0'+second;
    }   
	if(hour > 12) {
		hour = hour - 12
		suffix = "PM"
	} else {
		suffix = "AM"
	}
	
	// year+'/'+month+'/'+day+' '+
    var dateTime = hour+':'+minute+':'+second+suffix;   
     return dateTime;
}

function processbatch() { 
	$(":input").disabled = true;  
	$("#all").hide();
	
	$(":input").hide(); 
	
	var src = 'uploadData.asp';
	var finished = "";
	
	var auto = $("#auto").val();
		
	$.ajax({
		type:"POST",
		url:src,
		context: document.body,
		data: "action=ADD&auto=" + auto,
		beforeSend: function() {
			$("#glypinv").addClass("red");
			$("#glypall").removeClass("red");},
		success: function(outputhtml){
			$("#ajaxdivInv").html(outputhtml);
			
			$("#glypinv").html(getDateTime());

			
			src = 'uploadDataUsers.asp'
						
			$.ajax({
				type:"POST",
				url:src,
				context: document.body,
				data: "action=ADD",
				beforeSend: function() {
					$("#glypusers").addClass("red");
					$("#glypinv").removeClass("red");},
				success: function(outputhtml){
					$("#ajaxdivUsers").html(outputhtml);
					
					$("#glypusers").html(getDateTime());

					
					src = 'uploadDataFamily.asp'
											
					$.ajax({
						type:"POST",
						url:src,
						context: document.body,
						data: "action=ADD",
						beforeSend: function() {
							$("#glypfam").addClass("red");
							$("#glypusers").removeClass("red");},
						success: function(outputhtml){
							$("#ajaxdivFam").html(outputhtml);
							
							$("#glypfam").html(getDateTime());

							
							src = 'uploadDataInvoice.asp'
							
							$.ajax({
								type:"POST",
								url:src,
								context: document.body,
								data: "action=ADD",
								beforeSend: function() {
									$("#glypinvoice").addClass("red");
									$("#glypfam").removeClass("red");},
								success: function(outputhtml){
									$("#ajaxdivFact").html(outputhtml);	
									
									$("#glypinvoice").html(getDateTime());

									
									src = 'uploadDataSimilares.asp'
									
									$.ajax({
										type:"POST",
										url:src,
										context: document.body,
										data: "action=ADD",
										beforeSend: function() {
											$("#glypsimilar").addClass("red");
											$("#glypinvoice").removeClass("red");},
										success: function(outputhtml){
											$("#ajaxdivSim").html(outputhtml);	
											
											$("#glypsimilar").html(getDateTime());

											
											src = 'uploadDataSpecials.asp'
											
											
											$.ajax({
												type:"POST",
												url:src,
												context: document.body,
												data: "action=ADD",
												beforeSend: function() {
													$("#glypspecials").addClass("red");
													$("#glypsimilar").removeClass("red");},
												success: function(outputhtml){
													$("#ajaxdivspecials").html(outputhtml);	
													
													$("#glypspecials").html(getDateTime());

													
													src = 'uploadDatalong.asp'
																										
													$.ajax({
														type:"POST",
														url:src,
														context: document.body,
														data: "action=ADD",
														beforeSend: function() {
															$("#glyplong").addClass("red");
															$("#glypspecials").removeClass("red");},
														success: function(outputhtml){
															$("#ajaxdivlong").html(outputhtml);	
															
															$("#glyplong").html(getDateTime());

															
															//CONVERSIONES
															src = 'uploadDataConversions.asp'
															$.ajax({
																type:"POST",
																url:src,
																context: document.body,
																data: "action=ADD",
																beforeSend: function() {
																	$("#glypconv").addClass("red");
																	$("#glyplong").removeClass("red");},
																success: function(outputhtml){
																	$("#ajaxdivconv").html(outputhtml);	
																	$("#glypconv").html(getDateTime());
																	
																	finished = "OK";
																	window.print();
																	window.location.href = "http://www.overseaspr.com";
																	
																	$("#glypconv").removeClass("red");
																	$(":input").show();
																	$(":input").disabled = false;		
																	
																},
																error: function(xhr,textStatus, errorThrown){
																	$('#ajaxdivconv').html(textStatus);
																	
																	$("#glypconv").removeClass("red");
																	$(":input").show();
																	$(":input").disabled = false;
																}               
															});																	
														},
														error: function(xhr,textStatus, errorThrown){
															$('#ajaxdivlong').html(textStatus);
															$(":input").disabled = false;
														}               
													});	
													
												},
												error: function(xhr,textStatus, errorThrown){
													$('#ajaxdivspecials').html(textStatus);
													$(":input").disabled = false;
												}               
											});	

										},
										error: function(xhr,textStatus, errorThrown){
											$('#ajaxdivSim').html(textStatus);
											$(":input").disabled = false;
										}               
									});	
															
								},
								error: function(xhr,textStatus, errorThrown){
									$('#ajaxdivFact').html(textStatus);
									$(":input").disabled = false;
								}               
							});	
	
						},
						error: function(xhr,textStatus, errorThrown){
							$('#ajaxdivFam').html(textStatus);
							$(":input").disabled = false;
						}               
					});	
							
				},
				error: function(xhr,textStatus, errorThrown){
					$('#ajaxdivUsers').html(textStatus);
					$(":input").disabled = false;
				}               
			});	
	
		},
		error: function(xhr,textStatus, errorThrown){
			$('#ajaxdivInv').html(textStatus);
			$(":input").disabled = false;
		}               
	});	
	
	if(finished == "OK") {
		//window.close();
		window.location.href = "http://www.overseaspr.com";	
	}
};

function processinventory() { 
	var src
	src = 'uploadData.asp'
	$(":input").disabled = true;  
	$("#all").hide();
	
	$(":input").hide(); 
	
	$.ajax({
		type:"POST",
		url:src,
		context: document.body,
		data: "action=ADD",
		beforeSend: function() {
			$("#glypinv").addClass("red");
			$("#glypall").removeClass("red");},
		success: function(outputhtml){
			$("#ajaxdivInv").html(outputhtml);
			
			$("#glypinv").html(getDateTime());

			
			//CONVERSIONES
															src = 'uploadDataConversions.asp'
															$.ajax({
																type:"POST",
																url:src,
																context: document.body,
																data: "action=ADD",
																beforeSend: function() {
																	$("#glypconv").addClass("red");
																	$("#glyplong").removeClass("red");},
																success: function(outputhtml){
																	$("#ajaxdivconv").html(outputhtml);	
																
																	$("#glypconv").html(getDateTime());

																	
																	$("#glypconv").removeClass("red");
																	$(":input").show();
																	$(":input").disabled = false;		
																	
																},
																error: function(xhr,textStatus, errorThrown){
																	$('#ajaxdivconv').html(textStatus);
																	
																	$("#glypconv").removeClass("red");
																	$(":input").show();
																	$(":input").disabled = false;
																}               
															});				
															
									
		},
		error: function(xhr,textStatus, errorThrown){
			$('#ajaxdivInv').html(textStatus);
			$(":input").disabled = false;
		}               
	});	
};

function processusers() { 
	var src
	src = 'uploadDataUsers.asp'
	$(":input").disabled = true;  
	$("#all").hide();
	
	$(":input").hide(); 

	$.ajax({
		type:"POST",
		url:src,
		context: document.body,
		data: "action=ADD",
		beforeSend: function() {
			$("#glypusers").addClass("red");
			$("#glypinv").removeClass("red");},
		success: function(outputhtml){
			$("#ajaxdivUsers").html(outputhtml);
			
			$("#glypusers").html(getDateTime());

			
			$(":input").disabled = false;								
		},
		error: function(xhr,textStatus, errorThrown){
			$('#ajaxdivUsers').html(textStatus);
			$(":input").disabled = false;
		}               
	});	
	
	$(":input").show();
	$(":input").disabled = false;		
};

function processfam() { 
	var src
	src = 'uploadDataFamily.asp'
	$(":input").disabled = true;  
	$("#all").hide();
	
	$(":input").hide(); 
	
	$.ajax({
		type:"POST",
		url:src,
		context: document.body,
		data: "action=ADD",
		beforeSend: function() {
			$("#glypfam").addClass("red");
			$("#glypusers").removeClass("red");},
		success: function(outputhtml){
			$("#ajaxdivFam").html(outputhtml);
			
			$("#glypfam").html(getDateTime());

			
			//CONVERSIONES
															src = 'uploadDataConversions.asp'
															$.ajax({
																type:"POST",
																url:src,
																context: document.body,
																data: "action=ADD",
																beforeSend: function() {
																	$("#glypconv").addClass("red");
																	$("#glypfam").removeClass("red");},
																success: function(outputhtml){
																	$("#ajaxdivconv").html(outputhtml);	
																
																	$("#glypconv").html(getDateTime());

																	
																	$("#glypconv").removeClass("red");
																	$(":input").show();
																	$(":input").disabled = false;		
																	
																},
																error: function(xhr,textStatus, errorThrown){
																	$('#ajaxdivconv').html(textStatus);
																	
																	$("#glypconv").removeClass("red");
																	$(":input").show();
																	$(":input").disabled = false;
																}               
															});				
																					
		},
		error: function(xhr,textStatus, errorThrown){
			$('#ajaxdivFam').html(textStatus);
			$(":input").disabled = false;
		}               
	});	
};

function processinvoice() { 
	var src
	src = 'uploadDataInvoice.asp'
	$(":input").disabled = true;  
	$("#all").hide();
	
	$(":input").hide(); 
	
	$.ajax({
		type:"POST",
		url:src,
		context: document.body,
		data: "action=ADD",
		beforeSend: function() {
			$("#glypinvoice").addClass("red");
			$("#glypfam").removeClass("red");},
		success: function(outputhtml){
			$("#ajaxdivFact").html(outputhtml);	
			
			$("#glypinvoice").html(getDateTime());

			$(":input").disabled = false;							
		},
		error: function(xhr,textStatus, errorThrown){
			$('#ajaxdivFact').html(textStatus);
			$(":input").disabled = false;
		}               
	});	

	$(":input").show();
	$(":input").disabled = false;		

};

function processimilares() { 
	var src
	src = 'uploadDataSimilares.asp'
	$(":input").disabled = true;  
	$("#all").hide();
	
	$(":input").hide(); 
	
	$.ajax({
		type:"POST",
		url:src,
		context: document.body,
		data: "action=ADD",
		beforeSend: function() {
			$("#glypsimilar").addClass("red");
			$("#glypinvoice").removeClass("red");},
		success: function(outputhtml){
			$("#ajaxdivSim").html(outputhtml);	
			
			$("#glypsimilar").html(getDateTime());

			
			//CONVERSIONES
															src = 'uploadDataConversions.asp'
															$.ajax({
																type:"POST",
																url:src,
																context: document.body,
																data: "action=ADD",
																beforeSend: function() {
																	$("#glypconv").addClass("red");
																	$("#glyplong").removeClass("red");},
																success: function(outputhtml){
																	$("#ajaxdivconv").html(outputhtml);	
																
																	$("#glypconv").html(getDateTime());

																	
																	$("#glypconv").removeClass("red");
																	$(":input").show();
																	$(":input").disabled = false;		
																	
																},
																error: function(xhr,textStatus, errorThrown){
																	$('#ajaxdivconv').html(textStatus);
																	
																	$("#glypconv").removeClass("red");
																	$(":input").show();
																	$(":input").disabled = false;
																}               
															});				
									
		},
		error: function(xhr,textStatus, errorThrown){
			$('#ajaxdivSim').html(textStatus);
			$(":input").disabled = false;
		}               
	});	
};

function processpecials() { 
	var src
	src = 'uploadDataSpecials.asp'
	$(":input").disabled = true;  
	$("#all").hide();
	
	$(":input").hide(); 
	
	$.ajax({
		type:"POST",
		url:src,
		context: document.body,
		data: "action=ADD",
		beforeSend: function() {
			$("#glypspecials").addClass("red");
			$("#glypsimilar").removeClass("red");},
		success: function(outputhtml){
			$("#ajaxdivspecials").html(outputhtml);	
			
			$("#glypspecials").html(getDateTime());

			
			$(":input").disabled = false;							
		},
		error: function(xhr,textStatus, errorThrown){
			$('#ajaxdivspecials').html(textStatus);
			$(":input").disabled = false;
		}               
	});	
	
	$(":input").show();
	$(":input").disabled = false;		

};

function processlong() { 
	var src
	src = 'uploadDatalong.asp'
	$(":input").disabled = true;  
	$("#all").hide();
	
	$(":input").hide(); 
	
	$.ajax({
		type:"POST",
		url:src,
		context: document.body,
		data: "action=ADD",
		beforeSend: function() {
			$("#glyplong").addClass("red");
			$("#glypspecials").removeClass("red");},
		success: function(outputhtml){
			$("#ajaxdivlong").html(outputhtml);	
			
			$("#glyplong").html(getDateTime());

			
			//CONVERSIONES
															src = 'uploadDataConversions.asp'
															$.ajax({
																type:"POST",
																url:src,
																context: document.body,
																data: "action=ADD",
																beforeSend: function() {
																	$("#glypconv").addClass("red");
																	$("#glyplong").removeClass("red");},
																success: function(outputhtml){
																	$("#ajaxdivconv").html(outputhtml);	
																
																	$("#glypconv").html(getDateTime());

																	
																	$("#glypconv").removeClass("red");
																	$(":input").show();
																	$(":input").disabled = false;		
																	
																},
																error: function(xhr,textStatus, errorThrown){
																	$('#ajaxdivconv').html(textStatus);
																	
																	$("#glypconv").removeClass("red");
																	$(":input").show();
																	$(":input").disabled = false;
																}               
															});				
																					
		},
		error: function(xhr,textStatus, errorThrown){
			$('#ajaxdivlong').html(textStatus);
			$(":input").disabled = false;
		}               
	});	
};

$(document).ready(function() {
	
	var auto = $("#auto").val();
	if (auto == "true") {
		processbatch();
	}
	
	$("#all").click(function(ev) {
		processbatch();
	});
	
	$("#binv").click(function(ev) {
		processinventory();
	});

	$("#buser").click(function(ev) {
		processusers();
	});

	$("#bfam").click(function(ev) {
		processfam();
	});

	$("#binvoice").click(function(ev) {
		processinvoice();
	});

	$("#bsim").click(function(ev) {
		processimilares();
	});
	
	$("#bspec").click(function(ev) {
		processpecials();
	});
	
	
	$("#blong").click(function(ev) {
		processlong();
	});
	
	
	// lock
	$("input").prop('disabled', true);

	
	
});		
</script>
<!-- InstanceEndEditable -->

<style type="text/css">
body,td,th {
	font-size: 14px;
}
body {
	margin:0px;
}
body,tr,td,th {
	font-family: "Segoe UI", "Segoe UI Light", "Segoe UI Semibold", "Lucida Console";
	font-size: 14px;
}
</style>
</head>

<body>
<div class="container">


<!-- InstanceBeginEditable name="main_page" -->

<nav class="navbar navbar-inverse navbar-static-top" role="navigation">
        <div id="container" class="container">
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
                    <li class="active"><a href="uploadDataDaily.asp" title="Import parts" target="_self">Import parts</a></li>
                    <li><a href="monitor_beta.asp" title="Monitor" target="_self">Monitor  <span class="badge"><%= GetCurrentSales(Session("MM_UserName")) %> </span></a></li>
                    <li><a href="usuarios.asp" title="Usuarios" target="_self">Usuarios</a></li>
                    <% End If %>
                </ul>
            </div>
        </div>
    </nav>
    

    <a class="navbar-brand" rel="home" href="#" title="Overseas Import Corporation">
        			<img style="max-width:256px; margin-top: -7px;" src="/images/oiclogo2.gif"></a>
                    
	<br>
    <!--<p style="color:red">locked!</p>-->
   
    
<div class="row">	
	<div class="col-xs-12">
		<h4>Rundate&nbsp;<%= response.write(now()) %></h4>
    </div>
</div>
<div class="row">	
	<div class="col-xs-12">
        <label for="todo" class="col-xs-4 col-lg-2">Importar archivos</label>
        
        <button id="all" type="button" class="btn btn-default btn-xs col-xs-4 col-lg-2" > <span class="glyphicon glyphicon-play-circle"></span> Importar </button>
          
        <div class="col-xs-1 col-lg-2">
        <!--<span id="glypall" class="glyphicon glyphicon-flash"></span>-->
        </div>
     </div>       
     <div class="col-xs-10 col-xs-offset-2">
        <div id="ajaxdivAll"></div>
	</div>
</div>


                    
<div class="row">	
	<div class="col-xs-12">
        <label for="inventario" class="col-xs-4 col-lg-2">Inventario</label>
        
        <button id="binv" type="button" class="btn btn-default btn-xs col-xs-4 col-lg-2"> <span class="glyphicon glyphicon-tag"></span> Inventario</button>
        
        <div class="col-xs-1 col-lg-2">
        	<span id="glypinv" class="glyphicon glyphicon-flash"></span>
        </div>
    </div>
    <div class="col-xs-10 col-xs-offset-2">    
        <div id="ajaxdivInv"></div>
	</div>
</div>

<div class="row">	
	<div class="col-xs-12">
        <label for="inventario" class="col-xs-4 col-lg-2">Clientes</label>
        
        <button id="buser" type="button" class="btn btn-default btn-xs col-xs-4 col-lg-2"> <span class="glyphicon glyphicon-user"></span> Usuarios</button>
        
        <div class="col-xs-1 col-lg-2">
        <span id="glypusers" class="glyphicon glyphicon-flash"></span>
        </div>
    </div>    
    <div class="col-xs-10 col-xs-offset-2">    
        <div id="ajaxdivUsers"></div>
	</div>
</div>

<div class="row">
	<div class="col-xs-12">	
        <label for="familia" class="col-xs-4 col-lg-2">Familia</label>
        
        <button id="bfam" type="button" class="btn btn-default btn-xs col-xs-4 col-lg-2"> <span class="glyphicon glyphicon-sort-by-alphabet"></span> Familia </button>
        
        <div class="col-xs-1 col-lg-2">
        <span id="glypfam" class="glyphicon glyphicon-flash"></span>
        </div>
    </div>
    <div class="col-xs-10 col-xs-offset-2">    
        <div id="ajaxdivFam"></div>
    </div> 
</div>

<div class="row">
	<div class="col-xs-12">	
        <label for="facturas" class="col-xs-4 col-lg-2">Facturas</label>
        
        <button id="binvoice" type="button" class="btn btn-default btn-xs col-xs-4 col-lg-2"> <span class="glyphicon glyphicon-list-alt"></span> Facturas </button>
        
        <div class="col-xs-1 col-lg-2">
        <span id="glypinvoice" class="glyphicon glyphicon-flash"></span>
        </div>
    </div>
        
    <div class="col-xs-10 col-xs-offset-2">
        <div id="ajaxdivFact"></div>
    </div> 
</div>

<div class="row">
	<div class="col-xs-12">	
        <label for="similares" class="col-xs-4 col-lg-2">Similares</label>
        
        <button id="bsim" type="button" class="btn btn-default btn-xs col-xs-4 col-lg-2"> <span class="glyphicon glyphicon-equalizer"></span> Similares </button>
        
        <div class="col-xs-1 col-lg-1">
        	<span id="glypsimilar" class="glyphicon glyphicon-flash"></span>
        </div>
    </div>
    <div class="col-xs-10 col-xs-offset-2">	    
        <div id="ajaxdivSim"></div>
    </div> 
</div>

<div class="row">
	<div class="col-xs-12">	
        <label for="specials" class="col-xs-4 col-lg-2">Especiales</label>
        
        <button id="bspec" type="button" class="btn btn-default btn-xs col-xs-4 col-lg-2"> <span class="glyphicon glyphicon-usd"></span> Especiales </button>
        
        <div class="col-xs-1 col-lg-1">
        	<span id="glypspecials" class="glyphicon glyphicon-flash"></span>
        </div>
    </div>
        
    <div class="col-xs-10 col-xs-offset-2">
        <div id="ajaxdivspecials"></div>
    </div> 
</div>

<div class="row">
	<div class="col-xs-12">	
        <label for="long" class="col-xs-4 col-lg-2">Long</label>
        
        <button id="blong" type="button" class="btn btn-default btn-xs col-xs-4 col-lg-2"> <span class="glyphicon glyphicon-barcode"></span> Long </button>
        
        <div class="col-xs-1 col-lg-1">
        	<span id="glyplong" class="glyphicon glyphicon-flash"></span>
        </div>
    </div>    
    <div class="col-xs-10 col-xs-offset-2">
        <div id="ajaxdivlong"></div>
    </div> 
</div>

<div class="row">
	<div class="col-xs-12">	
        <label for="long" class="col-xs-4 col-lg-2">Conversiones</label>
        
        <button  id="bconv" type="button" class="btn btn-default btn-xs col-xs-4 col-lg-2">&nbsp;</button>
        
        <div class="col-xs-1 col-lg-1">
        	<span id="glypconv" class="glyphicon glyphicon-flash"></span>
        </div>
        
    </div>    
    <div class="col-xs-10 col-xs-offset-2">
        <div id="ajaxdivconv"></div>
    </div> 
</div>

<input id="auto" type="hidden" value="<%= auto %>">

<!-- InstanceEndEditable -->


<script src="bootstrap-3.3.6-dist/js/bootstrap.min.js"></script>
</div>
</body>
<!-- InstanceEnd --></html>
