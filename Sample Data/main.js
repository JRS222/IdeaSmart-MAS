////////////////////////////////////////////////////////////////////////////
//window resize event -- style control
$(window).resize(function(){
	var cur_size = $(window).width();
	var per = "80%";
	if( cur_size < 850 ){
		per = "65%";
	}	
	else{
		per = "75%";	
	}
	$("div#ie7fix").css("width", per);
});                        
///////////////////////////////////////////////////////////////////////////

//get viewerflag
var viewerflag = getviewerflag();


//number validation function
function isNum(passedVal){
	if(passedVal == "") {
		return false;
	}
	for ( i = 0; i < passedVal.length; i++){
		if (i == 0 && passedVal.charAt(i) == "0" && passedVal.length != 1){
			return false;
		}
		if (passedVal.charAt(i) < "0" ){
			return false;
		}
		if (passedVal.charAt(i) > "9" ) {
			return false;
		}
	}
	return true;
}

function isDigit(passedVal){
	if(passedVal == "") {
		return false;
	}
	for ( i = 0; i < passedVal.length; i++){
		if (passedVal.charAt(i) != "0" && passedVal.charAt(i) != "1" && passedVal.charAt(i) != "2" && passedVal.charAt(i) != "3" && 
		passedVal.charAt(i) != "4" && passedVal.charAt(i) != "5" && passedVal.charAt(i) != "6" && passedVal.charAt(i) != "7" && 
		passedVal.charAt(i) != "8" && passedVal.charAt(i) != "9"){
			return false;
		}
	}
	return true;
}

function showalt(divobjid, nsn, oem, allflag){
	var myurl = "ajax/listalt.php?nsn=" + nsn + "&oem=" + oem;
	var divobj = document.getElementById(divobjid); 
	if (allflag == false){
		if (divobj.style.display == "block"){
			divobj.innerHTML = "";
			divobj.style.display = "none";
		}
		else{
			$.ajax({url:myurl, success:function(result){
				divobj.innerHTML = result;
				divobj.style.display = "block";
			}});
		}
	}
	else{
		$.ajax({url:myurl, success:function(result){
			divobj.innerHTML = result;
			divobj.style.display = "block";
		}});
	}
	return true;
}

function gostockroom(divcount,stroem,strnsn){
	var strarea = $("#area"+divcount).val();
	var strsearchby = $("#searchby"+divcount).val();
	var url = "";
	if (strarea == "all"){
		if (strsearchby == "OEM" || strnsn == "" || strnsn == "NSL"){
			url = "http://emars.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_national?p_search_type=OEM&p_search_string=" + stroem + "&p_boh_radio=-1";
		}
		else{
			url = "http://emars.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_national?p_search_type=NSN&p_search_string=" + strnsn + "&p_boh_radio=-1";
		}
		window.open(url);
	}
	else{
		if (strsearchby == "OEM" || strnsn == "" || strnsn == "NSL"){
			url = "http://emars.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_area?p_area_id=" + strarea + "&p_search_type=OEM&p_search_string=" + stroem + "&p_boh_radio=-1";
		}
		else{
			url = "http://emars.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_area?p_area_id=" + strarea + "&p_search_type=NSN&p_search_string=" + strnsn + "&p_boh_radio=-1";
		}
		window.open(url);
	}
}

$(document).ready(function(){   
	var left_btm_height = screen.height * 0.64;
	$("div.left_btm").css("height", left_btm_height);

	//section tree ini call
	$("#book_acro").html(acro);
	ini_tree(); //from treeControl.js
	load_front_page();
	leaf_load_figure_table();
	$("div.right_btm").hide();

	// STRY0013240: If we have a sno and figno but no item no, it might 
	// be an assembly like the OEM 1ENS311597 "Drive Module Cover Assembly". 
	// Give the itemno a dummy value so the sno-related click event will 
	// trigger and load the image. It won't auto-scroll the table to a 
	// specific item number, but that is ok, because the "Drive Module Cover 
	// Assembly" is at the top of the table anyway.

	if (sno !="" && figno !="" && itemno == "") {
		itemno = "-";
	}
	// The itemno is not used anywhere else.

	// STRY0013240: This is the code that was preventing the "show figure" 
	// event from happening if there was no itemno:

	//check if there are parameters sno, figno, and itemno passed through URL parts search result
	if (sno != "" && figno != "" && itemno != ""){    
		$("li[sno='" + sno + "'] div").trigger("click");
	}
	//parts book change click event
	$("#change_book").click(function(){
		//load_front_page();
		$("#figdiv").hide();
		$.colorbox({
			href:"#book_nav_wrapper",
			inline: true,
			innerWidth: 700,
			innerHeight: 200,
			opacity: 0.85,
			transition: "elastic",
			//close: "close",
			onClosed:function(){
				$("#figdiv").show();
			},
			escKey: false
		});
	});
	create_print_event_handler();
	//book onchange event
	$("#booklist").change(function(){
		booknum = $(this).val();
		if (booknum != 209){       
		vol = $("#booklist option:selected").attr("volno");
		acro = $("#booklist option:selected").attr("acro");
		//alert(booknum);
		//alert(vol);
		$("#book_acro").html(acro);
		$.ajax({
			url: "content/build_Tree_sections.php",
			type: "POST",
			data:{booknum:booknum, vol:vol},
			success:function(data){
				$("div.col-md-2.sidebar.left_btm #Ryan_fault").html(data);                    
				$.colorbox.close();
				ini_tree();
				load_front_page();
				book_title();
				leaf_load_figure_table();
				$("div.right_btm").hide();
				create_print_event_handler();  
			}
		});
		}
		else{
			$.colorbox.close();
			var srvr;
			var str = document.domain;
			switch (true) {
					case (str.indexOf("devphp") >= 0):
					srvr = "devhandbooks";
				break;
					case (str.indexOf("phpsit") >= 0):
					srvr = "handbookssit";
				break;
					case (str.indexOf("phpstg") >= 0):
					srvr = "handbooksstg";
				break;
					default:
					srvr = "handbooks";
			}
			window.location.href = "https://"+srvr+".mtsc.usps.gov/apps/HBK/index.php#209/I/1/IP_129672";
		}
	})

	$("a.grace_link1").click(function(){
		var url = $(this).attr("ghref") + "msbookno="+booknum+"&volno="+vol;
		window.open(url);
	});     
	$("a.grace_link2").click(function(){
		var url = $(this).attr("ghref") + booknum+"&"+vol;
		window.open(url);
	});
	
	//add cage code click event
	$("#part_tbody").on("click", "span.show_cage", function(){
		var cage = $(this).attr("cage");
		//LOAD cage information
		$.ajax({
			url: "ajax/get_cage.php", 
			type:"POST",
			data:{booknum:booknum, vol:vol, cage: cage},
			success: function(data){
				//try to hide figdiv instead of load_front_page
				$("#figdiv").hide();
				//load cage info in colorbox
				$("#cage_wrapper").html(data);
				$.colorbox({
					href:"#cage_wrapper",
					inline: true,
					title: "Vendor Information",
					innerWidth: 600,
					innerHeight: 300,
					opacity: 0.85,
					transition: "elastic",
					//close: "close",
					onClosed:function(){
						$("#figdiv").show();
					},
					escKey: false
				}); 
			}
		});

	});

	//emars button click event
	$("#part_tbody").on("click", "img.emarsbtn", function(){
		$(this).next().slideToggle("fast");
	});
	
	//book_title click event
	$("#book_title").click(function(){
		load_front_page();
		$("div.right_btm").hide();
	});
	
	// book_title click event
	function book_title(){
		$("#book_title").click(function(){
			load_front_page();
			$("div.right_btm").hide();
		});
	}
	//load the front page
	function load_front_page()
	{
		//LOAD PARTS HANDBOOK INDEX PAGE            
		$.ajax({
			url: "ajax/msindex.php", 
			type:"POST",
			data:{booknum:booknum, vol:vol},
			success: function(data){    
				$("#figdiv").html(data);   
				$("#section_1_content").hide()     
				$("#ie7fix").show();
			}
		});
	}

	//tree leaf click event
	function leaf_load_figure_table()
	{
		//Figure tree leaf click event
		$("#phbk_tree").on("click", "span.go_fig", function(event){
			event.stopPropagation();
			var _parent = $(this).parent();
			var dwg = _parent.attr("dwg");
			var figno = _parent.attr("figno");
			var dwgid = _parent.attr("dwgid");
			//commented out for testing purpose                
			load_figure(figno);
			load_figure_table(figno,0); 
		});
	}

	//print control
	function create_print_event_handler()
	{   
		$("#phbk_tree").on("click", "img.section_print", function(){
			var url = "./content/printsection.php?"+
				"msbookno=" + booknum + 
				"&volno=" + vol +
				"&secno=" + $(this).closest("li").attr("sno") +
				"&viewerflag=" + viewerflag + "&layout=L11";
			window.open(url);
		}).on("click", "img.fig_print", function(){
			var url = "./content/printfigandtable.php?"+
				"msbookno=" + booknum + 
				"&volno=" + vol +
				"&secno=" + $(this).closest("ul").parent().attr("sno") +
				"&figno=" + $(this).closest("li").attr("figno")+
				"&viewerflag=" + viewerflag + "&layout=L11";
			window.open(url); 
		});
	}   
}); //end document ready

function load_figure(fignum){
	getFigInfo(fignum)
	.then(function(result){
		$("#figdiv").empty();
		var loader = $("<img>").prop('id', "loader").attr('src', 'images/loader.gif');
		$("#figdiv").append(result.title);
		$("#figdiv").append(loader);
		$.when($.get(result.path))
		.done(function(data) {
			var file = JSON.parse(data);
			$(loader).remove();
			$("#figdiv").append(file.figure);
			if(file.extention){
				//Clear controls and styles
				$('#figure_controls').remove();
				$('#figureControlCSS').remove();
				$('#figureControlIcons').remove();	
				switch(file.extention) {
					case "img":
						var control = new ImgFigureControls("#container", "#figimg");
						break;
					case "svg":
						var control = new SVGFigureControls("#figdiv", "#svgAll");
						break;
					default:
					// code block
				}
			}
		});
	});
}

function getFigInfo(fignum){
	return new Promise((resolve, reject) => {
		var proceed = "good";
		$.ajax({
			url: "ajax/get_fig_info.php", 
			type:"POST",
			data:{booknum:booknum, vol:vol, fignum:fignum, dwg:"", viewerflag:viewerflag},
			dataType: "json",
			success: function (info){
				resolve(info);				
			},
			error: function (secXHR, secException) {reject(secXHR, secException)},
		})
	});
}

function load_figure_table(fignum,trindex){
	$.ajax({
		url: "ajax/figure_table.php",
		type:"POST",
		data:{booknum:booknum, vol:vol, fignum:fignum, trindex:trindex},
		success: function(data){
			if(data == ""){
				$("div.right_btm").hide();
			}
			else{
				$("#figno").html(fignum);
				$("#part_tbody").html(data);
				$("div.right_btm").show();
				if (trindex != 0){
					scrollToRow(trindex);
				}
				else{
					var row = document.getElementById("figno");
					row.scrollIntoView(false);
				}
			}
		}
	});
}

//scroll to row inside div
function scrollToRow(index){   
	$( "#ptrow" + index)[0].scrollIntoView();
	$("tbody#part_tbody tr").each(function(){
		if ($(this).attr("id") == "ptrow" + index){
			$(this).css('background-color','#FFF8DC');
		}
		else{
			$(this).css('background-color','#FFFFFF');
		}
	});    
}

//ptable function for dwf callout 
function ptable(pstring){
	//alert(pstring);
	var strIndex;
	var intPos = pstring.indexOf('#');
	var intLen = pstring.length;
	strIndex = pstring.substr(intPos+1,intLen-intPos-1); // returns index number
	strDWG = pstring.substr(10,8);
	scrollToRow(strIndex);
}

function isNumeric(input){
	var number = /^\-{0,1}(?:[0-9]+){0,1}(?:\.[0-9]+){0,1}$/i;
	var regex = RegExp(number);
	return regex.test(input) && input.length>0;
} 
