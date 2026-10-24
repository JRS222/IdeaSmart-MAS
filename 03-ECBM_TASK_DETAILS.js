var currentWhepFacID;
var currentMachineType;
var currentMachineNo;
var currentMachinePtr;
var currentMachineClassCode;
var currentTaskID = '';
var siteName;

function bodyOnLoad() { 
	currentMachineType=opener.currentMachineType;
	currentMachineNo=opener.currentMachineNo;
	currentMachinePtr=opener.currentMachinePtr;
	currentMachineClassCode = opener.currentMachineClassCode;
	currentRoute=opener.currentRoute;
	currentTaskID=opener.currentTaskID;
	currentWhepFacID = opener.currentWhepFacID;
	siteName = opener.siteName;

	var strheader = '<span style="align:center;"><table width="100%"><tr>' + 
	'<td bgcolor="000099"><font color="cccccc">&nbsp;Subsite Name:&nbsp;</font></td>' + 
	'<td bgcolor="cccccc"><font color="000099">&nbsp;&nbsp;' + siteName + '&nbsp;&nbsp;</font></td>' +
	'<td bgcolor="000099"><font color="cccccc">&nbsp;Equipment Type:&nbsp;</font></td>' + '<td bgcolor="cccccc"><font color="000099">&nbsp;&nbsp;' + currentMachineType + '&nbsp;&nbsp;</font></td>' + 
	'<td bgcolor="000099"><font color="cccccc">&nbsp;Equipment No.:&nbsp;</font></td>' + 
	'<td bgcolor="cccccc"><font color="000099">&nbsp;&nbsp;' + currentMachineNo + '&nbsp;&nbsp;</font></td>' + 
	'<td bgcolor="000099"><font color="cccccc">&nbsp;Class Code.:&nbsp;</font></td>' + 
	'<td bgcolor="cccccc"><font color="000099">&nbsp;&nbsp;' + currentMachineClassCode + '&nbsp;&nbsp;</font></td>' + 
	'</tr></table></span>';
	
	var strStylesheet = '<style type="text/css">'+
	'DIV.danger {'+
		'background-image:url(/'+vReportUserIdMapping+'/warning.gif);'+
		'background-repeat:no-repeat;'+
		'background-position:center top;'+
		'margin: 0 0 2em 1em;'+
		'width: 300px;'+
		'border:#000000 solid thin;'+
		'border-collapse:collapse;'+
		'font-weight:bold;'+
		'padding: 4em 1em 1em 1em;'+
	'}'+
	'DIV.caution {	background-image:url(/'+vReportUserIdMapping+
	'/caution.gif);	background-repeat:no-repeat;	background-position:center top;margin: 0 0 2em 1em;	width: 300px;'+
	'border:#000000 solid thin;	border-collapse:collapse;font-weight:bold;padding: 4em 1em 1em 1em;}'+
	'DIV.note {background-image:url(/'+vReportUserIdMapping+
	'/note.gif);background-repeat:no-repeat;background-position:center top;margin: 0 0 2em 1em;width: 300px;'+
	'border:#000000 solid thin;border-collapse:collapse;font-weight:bold;padding: 3.5em 1em 1em 1em;}'+
	'</style>';

	document.all.div_taskheader.innerHTML = strheader + strStylesheet;
	document.getElementById('div_task').innerHTML=showTask() ;
}

function getCiossSize()
{
	var objHTTP=new ActiveXObject('microsoft.XMLHTTP');	
    var urlStr=getServerPath()+'ecbm.ecbm_machine_lists.getMachineSize?p_machineid='+currentMachinePtr+'&p_site='+currentSite;
    objHTTP.open('GET', urlStr, false);
    objHTTP.send();

	return objHTTP.responseText;	
}

// if(currentMachineType == "CIOSS" )
function getCIOSSTime(tid)
{ 
   			var objHTTP=new ActiveXObject('microsoft.XMLHTTP');	
   				var urlStr=getServerPath()+'ecbm.ecbm_forecast.getTaskEstTime?p_site='+currentSite+'&p_machinetype='+currentMachineType+'&p_classcode='+currentMachineClassCode+'&p_machineid='+currentMachinePtr+'&p_taskid='+tid;
    			objHTTP.open('GET', urlStr, false);
   				objHTTP.send();
						
				if(objHTTP.responseText*1 == -1  || objHTTP.responseText*1 == -3){
						alert("Failed to retrieve the estimated time for task "+ tid + ". Please call the eMARS HelpDesk, thanks.");
						return -1;
				}else{
					if(objHTTP.responseText*1 == -2){	
						return -2;
					}else{
					    minutesPerTask = objHTTP.responseText*1; 
						return  minutesPerTask; 
					}
				}
		
} 

//display the step details
function showTask()
{		 
	var tempStr = null; 
	var tempNum = -1;
	var objAcroList = null;
	var objAcros=new ActiveXObject('microsoft.XMLDOM');
   	objAcros.async=false;
	
	/* get the checklist*/
   objAcros.load("ecbm.ecbm_report_j.getTaskDetails?p_machinetype=" + currentMachineType + "&p_classcode=" + currentMachineClassCode+ "&p_taskid=" + currentTaskID+ "&p_machineid=" +currentMachinePtr+ "&p_site=" + currentSite);	
   if (objAcros.parsed) {
      objAcroList=objAcros.getElementsByTagName("ROW");
	  var childList=objAcroList(0).childNodes;
	
	  tempStr = '<p style="font:arial black;font-size:large;text-align:center;">Details of '+currentMachineType+' PM Task No.'+childList[24].text+'</p>'+
	  		'<table border="1" align="center" width="830">';
   
	  var descStr = new String(childList[2].text);	  
			 
       //remove <p>,</p>
	  while (descStr.search(/<\/?[pP]>/) >= 0){
			descStr = 	descStr.replace(/<\/?[pP]>/,'  ');	 	
	  }
	  while (descStr.search(/<br><br>/) >= 0){
			descStr = 	descStr.replace(/<br><br>/,'<br> ');	 	
	  }
  
	  var noteStr = new String(childList[21].text);  
	  //remove <p>,</p>
	  while (noteStr.search(/<\/?[pP]>/) >= 0){
			noteStr = 	noteStr.replace(/<\/?[pP]>/,', ');	 	
	  }
	  while (noteStr.search(/<br><br>/) >= 0){
			noteStr = 	noteStr.replace(/<br><br>/,'<br> ');	 	
	  }
	  
	  tempStr += '<tr bgcolor="silver"><th  width="150">MMO Item No.</th><td  width="600">'+childList[24].text+'</td></tr>'+
	  			'<tr bgcolor="lightgrey"><th  width="150">Task Type</th><td  width="">'+childList[23].text+'</td></tr>'+
	  			'<tr bgcolor="silver"><th  width="150">Sub-Assembly</th><td  width="">'+childList[0].text+'</td></tr>'+
				'<tr bgcolor="lightgrey"  style="font-size:14px;"><th>Part or Component</th><td>'+childList[1].text+'</td></tr>';
				
	 var minutesPerTask = new Number(childList[3].text).valueOf();
	
	 //calculate the time reuired for APPS
	 if(currentMachineType == "APPS"){    	//<!--- if 2 --->
		var multilierValues = getMultilierValues();
		var multiFactors = new Array("feed","carriercells","bins","apcu","pun","gcpu","vdcsupply","tencells","twentyfivebins","apcupun","apcugate","estop","sortconfig","mab","occ","dailyhours");

		if(childList[11].text == "T"){ //perUnit	<!--- if 3 --->
			var multiFactor = childList[13].text.toLowerCase();

			if(multiFactor == "fixed"){ //multiplevar
				minutesPerTask = minutesPerTask*new Number(childList[12].text).valueOf();							
			}else{  //<!--- Feed ... E-stop  --->
				for(var x=0;x<multiFactors.length;x++)
				{  //find the position of multiplier in database procedure, see procedure ecbm_checklist.getEORData(...) 
					if(multiFactors[x] == multiFactor){
						minutesPerTask = minutesPerTask*new Number(multilierValues[x+22].text).valueOf();

						break; 
					}
				}							
			} 
	 	}  // <!--- if 3 --->
	 } //<!--- if 2 --->
	
	tempNum =  getCIOSSTime(currentTaskID);
	//if((currentMachineType == "CIOSS" || currentMachineType == "DIOSS"  || currentMachineType == "DBCS")  && childList[11].text == "T"){ //perUnit	
		//CIOSS ....,  the estimated time will be multiplied by the equipment size %$&* ....... 
	if(tempNum > 0 ){
		 minutesPerTask =  tempNum; 
	}

	var localT = 'No';
	if(childList[14].text != '0') { localT = 'Yes'; }
	var runT =  childList[7].text;
	var piecesT = childList[8].text;
	var weeklyT = childList[9].text; 
	var freqT = childList[22].text;
	var issued_by_date_interval = childList[25].text;
	
	if(runT == '-1'){ runT = 'n/a'; }
	if(	piecesT == '-1') { piecesT = 'n/a'; }
	if(weeklyT == '-1'){ 
		if(issued_by_date_interval == '-1') {
			weeklyT = 'n/a'; 
		}else{
			weeklyT = 'Issued every ' + issued_by_date_interval + ' days, and the next date will be  ' + childList[26].text;
		}		
	}
	if(freqT == '-1'){	freqT = 'n/a'; }
	
	tempStr += '<tr bgcolor="lightgrey"><th>Est. Time Req`d (min)</th><td>'+ Math.ceil(minutesPerTask)+'</td></tr>'+
				'<tr bgcolor="silver"><th>Min. Skill Level</th><td>'+childList[4].text+'</td></tr>'+
				'<tr bgcolor="lightgrey"><th>Power On/Off</th><td>'+childList[5].text+'</td></tr>'+
				'<tr bgcolor="silver"><th>Unit Name</th><td>'+childList[6].text+'</td></tr>'+
				'<tr bgcolor="lightgrey"><th>Runhours Threshold</th><td>'+runT+'</td></tr>'+
				'<tr bgcolor="silver"><th>Pieces Fed Threshold</th><td>'+piecesT+'</td></tr>'+
				'<tr bgcolor="lightgrey"><th>Weekly Threshold</th><td>'+weeklyT+'</td></tr>'+
				'<tr bgcolor="silver"><th>Tourly,Daily or Monthly</th><td>'+freqT+'</td></tr>'+
				'<tr bgcolor="lightgrey"><th>Days To Live</th><td>'+childList[10].text+'</td></tr>'+
				'<tr bgcolor="silver"><th>Local Task</th><td>'+localT+'</td></tr>'+
				'<tr bgcolor="lightgrey"  style="font-size:14px;"><th>Note</th><td>'+noteStr+'</td></tr>';
		
	 tempStr += '<tr bgcolor="silver"><th>Task Statement and Instruction</th><td   style="font-size:12px;">';
	 tempStr +=  descStr+'</td></tr>';

	  tempStr += '</table>';	  
	}else{
		tempStr = "ERROR in Loading ... task No. " + currentTaskID + " of " + currentMachineType;
	}	

	return tempStr;
}

//get multipliers for APPS 
function getMultilierValues()
{	
	var tempStr = null;
	var objAcroList = null;
	var objAcros=new ActiveXObject('microsoft.XMLDOM');
   	objAcros.async=false;

	//get the checklist
   objAcros.load('ecbm.ecbm_checklist.getEORData?p_machineid='+currentMachinePtr+'&p_machinetype='+currentMachineType+'&p_classcode='+currentMachineClassCode);  
   if (objAcros.parsed) {
      objAcroList=objAcros.getElementsByTagName("ROW");
	  var childList=objAcroList(0).childNodes;  

	  return childList;
   }

   return null;  
}

function btn_print_onClick()
{
		var newWin = null;
		newWin = window.open("","win2","alwaysRaised,status,menubar,height=600,width=800,resizable,scrollbars,top=0,left=0",true);
	    newWin.document.write(showTask());
		newWin.document.close();
	    newWin.focus();
		newWin.print();			
}

var vEnvironment = getEnvVar().toUpperCase(); 
var vReportUserIdMapping = 'emarsprdrpt';
if(vEnvironment == 'DEV') { vReportUserIdMapping = 'emarsdevrpt'; }
else if (vEnvironment == 'TEST') { vReportUserIdMapping = 'emarststrpt'; } // staging
else if (vEnvironment == 'PROD') { vReportUserIdMapping = 'pemarsrpt'; } // 'pemarsrpt'
else if (vEnvironment == 'CAT') { vReportUserIdMapping = 'cemarsrpt'; }
else if (vEnvironment == 'VIEW') { vReportUserIdMapping = 'emarsdevrpt'; }  //'emarsdevrpt'

function getEnvVar()
{
	var objAcroList = null;
	var objAcros=new ActiveXObject('microsoft.XMLDOM');
   	objAcros.async=false;
	
	var sysEnv = null;
	
	// get the system environment variable  
   objAcros.load("ecbm.ecbm_checklist.getEnvironmentVar");
   
   if (objAcros.parsed) {
      objAcroList=objAcros.getElementsByTagName("ROW");
	  
	 for(i=0;i<objAcroList.length;i++)
	  { 
	  	childList=objAcroList(i).childNodes;
		sysEnv = childList[0].text;
	  }	  	  
   }
  
   return sysEnv;
}	



