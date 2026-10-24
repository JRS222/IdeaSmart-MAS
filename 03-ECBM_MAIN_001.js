//jesse testing in 3/15/2006  
var gWOHtml;
var gWOHeight = 64;
var findorder='';
var findnsn='';
var findwono='';
var findconsn='';
var findWin='';
var findhist='';
var findlocal='';
var findWinlocalauto;
var l_stock;
var l_for_stock_local;
var otype;
var reg;
var l_main = 'M';
var gCurrentPtr;
var gOrderType;
var w_report;
var ecbm_view_only='N';
var objVc=new ActiveXObject("Microsoft.XMLDOM");

var currentWhepFacID = -9999;
var currentReturnVal;

var l_facility_passed = false;

//Jesse
var currentMachineType = '-1'; //default value in the original list is -1 (blank or please select ...)
var currentMachineNo = '-1';   //default value in the original list is -1 (blank or please select ...)
var currentMachinePtr = '-1';
var currentMachineClassCode = "";
var currentTaskID = 0;
var currentECBMSubsite;
var mmo_no = '  '; //MMO NOs that detail the PM tasks
//jesse
var userLevel = 'MU'; 
/*
	MU 	MARS USER
	HM 	HQ MANAGER
	LA 	LOCAL ADMINISTRATOR
	AD 	HQs ADMINISTRATOR
	AM 	AREA MANAGER
*/

var currentUserRightECBM = "N"; //N=None,R=Read,W=Write,S=Super(visor). See em_user_t
var siteName = ""; 
var currentUserFirstName = "";
var currentUserLastName = ""; 
var currentUserCrew = " ";  

//menu str for pm task
var menuequipmentprofileStr = '<br><DIV><SPAN STYLE="width:630;align:center;font: 10pt arial black;text-align:center;top:5px;left:5px; border-style: outset;border-color:grey;border-width:thick;color: #ff0000; !important; padding:paddingThickness{2,40}">please specify an equipment No. above before going any further!</SPAN></DIV><br>'+
'<fieldset style="width:630">'+
	'<div align="center">'+
	'<table cellpadding="5" cellspacing="0" width="630" height="419" border="0" id="table1">'+
		<!-- MSTableType="layout" -->
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td height="29">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3">'+
			'<img border="0" src="/emars-img/ecbmLogo.JPG" width="200" height="50" align="center" hspace="3"></td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
		'<td height="55">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td height="29">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("UpdateAPPSConfig") style="width: 270; height:40; color: blue; font-weight:bold" value="APPS Configuration" /></td>'+
			'<td>&nbsp;</td>'+ 
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("UpdatePSMConfig") style="width: 270; height:40; color: blue; font-weight:bold" value="PSM Configuration" /></td>'+		
			'<td height="57">&nbsp;</td>'+
		'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("UpdateCIOSSConfig") style="width: 270; height:40; color: blue; font-weight:bold" value="CIOSS/DIOSS/DBCS Equipment Size" /></td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("UpdateSPBSConfig") style="width: 270; height:40; color: blue; font-weight:bold" value="SPBS/APBS Configuration" /></td>'+
			'<td height="57">&nbsp;</td>'+
		'</tr>'+ 
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("UpdateEPPSConfig") style="width: 270; height:40; color: blue; font-weight:bold" value="EPPS CONFIGURATION" /></td>'+
			'<td>&nbsp;</td>'+
			//20250707 '<td valign="top" colspan="3"><input type="button" onClick=doMenu("UpdateMaintWin") style="width: 270; height:40; color: blue; font-weight:bold" value="Maintenance Window" /></td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("UpdateHTPSConfig") style="width: 270; height:40; color: blue; font-weight:bold" value="HTPS Configuration" /></td>'+
			
			'<td height="57">&nbsp;</td>'+
		'</tr>'+  
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("UpdateRUNNINGDAY") style="width: 270; height:40; color: blue; font-weight:bold" value="Non-PM Day" /></td>'+
			'<td>&nbsp;</td>'+
			//20250707 '<td valign="top" colspan="3"><input type="button" onClick=doMenu("UpdateMaintWin") style="width: 270; height:40; color: blue; font-weight:bold" value="Maintenance Window" /></td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("UpdateHOPSConfig") style="width: 270; height:40; color: blue; font-weight:bold" value="HOPS Configuration" /></td>'+
			
			'<td height="57">&nbsp;</td>'+
		'</tr>'+  
		/*'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="5">'+
			'<input type="button" onClick=doMenu("UpdateRUNNINGDAY") style="width: 319; height:40; color: blue; font-weight:bold" value="Non-PM Day" /></td>'+
			'<td>&nbsp;</td>'+
			'<td height="52">&nbsp;</td>'+
		'</tr>'+ */
		'<tr>'+
			'<td width="27">&nbsp;</td>'+
			'<td width="127">&nbsp;</td>'+
			'<td width="50">&nbsp;</td>'+
			'<td width="88">&nbsp;</td>'+
			'<td width="4">&nbsp;</td>'+
			'<td width="97">&nbsp;</td>'+
			'<td width="46">&nbsp;</td>'+
			'<td width="115">&nbsp;</td>'+
			'<td height="62" width="18">&nbsp;</td>'+
		'</tr>'+
	'</table>'+
	'</div>'+
	'</fieldset>';

//menu str for pm task
var menupmtaskStr = '<br><DIV><SPAN STYLE="width:630;align:center;font: 10pt arial black;text-align:center;top:5px;left:5px; border-style: outset;border-color:grey;border-width:thick;color: #ff0000; !important; padding:paddingThickness{2,40}">please specify an equipment No. above before going any further!</SPAN></DIV><br>'+
'<fieldset style="width:630">'+
	'<div align="center">'+
	'<table cellpadding="5" cellspacing="0" width="630" height="419" border="0" id="table1">'+
		<!-- MSTableType="layout" -->
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td height="29">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3">'+
			'<img border="0" src="/emars-img/ecbmLogo.JPG" width="200" height="50" align="center" hspace="3"></td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td height="55">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td height="29">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="5">'+
			'<input type="button" onClick=doMenu("SuspendTask") style="width: 319; height:40; color: blue; font-weight:bold" value="Activate/Suspend Local Task" /></td>'+
			'<td>&nbsp;</td>'+
			'<td height="52">&nbsp;</td>'+
		'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("CreateLocalTask") style="width: 270; height:40; color: blue; font-weight:bold" value="Create Local Task" /></td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("EditLocalTask") style="width: 270; height:40; color: blue; font-weight:bold" value="Edit Local Task" /></td>'+
			'<td height="57">&nbsp;</td>'+
		'</tr>'+
		'<tr>'+
			'<td width="27">&nbsp;</td>'+
			'<td width="127">&nbsp;</td>'+
			'<td width="50">&nbsp;</td>'+
			'<td width="88">&nbsp;</td>'+
			'<td width="4">&nbsp;</td>'+
			'<td width="97">&nbsp;</td>'+
			'<td width="46">&nbsp;</td>'+
			'<td width="115">&nbsp;</td>'+
			'<td height="62" width="18">&nbsp;</td>'+
		'</tr>'+
	'</table>'+
	'</div>'+
	'</fieldset>';	

//menu str for management
var menumanageStr = '<br><DIV><SPAN STYLE="width:630;align:center;font: 10pt arial black;text-align:center;top:5px;left:5px; border-style: outset;border-color:grey;border-width:thick;color: #ff0000; !important; padding:paddingThickness{2,40}">please specify an equipment No. above before going any further!</SPAN></DIV><br>'+
'<fieldset style="width:630">'+
	'<div align="center">'+
	'<table cellpadding="5" cellspacing="0" width="630" height="419" border="0" id="table1">'+
		<!-- MSTableType="layout" -->
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td height="29">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3">'+
			'<img border="0" src="/emars-img/ecbmLogo.JPG" width="200" height="50" align="center" hspace="3"></td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td height="55">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td height="29">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="5">'+
			'<input type="button" onClick=doMenu("AnnualEstimate") style="width: 319; height:40; color: blue; font-weight:bold" value="APPS Annual Workload Estimate" /></td>'+
			'<td>&nbsp;</td>'+
			'<td height="52">&nbsp;</td>'+
		'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("DailyEstimate") style="width: 270; height:40; color: blue; font-weight:bold" value="APPS Daily-Task Workload" /></td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("AllChecklists") style="width: 270; height:40; color: blue; font-weight:bold" value="ALL Checklists" /></td>'+
			'<td height="57">&nbsp;</td>'+
		'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("CompletionReports") style="width: 270; height:40; color: blue; font-weight:bold" value="Completion Reports" /></td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("AllLocalTasks") style="width: 270; height:40; color: blue; font-weight:bold" value="ALL Local Tasks" /></td>'+
			'<td height="57">&nbsp;</td>'+
		'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("AllSeniorTasksDue") style="width: 270; height:40; color: blue; font-weight:bold" value="All Tasks Due" /></td>'+
			'<td>&nbsp;</td>'+
			
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("ChecklistHours") style="width: 270; height:40; color: blue; font-weight:bold" value="Checklist Estimated Hours" /></td>'+
			//'<td style="vertical-align:middle;" colspan="5">'+
			//<input type="button" onClick=doMenu("ChecklistHours") style="width: 319; height:40; color: blue; font-weight:bold" value="All Tasks Due" /></td>'+
			'<td>&nbsp;</td>'+
			'<td height="57">&nbsp;</td>'+
		'</tr>'+ 

		'<tr>'+

			'<td width="27">&nbsp;</td>'+

			'<td width="127">&nbsp;</td>'+

			'<td width="50">&nbsp;</td>'+

			'<td width="88">&nbsp;</td>'+

			'<td width="4">&nbsp;</td>'+

			'<td width="97">&nbsp;</td>'+

			'<td width="46">&nbsp;</td>'+

			'<td width="115">&nbsp;</td>'+

			'<td height="62" width="18">&nbsp;</td>'+

		'</tr>'+

	'</table>'+

	'</div>'+

	'</fieldset>';


//menu str for reports
var menureportStr = '<br><DIV><SPAN STYLE="width:630;align:center;font: 10pt arial black;text-align:center;top:5px;left:5px; border-style: outset;border-color:grey;border-width:thick;color: #ff0000; !important; padding:paddingThickness{2,40}">please specify an equipment No. above before going any further!</SPAN></DIV><br>'+
'<fieldset style="width:630">'+

	'<div align="center">'+

	'<table cellpadding="5" cellspacing="0" width="630" height="419" border="0" id="table1">'+

		<!-- MSTableType="layout" -->

		'<tr>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td height="29">&nbsp;</td>'+

			'</tr>'+

		'<tr>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td valign="top" colspan="3">'+

			'<img border="0" src="/emars-img/ecbmLogo.JPG" width="200" height="50" align="center" hspace="3"></td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td height="55">&nbsp;</td>'+

			'</tr>'+

		'<tr>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td height="29">&nbsp;</td>'+

			'</tr>'+

		'<tr>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td valign="top" colspan="5">'+

			'<input type="button" onClick=doMenu("PMHistory") style="width: 319; height:40; color: blue; font-weight:bold" value="PM History" /></td>'+

			'<td>&nbsp;</td>'+

			'<td height="52">&nbsp;</td>'+

		'</tr>'+

		'<tr>'+

			'<td>&nbsp;</td>'+

			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("EORData") style="width: 270; height:40; color: blue; font-weight:bold" value="EOR Data" /></td>'+

			'<td>&nbsp;</td>'+

			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("TaskThreshold") style="width: 270; height:40; color: blue; font-weight:bold" value="Tasks & Thresholds" /></td>'+

			'<td height="57">&nbsp;</td>'+

		'</tr>'+

		'<tr>'+

			'<td>&nbsp;</td>'+

			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("CompletinRate") style="width: 270; height:40; color: blue; font-weight:bold" value="Completion Rate" /></td>'+

			'<td>&nbsp;</td>'+

			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("BypassedTask") style="width: 270; height:40; color: blue; font-weight:bold" value="Bypassed Tasks" /></td>'+

			'<td height="52">&nbsp;</td>'+

		'</tr>'+
		
		'<tr>'+

			'<td>&nbsp;</td>'+

			'<td>&nbsp;</td>'+

			'<td valign="top" colspan="5">'+

			'<input type="button" onClick=doMenu("TaskCompleted") style="width: 319; height:40; color: blue; font-weight:bold" value="Last Time Tasks Completed" /></td>'+

			'<td>&nbsp;</td>'+

			'<td height="52">&nbsp;</td>'+

		'</tr>'+

		'<tr>'+

			'<td width="27">&nbsp;</td>'+

			'<td width="127">&nbsp;</td>'+

			'<td width="50">&nbsp;</td>'+

			'<td width="88">&nbsp;</td>'+

			'<td width="4">&nbsp;</td>'+

			'<td width="97">&nbsp;</td>'+

			'<td width="46">&nbsp;</td>'+

			'<td width="115">&nbsp;</td>'+

			'<td height="62" width="18">&nbsp;</td>'+

		'</tr>'+

	'</table>'+

	'</div>'+

	'</fieldset>';

	

function bodyOnLoad() { 
  document.all.divTip.innerHTML="Loading ECBM Main Menu...One Moment Please...000";
//hirad
  var ll = launchModalModule('ECBM_FACILITY_LOGIN',330,360);
  var xmlObj = getUserInfo();
	if(xmlObj.length != 0)	{ 			
	  var childList=xmlObj(0).childNodes;
	  currentUserFirstName = childList(0).text;
		currentUserLastName = childList(1).text;
		userLevel = childList(2).text.toUpperCase();
		currentUserRightECBM = childList(4).text.toUpperCase();
	}	   

  if (currentReturnVal== 'NONE') {
    self.close();
  }
	//display eCBM menu
  ECBMMenu(); 

  document.all.div_menuheader.innerHTML = '*** Main Menu ***';
  document.all.divTip.innerHTML='<br><SPAN STYLE="width:800;font:bold;top: 5px; left: 5px;">Subsite Name:<input type="text" value="'+siteName+'" readonly="readonly">'+
 		 '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Equipment Type:<select name="ecbm_machinetype_list"  onChange="machineType_onChange();"><option value="x1">x1</option></select>'+
		 '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Equipment No:<select name="ecbm_machineno_list" onChange="machineNo_onChange();"><option value=""></option></select></span>'; 

	//Jesse 1/19/2007, 5/8/2008 remove it 		
	//document.all.divTip.innerHTML+= '<br><font size="-5" color="navy">Notice: Currently, eCBM is only available for APPS. If you do not have APPS, the Equipment Type list above could be empty.</font>'; 

	loadList('ecbm_machinetype_list', getServerPath()+"ecbm.ecbm_machine_lists.getMachineTypeList?p_site="+currentSite+"&p_subsite="+currentECBMSubsite+"&p_facid="+currentWhepFacID,'Y');
	
	//close the eCBM automatically after 8 hours (28,800,000 milliseconds )
	setTimeout('self.close();', 8*60*60*1000);
}

function machineType_onChange()
{
	
	if (document.all.ecbm_machinetype_list.selectedIndex == 0 ){
		document.all.ecbm_machineno_list.length = 0;
		alert("Please select equipment type!");		
	}else{
		currentMachineType = document.all.ecbm_machinetype_list.value;
		loadList('ecbm_machineno_list', getServerPath()+"ecbm.ecbm_machine_lists.getMachineNoList?p_site="+currentSite+"&p_subsite="+currentECBMSubsite+"&p_equiptype="+currentMachineType, 'Y');
	}
}

function machineNo_onChange()
{
	if (document.all.ecbm_machineno_list.selectedIndex == 0 ){
		alert("Please select equipment no.!");
	}else{
		currentMachinePtr = document.all.ecbm_machineno_list.value;
		var myindex = document.all.ecbm_machineno_list.selectedIndex;
		currentMachineNo = document.all.ecbm_machineno_list.options[myindex].text;

	mmo_no = ' ';
    var childList = null;
  	var objAcroList = null;
	var objAcros=new ActiveXObject('microsoft.XMLDOM');
    objAcros.async=false;

	
	// alert('site ID=' + currentSite);
	 //alert('equip ptr = ' +  currentMachinePtr);
	
	  /* get class-code view_only_ecbm for current machine - equipment_ptr */
    objAcros.load("ecbm.ecbm_machine_lists.getEquipClassView?p_site_id="+currentSite+"&p_equip_ptr="+currentMachinePtr);
    if(objAcros.parsed){
      currentMachineClassCode = objAcros.getElementsByTagName("CLASS")[0].childNodes[0].nodeValue; 
      ecbm_view_only = objAcros.getElementsByTagName("VIEW")[0].childNodes[0].nodeValue;
	  
	  //8/6/2010 get all mmo numbers 
	  objAcroList=objAcros.getElementsByTagName("ROW");
	 for(i=0;i<objAcroList.length;i++)
	  { 
	  	childList=objAcroList(i).childNodes;
		mmo_no = mmo_no  + "  " +  childList[3].text;
	  }	  	  
	  
    }

	//alert('currentMachineType = ' +  currentMachineType);
	//alert('verifyMachineReq() = ' +  verifyMachineReq());
	
		//copy all task into table pmstepsupdate, if it is first time selected
		if(verifyMachineReq() == 0){ //not registerred  yet
			
			//configuration of APPS
			if(currentMachineType == "APPS"){
				openSwilWin('ECBM_APPS_CONFIG');
			}

			//configuration of AFCS
			if(currentMachineType == "AFCS"){
				openSwilWin('ECBM_AFCS_CONFIG');
			}
			
			//configuration of CIOSS/DIOSS
			if(currentMachineType == "CIOSS" || currentMachineType == "DIOSS"  || currentMachineType == "DBCS"){
				openSwilWin('ECBM_CIOSS_CONFIG');
			}
			
			if(currentMachineType == "SPBS" || currentMachineType == "APBS"  || currentMachineType == "SPBM"){
				openSwilWin('ECBM_SPBS_CONFIG');
			}
			
			if(currentMachineType == "PSM"){
				openSwilWin('ECBM_PSM_CONFIG');
			}
			
			if(currentMachineType == "HOPS"){
				openSwilWin('ECBM_HOPS_CONFIG');
			}
			
			if(currentMachineType == "EPPS"){
				openSwilWin('ECBM_EPPS_CONFIG');
			}
			
			var result = copyPMTasks();
			if(result.length < 0)  //return -1 or the total number of tasks copied, eg, apps==120,afcs==94,afcs=43
			{alert("Error to add the machine "+currentMachinePtr+"  into : "+result);}
			else{
				 ; //alert("OK to add the machine "+currentMachinePtr+"  into : "+result);
			} 
		}
		
		//SPBS - Number of Bins reminding
		if(currentMachineType == "SPBS" || currentMachineType == "APBS"  || currentMachineType == "SPBM"){
			if(getNoOfBins_SPBS()*1 == '999') { // num-of-bins is not updated yet.
				openSwilWin('ECBM_SPBS_CONFIG');
			}
		}
	}
}

//For SPBS/APBS.SPBM, return number of sort module and number of bins separated by "-". 6-100. 
//if num of bins is not available, return 'nn-999'
function getNoOfBins_SPBS()
{
	var objHTTP=new ActiveXObject('microsoft.XMLHTTP');	
    var urlStr=getServerPath()+'ecbm.ecbm_machine_lists.getNoOfUnits?p_machineid='+currentMachinePtr+'&p_site='+currentSite;
   objHTTP.open('GET', urlStr, false);
   objHTTP.send();

	//return objHTTP.responseText;	
	
	var v_arr = objHTTP.responseText.split("-");
	//var v_unitNo = v_arr[0];
	var v_binNo = v_arr[1];
	
	return v_binNo;
}



//get the User's level. return MU, LA,HQ,AD,AM, or NO FOUND if the currentUser is invalid 
function getUserLevel()
{
  var objHTTP=new ActiveXObject('microsoft.XMLHTTP');	
  var urlStr=getServerPath()+'ecbm.ECBM_MACHINE_LISTs.getUserLevel?p_site='+currentSite+'&p_user='+currentUser;
  objHTTP.open('GET', urlStr, false);
  objHTTP.send();

	return objHTTP.responseText;	
}

function getSiteName()
{
  	var objHTTP=new ActiveXObject('microsoft.XMLHTTP');	
    var urlStr=getServerPath()+'ecbm.ECBM_MACHINE_LISTs.getSiteName?p_facid='+currentWhepFacID;
   objHTTP.open('GET', urlStr, false);
   objHTTP.send();
	return objHTTP.responseText;	
}

function getUserInfo(){
	var j,i;
  var childList = null;
	var objAcroList = null;
	var objAcros=new ActiveXObject('microsoft.XMLDOM');
 	objAcros.async=false;
  objAcros.load("ecbm.ecbm_login.getUserInfo?p_user="+currentUser);
  if (objAcros.parsed) {
    objAcroList=objAcros.getElementsByTagName("ROW"); 
	  if(objAcroList.length != 0){ 
	  	return objAcroList;
	  }	   
  }  
  return null;
}  

//copy all regular tasks, except for the warnings
function copyPMTasks()
{	
  	var objHTTP=new ActiveXObject('microsoft.XMLHTTP');	
    var urlStr=getServerPath()+'ecbm.ECBM_MACHINE_LISTs.copyPmtasks?p_machineid='+currentMachinePtr+'&p_machinetype='+currentMachineType+'&p_classcode='+currentMachineClassCode+'&p_user='+currentUser+'&p_site='+currentSite+'&p_facid='+currentWhepFacID;
	
   objHTTP.open('GET', urlStr, false);
   objHTTP.send();

	return objHTTP.responseText;	
}

function verifyMachineReq()
{
  	var objHTTP=new ActiveXObject('microsoft.XMLHTTP');	
    var urlStr=getServerPath()+'ecbm.ECBM_MACHINE_LISTs.verifyMachineReq?p_machineid='+currentMachinePtr+'&p_site='+currentSite;

   objHTTP.open('GET', urlStr, false);
   objHTTP.send();
	return objHTTP.responseText;	
}

var reqElect ='';


function doModalResult(p_returnValue){
  currentReturnVal = p_returnValue;
  if (p_returnValue == "NONE") {
  	self.close();
  }

  if (!l_facility_passed) {
	var subsiteidandname = p_returnValue;
	currentECBMSubsite = subsiteidandname.substring(0,2);
	var len;
	len = subsiteidandname.length;
	siteName = subsiteidandname.substring(3, len);
    l_facility_passed = true;
  }
}  


function ECBMMenu() {
//hirad

//<!-- ================= -->
//<!-- Main canvas table -->
//<!-- ================= -->
document.all.div_menuheader.innerHTML = '*** Main Menu ***';
var menuStr = '<br><DIV><SPAN STYLE="width:630;align:center;font: 10pt arial black;text-align:center;top:5px;left:5px; border-style: outset;border-color:grey;border-width:thick;color: #ff0000; !important; padding:paddingThickness{2,40}">please specify an equipment No. above before going any further!</SPAN></DIV><br>'+
'<fieldset style="width:630">'+
	'<div align="center">'+
	'<table cellpadding="5" cellspacing="0" width="630" height="419" border="0" id="table1">'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td height="29">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3">'+
			'<img border="0" src="/emars-img/ecbmLogo.JPG" width="200" height="50" align="center" hspace="3"></td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td height="55">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td height="29">&nbsp;</td>'+
			'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="5">'+
			'<input type="button" onClick=doMenu("Checklist") style="width: 319; height:40; color: blue; font-weight:bold" value="Checklist" /></td>'+
			'<td>&nbsp;</td>'+
			'<td height="52">&nbsp;</td>'+
		'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=doMenu("Forecast") style="width: 270; height:40; color: blue; font-weight:bold" value="Forecast" /></td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3"><input type="button" onClick=dosubMenu("PMTask") style="width: 270; height:40; color: blue; font-weight:bold" value="PM Task" /></td>'+
			'<td height="57">&nbsp;</td>'+
		'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3">'+
			'<input type="button" onClick=dosubMenu("EquipmentProfile") style="width: 270; height:40; color: blue; font-weight:bold" value="Equipment Profile" /></td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="3">'+
			'<input type="button" onClick=dosubMenu("Management") style="width: 270; height:40; color: blue; font-weight:bold" value="Management" /></td>'+
			'<td height="52">&nbsp;</td>'+
		'</tr>'+
		'<tr>'+
			'<td>&nbsp;</td>'+
			'<td>&nbsp;</td>'+
			'<td valign="top" colspan="5">'+
			'<input type="button" onClick=dosubMenu("Reports") style="width: 320; height:40; color: blue; font-weight:bold" value="Reports" /></td>'+
			'<td>&nbsp;</td>'+
			'<td height="53">&nbsp;</td>'+
		'</tr>'+
		'<tr>'+
			'<td width="27">&nbsp;</td>'+
			'<td width="127">&nbsp;</td>'+
			'<td width="50">&nbsp;</td>'+
			'<td width="88">&nbsp;</td>'+
			'<td width="4">&nbsp;</td>'+
			'<td width="97">&nbsp;</td>'+
			'<td width="46">&nbsp;</td>'+
			'<td width="115">&nbsp;</td>'+
			'<td height="62" width="18">&nbsp;</td>'+
		'</tr>'+
	'</table>'+
	'</div>'+
	'</fieldset>';

	document.all.divECBMMenu.innerHTML=menuStr;
	document.getElementById("div_btn_s").style.display = "none";
}

//eCBM Menu items
function dosubMenu(val){
	//validate the input 
   	if (document.all.ecbm_machinetype_list.selectedIndex == 0 ){
		  document.all.ecbm_machineno_list.length = 0;
		  alert("Please select equipment type!");
		return;
	}
	
  // Check View Only
  if (ecbm_view_only != "N"){
    if (val=="EquipmentProfile" || val=="PMTask"){
    	 alert("Option Not Available for this Equipment.");
 		   return;
	  }	
	}  

	//'Management' module
	if (val=="Management") {
		if(currentUserRightECBM == "N"){
			alert("You do not have enough privilege to access this page.");
			return;
		}	

		document.all.div_menuheader.innerHTML = '*** Management ***';
   	    document.all.divECBMMenu.innerHTML=menumanageStr;
		document.getElementById("div_btn_s").style.display = "block";
		return;			
    }

	if (document.all.ecbm_machineno_list.selectedIndex == 0 ){
		alert("Please select equipment no.!");
		return;
	}

	//'Reports' module, everybody can access 	
	if (val == "Reports"){
		document.all.div_menuheader.innerHTML = '*** Reports ***';	
		document.all.divECBMMenu.innerHTML=menureportStr;
		document.getElementById("div_btn_s").style.display = "block";
	}

	//'PM Task' module
	if (val=="PMTask") {
		//only users who have 'S' or 'W' privildge can access
		if(currentUserRightECBM == "N"  || currentUserRightECBM == "R"){
			alert("You do not have enough privilege to access this page.");
			return;
		}
		document.all.div_menuheader.innerHTML = '*** PM Task ***';
   		document.all.divECBMMenu.innerHTML=menupmtaskStr;
		document.getElementById("div_btn_s").style.display = "block";			
    }
	
	//'Equipment Profile' module
	if (val=="EquipmentProfile") {
		//only users who have 'S' or 'W' privildge can access
		if(currentUserRightECBM == "N"){
			alert("You do not have enough privilege to access this page.");
			return;
		}
		document.all.div_menuheader.innerHTML = '*** Equipment Profile ***';
   		document.all.divECBMMenu.innerHTML=menuequipmentprofileStr;
		document.getElementById("div_btn_s").style.display = "block";			
    }
}

//eCBM submenu, eg. Reports-->History 
function doMenu(val) {
   gCurrentPtr='';
   gOrderType='';

//alert("debug: --- before open window of Checklist"); 

   //validate the input 
  if (document.all.ecbm_machinetype_list.selectedIndex == 0 ){
		document.all.ecbm_machineno_list.length = 0;
		alert("Please select equipment type!");
		return;
	}	

  if (ecbm_view_only != "N"){
    if (val=="Forecast" || val=="EORData"){
    	 alert("Option Not Available for this Equipment.");
 		   return;
	  }	
	}  
	

	if (document.all.ecbm_machineno_list.selectedIndex == 0 ){
		if(val=="AnnualEstimate" || val=="DailyEstimate" || val=="AllChecklists" || val=="CompletionReports" || val=="AllLocalTasks" || val=="AllSeniorTasksDue" || val=="ChecklistHours")
		{  //alert("Equipment No. does not matter in this module." ); 

		}else{
			alert("Please select Equipment No.!");
			return;
		}
	}

	//checklist module, everybody can access
   if (val=="Checklist") {
		
		//if(currentMachineType == "CIOSS"), verify the equipment size
		if(currentMachineType == "CIOSS" || currentMachineType == "DIOSS"  || currentMachineType == "DBCS"){
			var m_size = getCIOSSsize();
			
			if(isNaN(m_size) || m_size*1 < 110 || m_size*1 > 302){
				//w_report = openSwilWin('ECBM_CIOSS_CONFIG'); 
				alert("Please click on 'EQUIPMENT PROFILE' to update the equipment size, it is required for the further activities."); 
				return;
			}
		}
		
		w_report = openSwilWin('ECBM_SUBMIT_LIST'); 

   }

   //reports module	
   if (val=="PMHistory") {w_report = openSwilWin('ECBM_PM_HISTORY');} 
   if (val=="EORData") {w_report = openSwilWin('ECBM_EOR_DATA');}
   if (val=="TaskThreshold") {w_report = openSwilWin('ECBM_TASK_THRESHOLD');} 
   if (val=="BypassedTask") { w_report = openSwilWin('ECBM_BYPASSED_TASK');}
   if (val=="CompletinRate") { w_report = openSwilWin('ECBM_COMPLETION_RATE');}
   if (val=="TaskCompleted") { w_report = openSwilWin('ECBM_TASK_COMPLETED');}

 	//forecast module
   if (val=="Forecast") {
		//jesse 2014.3.25 temporary blocking it for FSS/SAMP/FSSOCR/FSSSMS 
		if (currentMachineType == "FSS" || currentMachineType == "FSSSMS" || currentMachineType == "FSSOCR" || currentMachineType == "SAMP"){
			alert("This tool is temporarily not available for " +  currentMachineType);
			return;
		}
		
   		//only users who have 'S' privilege can access 		
		if(currentUserRightECBM != "S"){
				alert("You do not have enough privilege to access this page.");
				return;
		}	
   		w_report = openSwilWin('ECBM_FORECAST');
	} 

   //pm task module
   if (val=="CreateLocalTask") { w_report = openSwilWin('ECBM_LOCAL_TASK');} 
   if (val=="EditLocalTask") { w_report = openSwilWin('ECBM_LOCAL_TASK_LIST');} 
   if (val=="SuspendTask") { w_report = openSwilWin('ECBM_SUSPEND_TASK1');} 
  //if (val=="Scheduling") { w_report = openSwilWin('ECBM_SHEDULING');} 

   //equipment profile module
   if (val=="UpdateAPPSConfig") { 
   		if (currentMachineType != "APPS"){
			alert("This page is only for APPS equipment!");
			return false;
		}		
   		w_report = openSwilWin('ECBM_APPS_CONFIG');
	} 
	
	if (val=="UpdatePSMConfig") { 
   		if (currentMachineType != "PSM"){
			alert("This page is only for PSM equipment!");
			return false;
		}		
   		w_report = openSwilWin('ECBM_PSM_CONFIG');
	}
	
	if (val=="UpdateHOPSConfig") { 
   		if (currentMachineType != "HOPS"){
			alert("This page is only for HOPS equipment!");
			return false;
		}		
   		w_report = openSwilWin('ECBM_HOPS_CONFIG');
	}
	//20260420 
	if (val=="UpdateEPPSConfig") { 
   		if (currentMachineType != "EPPS"){
			alert("This page is only for EPPS equipment!");
			return false;
		}		
   		w_report = openSwilWin('ECBM_EPPS_CONFIG');
	}
	if (val=="UpdateHTPSConfig") { 
   		if (currentMachineType != "HTPS"){
			alert("This page is only for HTPS equipment!");
			return false;
		}		
   		w_report = openSwilWin('ECBM_HTPS_CONFIG');
	}
	//20260420
	if (val=="UpdateAFCSConfig") { 
   		if (currentMachineType != "AFCS"){
			alert("This page is only for AFCS equipment!");
			return false;
		}		
   		w_report = openSwilWin('ECBM_AFCS_CONFIG');
	}
	
	if (val=="UpdateCIOSSConfig") { 
   		if (currentMachineType != "CIOSS"){
			if (currentMachineType != "DIOSS"){
				if (currentMachineType != "DBCS"){
					alert("This page is only for CIOSS, DIOSS or DBCS equipment!");
					return false;
				}
			}
		}		
   		w_report = openSwilWin('ECBM_CIOSS_CONFIG');
	}


	if (val=="UpdateSPBSConfig") { 
   		if (currentMachineType != "SPBS"){
			if (currentMachineType != "SPBM"){
				if (currentMachineType != "APBS"){ 
					alert("This page is only for SPBS or APBS equipment!");
					return false;
				}
			}
		}
		//alert("Please make sure the number of inductions is correct, which will impact the est. times of eCBM tasks.");
   		w_report = openSwilWin('ECBM_SPBS_CONFIG');  
	}
	
	if (val=="UpdateRUNNINGDAY") { 
		//alert("You will be directed to set up the NON-RUNNINGDAY.");
		//only users who have 'S' privilege can access 		
		if(currentUserRightECBM != "S"){
				alert("You do not have enough privilege to access this page.");
				return;
		}
   		w_report = openSwilWin('ECBM_RUNNINGDAY');
	}
	
	if (val=="UpdateMaintWin") { 
   		w_report = openSwilWin('ECBM_MAINT_WIN');
	}
	
   //management module
   if (val=="AnnualEstimate") {
        if (currentMachineType != "APPS"){
			alert("This page is only for APPS equipment!");
			return false;
		}
   		w_report = openSwilWin('ECBM_APPS_AWORKLOAD');
   } 

   if (val=="DailyEstimate") {
   		if (currentMachineType != "APPS"){
			alert("This page is only for APPS equipment!");
			return false;
		}
   		w_report = openSwilWin('ECBM_APPS_DWORKLOAD');
   } 
   
   if (val=="AllChecklists") { w_report = openSwilWin('ECBM_CHECKLISTS');}
   if (val=="CompletionReports") { w_report = openSwilWin('ECBM_COMPLETIONS');} 
   if (val=="AllLocalTasks") { w_report = openSwilWin('ECBM_ALL_LOCAL_TASKS');} 
   if (val=="AllSeniorTasksDue") { w_report = openSwilWin('ECBM_ALL_SENIORS_DUE');} 
   if (val=="ChecklistHours") { w_report = openSwilWin('ECBM_CHECKLIST_HOURS');} 
	//alert("debug: --- after open window of Checklist"); 
}

//number of stackers for CIOSS, DIOSS, 
function getCIOSSsize()
{
	var objHTTP=new ActiveXObject('microsoft.XMLHTTP');	
    var urlStr=getServerPath()+'ecbm.ecbm_machine_lists.getMachineSize?p_machineid='+currentMachinePtr+'&p_site='+currentSite;
   objHTTP.open('GET', urlStr, false);
   objHTTP.send();

	return objHTTP.responseText;	
} 

/*
//------------------------------------ testing only  --------------------------------------- 
function openSwilWin22(pModule,objName)
{
   if (objName==null || objName.length==0) {
      objName=pModule;
   } 
   
  // alert("inside em_dialog.js --- openSwilWin() ---  111 ");
   var winSize=getXMLresults("EM_XML_MODULE_SIZE_PRC?p_screen_id="+pModule); 
  // alert("inside em_dialog.js --- openSwilWin() ---  222 ");
   
   var h;
   var w;
   if (winSize.parsed) {
      var mHeight=winSize.getElementsByTagName("HEIGHT")
      var mWidth=winSize.getElementsByTagName("WIDTH")
      h=mHeight(0).text;
      w=mWidth(0).text;
   } else {
      h=600;
      w=800;
   } 
   
  // alert("inside em_dialog.js --- openSwilWin() ---  333 ");
   var urlStr="!vm_main_pkg.main?pUser=" + currentUser + "&pSiteID=" + currentSite + "&pModule="+pModule;
  // alert("inside em_dialog.js --- openSwilWin() ---  444 ");
											  
   var winleft = (screen.width / 2) - (w/2);
   var wintop = (screen.height / 2) - (h/2);
   var winProps="left="+winleft+",top="+wintop+",width="+w+",height="+h+",status=no,menubar=no,toolbar=no,resizable=no,scrollbars=no";
   var exeStr;
   
   try  { 
   		alert("window.ECBM_SUBMIT_LIST= " + window.ECBM_SUBMIT_LIST );
   		if (objName != null) {
			alert("objName = " + objName ) ;
			//objName = 'undefined';
		}
		
		//eval("window.ECBM_SUBMIT_LIST.close()");
		
	  // alert("inside em_dialog.js --- openSwilWin() ---  555 ");
	    exeStr=objName+"=window.open(\""+urlStr+"\",\""+pModule+"\",\""+winProps+"\")";
		 eval(exeStr);
		 
       exeStr=objName+".focus()"; 
	   //alert("inside em_dialog.js --- openSwilWin() ---  555666 : " + exeStr)
       eval(exeStr);
	  
	   //alert("inside em_dialog.js --- openSwilWin() ---  666 ");
	  
   } catch(e) {
	  exeStr=objName+"=window.open(\""+urlStr+"\",\""+pModule+"\",\""+winProps+"\")";
	  
	   alert("inside em_dialog.js --- openSwilWin() ---  6767: " + exeStr);
	   
      eval(exeStr);
      childWindows[childWindows.length]=objName;
	  
	  alert("inside em_dialog.js --- openSwilWin() ---  777 ");
   }
   
   //alert("inside em_dialog.js --- openSwilWin() ---  888 ");
}

*/




