/* Tehseen 10/10/2003*/
var currentFY = '2013';
var nextFY = '2014';
var currentSubsite;
var lmode="work";
var mode;
var equipall='diff';
var rights='X';
var findWinWO ='';
var allopms_called = 'N';
var ckhopms ='Y';
var endfy;
var listallregno='N'; 
var last_day_of_current_year='09/25/2015'; 
var last_day_of_next_year = '09/30/2017';
var currentLocalDate; 
var editmode='N';
var v_currentSite = 0; //Michael E. Harris - 12/07/2011
var v_cert_comments = ''; //Michael E. Harris - 04/10/2015
var v_past_due_saved = false;



function bodyOnLoad() { 

 
	
   v_currentSite = currentSite; 
	
   disableField('wo_status');   
   //cprice added onkeypress to this select 1/10/2005
   document.all['divEquipNo'].innerHTML='<font color=black>* </font><b>Equip No</b><BR><select name="equip_no" id="equip_no" style="width: 90" TABINDEX="1"  onKeyPress="typeAhead(); enterequip();"></select>';              
   document.all['divClass'].innerHTML='<font color=black></font><b>Class</b><BR><input type="text" name="class_code" id="class_code" TABINDEX="1" style="width: 90"  />';              
   document.all['divwono'].innerHTML='<font color=black>*</font><b>WO NO</b><BR><input type="text" name="wo_no" id="wo_no" size="6" TABINDEX="1" onKeyPress="enterwono()" onblur="localonBlur(this.name)" />';       
   disableField('btnMaterial'); 
   disableField('btnReport'); 
   disableField('btnMTSC'); 
   disableField('btnInfo'); 
   disableField('btnMaterialsAdd'); 
   disableField('btnSave'); 
   disableField('btnDelete'); 
   disableField('wo_desc_key');  
   disableField('hours_to_date');
   disableField('reg_no_key');
   disableField('fy_key');
   disableField('wo_no_key');
   disableField('wo_no');
   //disableField('modification_no');
   document.all['hours_to_date'].value=0;
   disableField('subsite_id');
   disableField('acronym');
   disableField('equip_no');
   disableField('class_code');
   //===============================
   // SASurikov
   
   disableField('nm_selBulletin_3Alpha');
   disableField('nm_txtBulletin_3Num');
   disableField('nm_selBulletin_2Num');
   //===============================
   
   //==============================
   // Michael E. Harris  -- Loading the "Past due reason dropdown"
   loadList('past_due_reason_code',getServerPath()+"em_work_pkg.em_XMLPastDueReasonList");
   disableField('past_due_reason_code');
   
   //document.all['txt_fssmaid'].readOnly = true;
   document.all['txt_fssmaid'].disabled = true

  //var objEndFY = new ActiveXObject('microsoft.XMLDOM');
  //objEndFY.async=false;
  //jesse this procedure will return the last day of current fiscal year in format of MM/DD/YYYY. 
  /*objEndFY.load('em_work_pkg.em_next_fiscal_yr_prc');
  if(objEndFY.parsed) {
   var loadlist = objEndFY.selectSingleNode('NEXTFYDT');
   nextfy = objEndFY.text;
   last_day_of_current_year = nextfy; 
  }
  objEndFY=null; */ 
  
  //jesse this procedure will return the last day of current fiscal year 
  //	based on the local time
  //	in format of MM/DD/YYYY.
  var objEndFY = new ActiveXObject('microsoft.XMLDOM');
  objEndFY.async=false;
  objEndFY.load('em_time_pkg.em_lastday_fiscal_yr_prc?p_siteID='+currentSite); 
  
  if(objEndFY.parsed) { 
   	var loadlist = objEndFY.selectSingleNode('LASTFYDT');
   	last_day_of_current_year = objEndFY.text;  
  }
  
  //jesse example 2014
  currentFY = last_day_of_current_year.substring(last_day_of_current_year.length-4,last_day_of_current_year.length);
  nextFY = currentFY*1 + 1;
  
  //get the last day of NEXT fiscal year
  objEndFY.load('em_time_pkg.em_lastday_fiscal_yr_prc?p_siteID='+currentSite+'&p_fy='+nextFY); 
  if(objEndFY.parsed) { 
   	var loadlist = objEndFY.selectSingleNode('LASTFYDT');
   	last_day_of_next_year = objEndFY.text;  
  }
  objEndFY=null;

}

function ps8210() {
        var so=window.open("http://blue.usps.gov/formflow/xft/psform8210.xft");

}

//Jesse 3/31/2014
// retrieve the next work order number for a non-standing work order
function get_non_standing_wono() {
	
	var objHTTP = new ActiveXObject('microsoft.XMLHTTP');	
	var urlStr = getServerPath()+'em_work_pkg.get_non_standing_wono?p_siteid='+currentSite+'&p_regno='+document.all['reg_no'].value;
	objHTTP.open('GET', urlStr, false);
	objHTTP.send();
	
	return objHTTP.responseText;	   
}

//most of settings are in onChange_RegisterNo()
function onREGChange(){ 
  
  //loadList('subsite_id', getServerPath()+"em_work_pkg.em_regsubsite_prc?pSiteID="+currentSite+"&pRegNo="+document.all['reg_no'].value, 'N');      
  //jesse see below  onssChange();
  
  //jesse 2012.5.18
  if(document.all['reg_no'].value == '-1') {
	  	//jesse 2015-1-29
		f_set_EmptStr_Into_FSSP();
		disableField('txt_fssp_call');
    	disableField('txt_fssp_problem');
    	disableField('txt_fssp_zip');
	
	 	disableField('acronym');
		document.all.reg_no.focus();
  }else{
	  
	  /* 2020.12.04 move this block to localblur() due to the selecting by keyboard 
	  	//jesse 2015.1.29
		var v_regno = document.getElementById('reg_no').value;
		var v_regno_text = document.getElementById('reg_no').options[document.getElementById('reg_no').options.selectedIndex].text;
		
		//alert('v_regno='+v_regno);alert('v_regno_text='+ v_regno_text);
		
		var fssp_enable = 'N';
		var reg_no23 = v_regno.substring(1,3);
		var objAcros=new ActiveXObject('microsoft.XMLDOM');              
   		objAcros.async=false; 
   		objAcros.load("em_work_pkg.em_get_fssp_status?pSiteID="+currentSite+"&pRegno="+document.all['reg_no'].value);
		if (objAcros.parsed) {       
      		var whatever = objAcros.selectSingleNode('fssp_enable');
			fssp_enable = objAcros.text;
		}else{
			fssp_enable = 'N';	
		}
		
	  	if(fssp_enable == 'Y' || v_regno_text.indexOf("FMO",0) != -1) {
	  		if(cust_confirm('FSSP Call','Does this work order require any FSSP calls?',3,4,1,0) != 6) {
				f_set_EmptStr_Into_FSSP();
				disableField('txt_fssp_call');
    			disableField('txt_fssp_problem');
    			disableField('txt_fssp_zip');
			}else{
				enableField('txt_fssp_call');
				enableField('txt_fssp_problem');
				enableField('txt_fssp_zip');
			}
		}else{
			f_set_EmptStr_Into_FSSP();
			disableField('txt_fssp_call');
    		disableField('txt_fssp_problem');
    		disableField('txt_fssp_zip');
		}
		//jesse
	*/
	
	  	onssChange(); //load acronym list
		enableField('acronym');  
		
		//move control to standing flag 
  		//document.getElementById('standing_flag').focus();
  }
  
  document.getElementById('wo_no').value = get_non_standing_wono();  //jesse 2014.3.31	   
}

//sub site onChange()
function onssChange() { 

	if(currentSite == '-303'){alert(' line 194 within ossChange() - begining');  };
	
	enableField('whep_fac_id');
	loadList('whep_fac_id', getServerPath()+"em_work_pkg.getXMLFacility_whep_fac_id?pSiteID="+currentSite+"&pSubsiteID="+document.all['subsite_id'].value, 'Y');
	
 	if (document.all.acronym.options.length <1 ) {  
    	loadList('acronym', getServerPath()+"em_work_pkg.em_XMLGetAcroListEquipCnt_add?pSiteID="+currentSite+"&pSubsite="+document.all['subsite_id'].value, 'Y'); 
 	}  
   	enableField('acronym');
	
	if(currentSite == '-303') {alert(' line 204 within ossChange() - ending'); };
}

function sfchange(){ //sf   (Standing Flag change)  
	if(currentSite == '-303'){ alert('line 208 starting of sfchange()'); };
	 
		//also see onChange_registerNo()
		var reg_no = document.all.reg_no.value; 
		 
		 if (document.all['standing_flag'].value == 'Y') {
			   document.all['divwono'].innerHTML='<font color=black>* </font><b>WO NO</b><BR><select name="wo_no" id="wo_no" TABINDEX="1" onKeyPress="enterwono()" style="width: 60"></select>';           
			   loadList('wo_no', getServerPath()+"em_work_pkg.em_workordernum_prc?pSite="+currentSite+"&pRegNo="+document.all['reg_no'].value, 'N'); 
			   if(document.getElementById('wo_no').options[document.getElementById('wo_no').options.selectedIndex].value == '-1'){
					//have created 100 standing work orders
					alert('Total 100 standing work orders have been created, no more is allowed.');
					document.all['standing_flag'].value == 'N';
					onChange_RegisterNo();
					return;  
			   }
			   document.all['priority_code'].value='D';                  
			   disableField('priority_code');
			   document.all.date_due.value = last_day_of_current_year;
			   disableField('date_due');
			   disableField('est_hours');
			   
			   if(reg_no.substring(1,3) == '01' || reg_no.substring(1,3) == '02' || reg_no.substring(1,3) == '09'  || reg_no.substring(1,3) == '14' ){
				   //2015.8.24 excluding work codes 05,08,18,23
				   //2020.11.12 excluding 09,24,11,12,06,21 for register x01,x02;  09 for x09; 24 for x14. 
				   loadList('work_code', getServerPath()+"em_work_pkg.em_getXMLCmWorkCodeListSWO_reg?pRegNo="+reg_no,'Y');
			   }else{
				   //2015.8.24 excluding work codes 05,08,18,23
				   //2020.11.12 excluding 21 
				   loadList('work_code', getServerPath()+"em_work_pkg.em_getXMLCmWorkCodeListSWO",'Y');
			   };
			   
			   
		 }else{
			   document.all['divwono'].innerHTML='<font color=black>* </font><b>WO NO</b><BR><input type="text" name="wo_no" id="wo_no" size="6" TABINDEX="1" onKeyPress="enterwono()" onblur="localonBlur(this.name)" />';      
			   document.all['wo_no'].value = get_non_standing_wono();  //jesse 2014.3.31		   
			   disableField('wo_no'); 
			   document.all['priority_code'].value='';                  
			   enableField('priority_code');
			   enableField('est_hours');
			   enableField('date_due');
			   document.all.date_due.value='';
			   
			   //2015.8.24 to reload the work code list including work codes 05,08,18,23
			   //2020.11.17 to exclude wc 06,11. 09 for x09; 24 for x14 
			   loadList('work_code', getServerPath()+"em_work_pkg.em_getXMLCmWorkCodeList_reg?pRegNo="+reg_no,'Y');
			   
		 } 
 
		//if(reg_no.substring(0,1) == '5'){ //fy2015, then work code for Mail Search changed from 24 to 14 
		if(1 == 2) { 
			if (reg_no.substring(1,3) == '09' || reg_no.substring(1,3) == '24'){ 
					if(reg_no.substring(1,3) == '09'){
						//alert('The work code must be 09.');
						document.all['work_code'].value = '09';
						
					}else{
						//alert('The work code must be 24.');
						document.all['work_code'].value = '24';
					}
					
					disableField('work_code');
			}else{
					enableField('work_code');
					document.all['work_code'].value = '';
			} 
		
		}else{
			
			if (reg_no.substring(1,3) == '09' || reg_no.substring(1,3) == '14'){ 
					if(reg_no.substring(1,3) == '09'){
						//alert('The work code must be 09.');
						document.all['work_code'].value = '09';
					}else{
						//alert('The work code must be 24.');
						document.all['work_code'].value = '24';
					}
					
					disableField('work_code');
			}else{
					enableField('work_code');
					document.all['work_code'].value = '';
			} 
		}
	 		   
	//jesse 2014.3.31 re-load acronyms ( ==> than  re-load equip no  ) 
	if(mode == 'add'){
		if(reg_no.substring(1,3) == '09') {
			loadList('acronym', getServerPath()+"em_work_pkg.em_XMLGetAcroListEquipCnt_add?pSiteID="+currentSite+"&pSubsite="+document.all['subsite_id'].value+"&pAcronymGroupCodes=A&pRegNo="+reg_no, 'Y'); 
		}else{
			if(reg_no.substring(1,3) == '01') {
				loadList('acronym', getServerPath()+"em_work_pkg.em_XMLGetAcroListEquipCnt_add?pSiteID="+currentSite+"&pSubsite="+document.all['subsite_id'].value+"&pAcronymGroupCodes=AMNO&pRegNo="+reg_no, 'Y'); 
			}else{ 
				if(reg_no.substring(1,3) == '02') {
					loadList('acronym', getServerPath()+"em_work_pkg.em_XMLGetAcroListEquipCnt_add?pSiteID="+currentSite+"&pSubsite="+document.all['subsite_id'].value+"&pAcronymGroupCodes=BO&pRegNo="+reg_no, 'Y'); 
				}else{
					loadList('acronym', getServerPath()+"em_work_pkg.em_XMLGetAcroListEquipCnt_add?pSiteID="+currentSite+"&pSubsite="+document.all['subsite_id'].value+"&pRegNo="+reg_no, 'Y'); 
				}
			}
		} 
		
		document.all.equip_no.value='';
		disableField('equip_no');
	
		//jesse 2014.9.18
		datevalidate();
	}	

	if(currentSite == '-303'){
		alert('line 316 ending of sfchange()'); 
		};
}

function todaydate(date,pAdd){
   var gNowa = new Date(date);
   var gNow = new Date(gNowa.getTime() + pAdd*864e5);   
   var vNowDay = gNow.getDate();
   var vNowMonth = gNow.getMonth();
   vNowMonth = vNowMonth +1;
   var vNowYear = gNow.getFullYear();
   vNowDay = vNowDay;
   var vret = vNowMonth+'/'+vNowDay+'/'+vNowYear;
   return (vret);
}

function todaydate1(s,pAdd) {
  var objDate = new ActiveXObject('microsoft.XMLDOM');
  objDate.async = false;
  objDate.load("em_work_pkg.em_date_prc?pStartDate="+s+"&pAdd="+pAdd);
  if (objDate.parsed) {
    var result = objDate.selectSingleNode('caldate');
    var newdate = objDate.text;
  }
  return (newdate);
}



function localvalidate(fieldNameStr) {
   if (fieldNameStr=='date_issued') {   ;
    /* //Jesse 2014.9.17 the date-input icon will go around the date_issued validation function datevalidated()
	 if (document.all['priority_code'].value=='A'){
         document.all['date_due'].value = todaydate1(document.all.date_issued.value,1);
      }
      if (document.all['priority_code'].value=='B'){
         document.all['date_due'].value = todaydate1(document.all.date_issued.value,7);
      }
      if (document.all['priority_code'].value=='C'){
         document.all['date_due'].value = todaydate1(document.all.date_issued.value,30);
      } */
   }
} 


//---------------------------------------------------------------------------------------------------
// Added by Michael E. Harris for Bulletin Work Order Certification
// This procedure displays the Work Order Certification Dialogue and prompts the user
// for his/her password to certify the current work order.
function f_btnCertify_OnClick() {

    var vDialogWidth=605;
    var vDialogHeight=520;     
    var winleft = (screen.width / 2) - (vDialogWidth / 2); 
    var wintop = (screen.height / 2) - (vDialogHeight / 2); 
	var v_wo_ptr = document.getElementById('wo_ptr').value;
	var v_result = '';
	var v_certify_result = '';
	
	do {
		v_result = window.showModalDialog(getServerPath()+"emapp.em_work_pkg.em_wo_certification_dialog?pSiteID=" + currentSite + "&pCurrentUser=" + currentUser + "&pWo_ptr="+v_wo_ptr +"", "", "dialogWidth:" + vDialogWidth + "px; dialogHeight:" + vDialogHeight + "px; center:yes; status:no; resizeable:no");
		if (v_result != '' && v_result != null && v_result =='SUCCESS') {
		/*	var v_urlStr = getServerPath()+"emapp.em_work_pkg.em_wo_certification?pSiteID="+currentSite + 
								"&pCurrentUser="+currentUser + 
								"&pWo_ptr="+v_wo_ptr;
			var objWOCert = new ActiveXObject('microsoft.XMLDOM');
			  objWOCert.async = false;
			  objWOCert.load(v_urlStr);
			if (objWOCert.parsed) {
				var nodelist = objWOCert.selectSingleNode('RESULT');
				v_certify_result = objWOCert.text;*/
				
				//if (v_certify_result == 'SUCCESS') {
							
			alert('The Work Order was certified.');
			document.getElementById('btnCertify').style.display = 'none';
			document.getElementById('capDiv').innerHTML = 'WORK ORDER DETAIL <FONT COLOR = "GREEN">(CERTIFIED)</FONT>';
			break;
		}else if (v_result == 'NO_PRIV') {
			alert('You are not authorized to certify this work order.  Only the Maintenance Manager can certify work orders.');
			document.getElementById('capDiv').innerHTML = 'WORK ORDER DETAIL <FONT COLOR = "RED">(NOT CERTIFIED)</FONT>';
			break;
		}else if (v_result == 'INCORRECT_PASSWORD') {
				alert('You have entered an incorrect password.  Please try certification again.');
				document.getElementById('capDiv').innerHTML = 'WORK ORDER DETAIL <FONT COLOR = "RED">(NOT CERTIFIED)</FONT>';
		}else if (v_result =='CANCEL' ) {
				document.getElementById('capDiv').innerHTML = 'WORK ORDER DETAIL <FONT COLOR = "RED">(NOT CERTIFIED)</FONT>';
				break;
		}else if (v_result=='ERROR' ) {
			alert('An error has occurred during the certification process.  If this persists, please notify the eMARS help desk.');
			document.getElementById('capDiv').innerHTML = 'WORK ORDER DETAIL <FONT COLOR = "RED">(NOT CERTIFIED)</FONT>';
			break;
		}
	} while (v_certify_result != 'SUCCESS');
}

// Added by Michael E. Harris for Bulletin Work Order Certification
// This procedure displays the Work Order Certification Dialogue and prompts the user
// for his/her password to certify multiple work orders.
function btnCertifyAllWos_onClick() {
	var objCert = new ActiveXObject('microsoft.XMLDOM');
	objCert.async = false;
	objCert.load('emapp.em_work_pkg.isUserWoCertifier?pSiteId='+currentSite+'&pCurrentUser='+currentUser);
	if (objCert.parsed) {
		var v_result = objCert.selectSingleNode('RESULT');
		var v_isCertifier = objCert.text;
		if (v_isCertifier == 'Y') {
			var vDialogWidth=700;
			var vDialogHeight=700;     
			var winleft = (screen.width / 2) - (vDialogWidth / 2); 
			var wintop = (screen.height / 2) - (vDialogHeight / 2); 
			var v_result=window.showModalDialog("emapp.em_work_pkg.em_certification_list_dialog?pSiteID=" + currentSite + "&pCurrentUser="+
			currentUser+"", "", "dialogWidth:" + vDialogWidth + "px; dialogHeight:" + vDialogHeight + "px; center:yes; status:no; resizeable:no");
		}
	}
}
/*function btn4774_onClick() {

		var v_wo_ptr = document.getElementById('wo_ptr').value;
		var vDialogWidth=1000;
		var vDialogHeight=700;
		var winleft = (screen.width / 2) - (vDialogWidth / 2);
		var wintop = (screen.height / 2) - (vDialogHeight / 2);
		var vyy=showModalDialog("!vm_main_pkg.main?pSiteID="+currentSite+"&pUser="+currentUser+"&pModule=EM_COMP_WO_4774",v_wo_ptr,"dialogWidth:1000px; dialogHeight:700px; center:yes; status:no; resizeable:no");
		var vyy = openSwilWin2('EM_COMP_WO_4774');
	
	// End code by Michael E. Harris
}*/

//----------------------------------------------------------------------------------------------------
//SASurikov
function onChange_RegisterNo(){
	
	loadList('subsite_id', getServerPath()+"em_work_pkg.em_regsubsite_prc?pSiteID="+currentSite+"&pRegNo="+document.all['reg_no'].value, 'N'); 
	
	//jesse 2014.9.18 has to reset the Standing flag and priority code
	enableField('standing_flag');
	document.all['standing_flag'].value = 'N';  //set the default value of STANDING = N 
	sfchange();  								//reload work code based on STANDING flag
	
	var nextdateobj = new Date(Date.parse(last_day_of_current_year)+86400000);
	var first_day_of_next_year = '0'+(nextdateobj.getMonth()+1) +'/'+nextdateobj.getDate()+'/'+nextdateobj.getFullYear();	
	if(nextdateobj.getMonth()*1 >= 9  ) { // no need the prefixed '0'
		  first_day_of_next_year = (nextdateobj.getMonth()+1) +'/'+nextdateobj.getDate()+'/'+nextdateobj.getFullYear();
	}  
	var reg_no = document.all.reg_no.value;
	if(reg_no.substring(0,1) != currentFY.substring(3) ){
		document.all['date_issued'].value = first_day_of_next_year;
	}else{
		document.all['date_issued'].value = currentDate;
	}
	
	/*  2015.9.13 to call the function sfchange() instead above */	
	/* jesse 2014-12-5 
	with(self.document.forms[0])
	{
		f_Enable_Desable_FSSP();
        f_set_EmptStr_Into_FSSP();
	}
	*/	
} 
// end of with() and function onChange_RegisterNo()
//
//------------------------------------------------------------------------------------------------------



//------------------------------------------------------------------------------------------------------
/*
//SASurikov:
function f_Enable_Desable_BulletinNumber(){with(self.document.forms[0])
{
// Per Kathy Carrico and Lloyd Bell: Three fields that "build" Bulletin Number must be available for the user if 
// the value of the last two numbers in Register ("reg_no") are equal or more than 91 or equal or less than 99.
//
	var vRegisterValue = "";
	var vTestValue 	   = "";
	var vRegisterTxt   = "";
	//
	for (var e=0; e<elements.length; e++)
    {
		if (elements[e].name == "reg_no" && elements[e].value != "")
		{
			vRegisterValue = reg_no.options[reg_no.options.selectedIndex].value;
			vRegisterTxt   = reg_no.options[reg_no.options.selectedIndex].text;
			//
			vTestValue = vRegisterValue.substr(1,2);
			vTestValue = parseInt(vTestValue, 10);
			//
			if ( 90 < vTestValue && vTestValue < 100)
			{
				enableField("nm_selBulletin_3Alpha");
				enableField("nm_txtBulletin_3Num");
				enableField("nm_selBulletin_2Num");
			}
			else
			{
				nm_selBulletin_3Alpha.options[0].selected = true;
				nm_txtBulletin_3Num.value = "";
				nm_selBulletin_2Num.options[0].selected = true;
				//
				disableField('nm_selBulletin_3Alpha');
   				disableField('nm_txtBulletin_3Num');
   				disableField('nm_selBulletin_2Num');
			}
		}
	}
}}
*/

function f_Enable_Desable_FSSP()
{
	with(self.document.forms[0])
	{//alert(modification_no.value)
        for (var e=0; e<elements.length; e++)
        {
                //alert(elements[e].name+"   "+elements[e].value)
				if (elements[e].name == "reg_no" && elements[e].value != "")
	   			{
					//--------------------------------------------------------------------------------
					var vRegisterValue = reg_no.options[reg_no.options.selectedIndex].value;
					var vRegisterTxt   = reg_no.options[reg_no.options.selectedIndex].text;
	
					if (vRegisterValue.indexOf("FMO",0) != -1 || vRegisterTxt.indexOf("FMO",0) != -1)					   
					{
							enableField('txt_fssp_call');
							enableField('txt_fssp_problem');
							enableField('txt_fssp_zip');
					}
					else if (vRegisterValue.indexOf("FMO",0) == -1 && vRegisterTxt.indexOf("FMO",0) == -1)
					{
							
							disableField('txt_fssp_call');
							disableField('txt_fssp_problem');
							disableField('txt_fssp_zip');
					}				
					//--------------------------------------------------------------------------------
					break;
				}
        }       
	}
} // end of with() and function f_Enable_Desable_FSSP()

//-----------------------------------------------------------------------------------------------------
function f_set_EmptStr_Into_FSSP(){with(self.document.forms[0])
{
        //
        txt_fssp_call.value = "";
        h_fssp_call.value   = "";
        //
        txt_fssp_problem.value = "";
        h_fssp_problem.value   = "";
        //
        txt_fssp_zip.value = "";
        h_fssp_zip.value   = "";
}} // end of with() and function f_Enable_Desable_FSSP()

//-------------------------------------------------------------------------------------------------------
function f_Trim(s) {
        s = s.replace(/(^\s*)|(\s*$)/gi,"");
        s = s.replace(/[ ]{2,}/gi," ");
        s = s.replace(/\n /,"\n");

        return s;
}

//=======================================================================================================
//jesse it is called ONLY when the FSSP calls are required 
function f_ValidateFSSP2(p_regno){with(self.document.forms[0])
{// this function is called from saveMainForm()
        var vIsValidationPassed = "YES";
        //var vRegisterValue = reg_no.options[reg_no.options.selectedIndex].value;
        //var vRegisterTxt   = reg_no.options[reg_no.options.selectedIndex].text;
        var objRegExp_onlyDigits  = /(^-?\d\d*$)/;
        var objRegExp_ZipCode     = /(^\d{5}$)|(^\d{5}-\d{4}$)/;
        var arrayOfCtrl = new Array("txt_fssp_call", "txt_fssp_problem", "txt_fssp_zip");
        var arrayOfLable = new Array("FSSP Call", "FSSP Problem", "FSSP ZIP");
        var vCtrlName  = "";
        var vCtrlValue = "";
        var vCtrlLable = "";
        var vMsg = "";
    
        for (var v = 0; v < arrayOfCtrl.length; v++){
        	vCtrlName  = arrayOfCtrl[v];
            vCtrlValue = eval(arrayOfCtrl[v]+".value.toUpperCase();"); 
            vCtrlValue = f_Trim(vCtrlValue);
            vCtrlLable = arrayOfLable[v];
                     
            if (vCtrlName == "txt_fssp_call"){
            	txt_fssp_call.value = vCtrlValue;
                      
                if (vCtrlValue == "" || (objRegExp_onlyDigits.test(vCtrlValue) == false) ){ // && ( vCtrlValue != "MPE" && vCtrlValue != "VENDING" ) )) {
                	vIsValidationPassed = "NO";
                    vMsg = "Field \""+vCtrlLable+"\" is required for the register " + p_regno + "\nand must be populated by numeric values only.";
                    break;
                }
            }else if(vCtrlName == "txt_fssp_problem"){
                txt_fssp_problem.value = vCtrlValue;
                         
                if(vCtrlValue == "" || objRegExp_onlyDigits.test(vCtrlValue) == false){
                    vIsValidationPassed = "NO";
                    vMsg = "Field \""+vCtrlLable+"\" is required for the register  " + p_regno + "\nand must be populated by numeric values only.";
                    break;
                }
            }else if(vCtrlName == "txt_fssp_zip"){
                txt_fssp_zip.value = vCtrlValue;
                             
                if(vCtrlValue == "" || objRegExp_ZipCode.test(vCtrlValue) == false || (vCtrlValue.length > 5 && vCtrlValue.length < 10)){
                    vIsValidationPassed = "NO";
                    vMsg = "Field \""+vCtrlLable+"\" is required for the register  " + p_regno + "\nand must have at least zip5."+
                                                   "\ne.g. 22084-9999 or 22084";
                    break;
                }
            }
        }                
      
        if(vIsValidationPassed == "NO"){
        	alert(vMsg);
            eval(vCtrlName+".focus();")
            eval(vCtrlName+".select();")
        }
      
        return vIsValidationPassed;
        
}} // end of with() and f_ValidateFSSP2()


//----------------------------------------------------------------------------------------------------
function f_onKeyUp(argCtrlName){with(self.document.forms[0])
{
        //
        var vStr_Length = 0;
       //
        var vStrTest = "";
        //
        var vCtrlValue = eval(argCtrlName+".value");
        vCtrlValue = f_Trim(vCtrlValue);
        //
		var objRegExp_7digits = /(^\d{7}$)/;
		var objRegExp_onlyDigits  = /(^-?\d\d*$)/;
		
        switch(argCtrlName)
        {
                //---------------------------------------------------------------
                case "txt_fssp_call":
                	if(objRegExp_onlyDigits.test(vCtrlValue) == false && vCtrlValue != ""){
						alert("The field 'FSSP Call #' field is required and is a Numeric-Only field.");
						eval(argCtrlName+".focus();")
					}
                break;
                //---------------------------------------------------------------
                case "txt_fssp_problem":
                	if(objRegExp_onlyDigits.test(vCtrlValue) == false && vCtrlValue != ""){
						alert("The field 'FSSP Problem #' field is required and is a Numeric-Only field.");
						eval(argCtrlName+".focus();")
					}
                break;
                //---------------------------------------------------------------
                case "txt_fssp_zip":
                    vStr_Length = eval(argCtrlName+".value.length;")
                    if (vStr_Length == 6) 
                    { 
                        vStrTest = vCtrlValue.substr(5,1) 
                        if (vStrTest != "-") 
                        {
                            vCtrlValue = vCtrlValue.substr(0,5)+"-"+vCtrlValue.substr(5,1);
                            txt_fssp_zip.value = vCtrlValue;
                        }
                    }         
                         
                break;
                //---------------------------------------------------------------
        } 
}}
//----------------------------------------------------------------------------------------------------

function saveMainForm(){
with(self.document.forms[0])
{
	var v_f4772 = document.getElementById('req_4772');
	document.all['wo_no_h'].value = document.all['wo_no'].value; 
	
    if (v_f4772.value == 'N' && mode == 'add') {
		
		if ((document.all['equip_no'].value.toLowerCase().indexOf('all') >= 0) || (document.all['equip_no'].value.toLowerCase().indexOf('opms') >= 0)) {
        	alert('"ALL or OPMS" is not valid as the part of Equip No.');
            document.all['equip_no'].focus();
            return;
       	}
                           
    	//02-22-05 cprice.  Added this validation for issue #156 resolution
    	if (!(validateBasicChars(document.all['equip_no'].value.trim()))) {
        	alert("Equipment number contains one or more invalid characters.  Please fix and try to save again.");
        	document.all['equip_no'].focus();
        	return;
    	}
			
    	var objchkequip = new ActiveXObject('microsoft.XMLDOM');
        objchkequip.async=false;
        objchkequip.load('em_work_pkg.em_chkequipno_prc?pSiteID='+currentSite+'&pSubsiteID='+document.all.subsite_id.value+'&pWoNo='+document.all.wo_no.value+'&pRegNo='+document.all.reg_no.value+'&pEquipNo='+document.all.equip_no.value);  
        if(objchkequip.parsed){
        	var loadequip = objchkequip.selectSingleNode('CHKDUPEQUIP');
            var chkuquip =objchkequip.text;
        }    
        if (chkuquip=='Y'){
            alert('This work order already exists, please use a different equipment number.');
            document.all.equip_no.focus();
            return;  
        }
    }
    document.all['equip_no_h'].value = document.all['equip_no'].value; 
	
	if(v_f4772.value == 'N' && document.all.equip_no.value != 'ALL'){    
        document.all['class_code_h'].value = document.all['class_code'].value;  
        if((document.all['class_code'].value == '') || (document.all['class_code'].value == -1)){
			if(document.all['equip_no'].value != 'OPMS') {  
            	alert('Entry: Class Code is required. Please complete this field, then try again.');
				document.all['class_code'].focus();
                return;
			}
        }
    }else{                    
    	//09/29/04 CM Updated so that if class selected on ALL
    	// value is stored so that work orders are only generated for equipment with that class         
    	if(document.all['equip_no'].value.toLowerCase() == "all"){
        	document.all['class_code_h'].value = document.all['class_code'].value;
    	}else{
        	document.all['class_code_h'].value = null;
    	}
	}
                        
    if ((document.all['equip_no'].value == '') || (document.all['equip_no'].value == '-1')){
    	alert('Entry: Equip Number is required. Please fill in all required fields, then save again.');
        if(document.all['equip_no'].disabled == false){ 
			document.all['equip_no'].focus();
		}
        return;
    }   

	//=========================================================================================================
	var v_regno = document.getElementById('reg_no').value;
	if(document.getElementById('txt_fssp_call').value != '' || document.getElementById('txt_fssp_problem').value != '' || document.getElementById('txt_fssp_zip').value != ''){
		
		var vIsValidationPassed = f_ValidateFSSP2(v_regno);
	  	if (vIsValidationPassed == "NO") {  
			
			return; 
		}else{
			h_fssp_call.value    = txt_fssp_call.value;
            h_fssp_problem.value = txt_fssp_problem.value;
            h_fssp_zip.value     = txt_fssp_zip.value ;
		}
	}else{
		//when users say yes to FSSP call, three fields are active. but leave all three fields empty
        if(document.getElementById('txt_fssp_call').disabled == false){ 
			alert('FSSP Call fields are required. Please fill in all required fields, then save again.');
			document.getElementById('txt_fssp_call').focus();
			
			return;
		}
	}                 
    //==========================================================================================================						   
    if ((document.all['priority_code'].value == 'A')||(document.all['priority_code'].value == 'B')||(document.all['priority_code'].value == 'C')){
        if (document.all['date_due'].value==''){
        	alert('Entry: Date Due is required. Please complete this field, then try again.');
            document.all['date_due'].focus();
            return;
        }    
    }
	// Form validation
	if (document.all['crew_id'].options[document.all.crew_id.selectedIndex].text.search(/inactive/i)!= -1) {
        alert('Cannot add corrective work to an inactive crew.');
        document.all.crew_id.focus();
        return;
    }
	
	/*jesse 2015.9.13 for registers x20/x50, 
		if 
			1) standing status = Y
			2) work code = 07 or 16 
			3) f4772 status = N 
		then 
			facility could be null (field whep_fac_id may be empty)
	*/  
	var isFacSelected = 0; 
	var v_workcode = document.getElementById('work_code').value;
	if(currentFY*1 == '2015'){
		if(v_regno.substring(0,1)*1 <= 5){ //existing registers
			if(document.getElementById('whep_fac_id').value == '-1'){
				isFacSelected = 0;
			}else{
				isFacSelected = 1;
			}
		}else{//register 6xx
			if(document.getElementById('whep_fac_id').value == '-1'){
				if(v_regno.substring(1,3) == '20' || v_regno.substring(1,3) == '50'){ 
					if(v_f4772.value == 'N' && document.getElementById('standing_flag').value == 'Y' && (v_workcode == '07' || v_workcode == '16') ){
						isFacSelected = 1;
					}else{
						isFacSelected = 0;	
					}
				}else{
					isFacSelected = 0;	
				}
			}else{
				isFacSelected = 1;	
			}		
		}
	}else{ //other future fiscal years
		if(document.getElementById('whep_fac_id').value == '-1'){
			if(v_regno.substring(1,3) == '20' || v_regno.substring(1,3) == '50'){ 
				if(v_f4772.value == 'N' && document.getElementById('standing_flag').value == 'Y' && (v_workcode == '07' || v_workcode == '16') ){
					isFacSelected = 1;
				}else{
					isFacSelected = 0;	
				}
			}else{
				isFacSelected = 0;	
			}
		}else{
			isFacSelected = 1;	
		}

	}
		
	if(isFacSelected == 0){
		alert('The Facility is required, please select one.');
		document.getElementById('whep_fac_id').focus();
			
		return;
	}
	
	//Jesse 2014.9.17 the date-input icon will go around the date_issued validation function datevalidated()
	if(mode == 'add') { datevalidate(); }    
                           
    if(document.all['project_id'].value == "-1"){
        document.all['project_id'].value = '';
    }
                         
	//validation added by Michael E. Harris on 12/02/2011 for 'past_due_reason_code' field 
	//this is a required field for all past due work orders						
	if (f_isPastDue() && mode != 'add') {
		if (document.all['past_due_reason_code'].value == '' || document.all['past_due_reason_code'].value == null) {
			alert('The "past due reason" field is a required field.  It must be selected for all past due work orders.  Please select a reason and try again.');
			return;
		}
	}							
						
    objchkequip=null;   

    if (mode != 'edit') {
        if (checkdup()) {
            saveForm(); 
        }   
    }else{
        saveForm();
		if (f_isPastDue()){ // Michael E. Harris
			v_past_due_saved = true;
		}
    }                        
}   // end of with()
	
	//2015.8.24 move th efocus to btn Exit
	document.getElementById('btnClose').focus(); 
	
}//end of function()

/***********************************************************
* Function Name: f_isPastDue							*
* Author: Michael E. Harris							*
* Purpose: This function closes the currentchecks to see 		*
* if the current Work Order is  past due.  If it is past due, it returns *
* true.  If not it returns false.    	 					*
* Creation Date: 12/6/2011							*
************************************************************/
function  f_isPastDue() {

	//alert('url=emapp.em_work_pkg.em_isWOPastDue?pSiteID='+v_currentSite+'&pWoPtr='+document.all['wo_ptr'].value);
	
	var objIsPastDue = new ActiveXObject('microsoft.XMLDOM');
	objIsPastDue.async=false;
	var urlStr = getServerPath()+'emapp.em_work_pkg.em_isWOPastDue?pSiteID='+v_currentSite+'&pWoPtr='+document.all['wo_ptr'].value;
	objIsPastDue.load(urlStr);
	if (objIsPastDue.parsed) {
		var nodelist = objIsPastDue.selectSingleNode('ISPASTDUE');
		var v_isPastDue = objIsPastDue.text;
								//alert('v_isPastDue='+v_isPastDue);
		if (v_isPastDue == 'Y') {
			return true;
		}
		else {
			return false;
		}
	}
}

/***********************************************************
* Function Name: bodyUnLoad							*
* Author: Michael E. Harris							*
* Purpose: This function closes the current work order window.	*
* Additionally, it checks to see if the current work order is past due	*
* If it is past due and the reason is left blank, it forces the user to	*
* select a reason of the work order being past due and save it		*
* before closing of the screen is allowed.					*
* Creation Date: 12/06/2011 							*
************************************************************/
function bodyOnUnLoad() {
	//Michael E. Harris 
	//If the current area is the work order area, test the past due status of the work order.
	//If the work order is past due, then the past_ due_reason_code is required and MUST be selected, populated and saved to em_wo_t.
	if (lower(document.all['AREAMain'].style.display) == 'block') {
		if (f_isPastDue() && !v_past_due_saved) {
			alert('The "past due reason" field is a required field.'+
			'  It must be selected for all past due work orders.  '+
			'  Please select a reason and save the work order.');
			return;
		}
	}
	self.close();
}

/***********************************************************
* Function Name: lower								*
* Author: Michael E. Harris							*
* Purpose: This function converts the passed string to 			*
*		lower case.								*
* Creation Date: 9/19/2006							*
************************************************************/
function lower(pString){
	var lString = pString;
	return lString.toLowerCase();
}

function checkdup() {
  var regno = document.all['reg_no'].value;
  var wono = document.all['wo_no'].value;
  var ecp = document.all['acronym'].value;
  var en = document.all['equip_no'].value; 
  
  var objchk = new ActiveXObject('microsoft.XMLDOM');
  objchk.async = false;
  objchk.load('em_work_pkg.em_chk_dupwo_prc?pSiteID='+currentSite+'&pRegNo='+regno+'&pWONO='+wono+'&pECP='+ecp+'&pEN='+en);
  if(objchk.parsed) {
   var nodelist = objchk.selectSingleNode('DUPWORK');
   var dup = objchk.text;  
  }
 
  if (dup == '4') {
    alert('This record already exists. Select a different work order number.')
    document.all['reg_no'].focus;
    return false;
  } else {
    return true;
  }  
}

function deletewo()
{
   if ((document.all['reg_no'].value.substring(1,3) < 65) && (document.all['standing_flag'].value == 'Y') ) {
           var objchkdel = new ActiveXObject('microsoft.XMLDOM');
           objchkdel.async = false;
           objchkdel.load("em_work_pkg.em_woexist?psiteid="+currentSite+"&p_wo_ptr="+document.all['wo_ptr'].value);
                   if (objchkdel.parsed) {      
                     var ckh=objchkdel.selectSingleNode('WOCOUNT');
                     var wono = objchkdel.text;
                   }  
   }
   
   if ((document.all['standing_flag'].value =='Y')&&(document.all['reg_no'].value.substring(1,3) < 65) && ((document.all['hours_to_date'].value=='')||(document.all['hours_to_date'].value==0))&&(wono != '1')){  
     var vMsg="Are you sure you wish to delete the selected record?";  

     vDeleteURL=""; 
     vDeleteURL="!vm_main_pkg.save?module=EM_WORK_MAINT_001&pSiteID="+currentSite+"+&puser="+currentUser+"&mode=delete&tab=ds&p_wo_ptr="+document.all['wo_ptr'].value;
     var myConfirm=newConfirm('Confirmation Request','Do you want to delete this record?',1,1,0);
     if (myConfirm) {
        mode='delete';
        saveForm('ds');
        disableField('btnDelete');
        disableField('btnSave');
        disableField('btnMaterial');
        document.all['btnMaterial'].style.display="none";
        document.all['btnMaterial'].style.visibility="hidden";
        disableField('btnReport'); 
		disableField('btnMTSC'); 
        disableField('btnInfo'); 
        disableField('btnMaterialsAdd');  

     } else { return; } 
     document.all['capDiv'].innerHTML='';
     clearFindScreen();
  }else{
     alert( 'Only standing work orders with a register number of less than 65 and no time or material charged against it may be deleted.');
     return false;
  }
}

//priority_code_onChange(), A,B,C, or D
function prioritychng(){ 
	var regno = document.all.reg_no.value; 
	//One day contains 86,400,000 millisecond.
	var nextdateobj = new Date(Date.parse(last_day_of_current_year)+86400000);
	//var first_day_of_next_year = '0'+(nextdateobj.getMonth()+1) +'/'+nextdateobj.getDate()+'/'+nextdateobj.getFullYear();
	var lastdateobj = new Date(Date.parse(last_day_of_current_year)+86400000*364);
	var last_day_of_next_year = '0'+(lastdateobj.getMonth()+1) +'/'+lastdateobj.getDate()+'/'+lastdateobj.getFullYear();
	
	if(document.all['priority_code'].value == ''){
		alert('Please select a Priority.');
		return;
	}

	if(mode == 'add'){
		//validate the day_issued first
		datevalidate();
	}else{
		//set up the due date based on the date issued
		if (document.all['priority_code'].value=='A'){
			document.all['date_due'].value = todaydate1(document.all.date_issued.value,1);
			disableField('date_due'); 
		}
		if (document.all['priority_code'].value=='B'){
			document.all['date_due'].value = todaydate1(document.all.date_issued.value,7);
			disableField('date_due');
		}
		if (document.all['priority_code'].value=='C'){
			document.all['date_due'].value = todaydate1(document.all.date_issued.value,30);
			disableField('date_due');
		}	
		if (document.all['priority_code'].value=='D')  {
			if(regno.substring(0,1) != currentFY.substring(3) ){
				document.all['date_due'].value = last_day_of_next_year;
			}else{
				document.all['date_due'].value = last_day_of_current_year;
			}		   
			enableField('date_due');
		} 
	}
	
	if(Date.parse(document.all['date_due'].value) > Date.parse(last_day_of_current_year)){
		if(regno.substring(0,1) == currentFY.substring(3) ){
			document.all['date_due'].value = last_day_of_current_year;
		}
	}	
}

//jesse validation on issued date, due date, and priority code 
function datevalidate(){   
	
      var dateissued = document.mainform.date_issued.value;
      var datedue = document.mainform.date_due.value;
      var today = new Date();
	  var priority = document.mainform.priority_code.value;
	  var regno = document.mainform.reg_no.value;
	  
	  //One day contains 86,400,000 millisecond.
	  var nextdateobj = new Date(Date.parse(last_day_of_current_year)+86400000);
	  var first_day_of_next_year = (nextdateobj.getMonth()+1) +'/'+nextdateobj.getDate()+'/'+nextdateobj.getFullYear();
	  if(nextdateobj.getMonth()*1 >= 9  ) { // no need the prefixed '0'
		  first_day_of_next_year = (nextdateobj.getMonth()+1) +'/'+nextdateobj.getDate()+'/'+nextdateobj.getFullYear();
	  }
	  
	//take care of the date issued first 
	if(regno == '-1') {
		alert('Please select a register number.');
		document.mainform.reg_no.focus();
		return;
	}else{	 
		if(regno.substring(0,1) != currentFY.substring(3) ){
			//it is in the last period of the fiscal year and a next year's register is selected 	
			if(Date.parse(dateissued) < Date.parse(first_day_of_next_year)){
				alert('Date Issued must be in next fiscal year, which starts from ' + first_day_of_next_year);
				document.mainform.date_issued.value = first_day_of_next_year;
			}		
		}else{
			if(Date.parse(dateissued) > Date.parse(last_day_of_current_year)){
				alert('Date Issued must be in current fiscal year, which ends in ' + last_day_of_current_year);
				document.mainform.date_issued.value = currentDate;
				document.mainform.date_issued.focus();
				//return;       
			} 	
		}	
	}
	  
	//could be 5 days past
    if(Date.parse(dateissued) < Date.parse(currentDate) - 5*86400000){
			alert('Date Issued cannot be beyond more than 5 days past from today.');
			document.mainform.date_issued.value = currentDate;			
			//return;	
    }
	
    if(Date.parse(dateissued) > Date.parse(datedue)){
        alert('Date Issued cannot be greater then Date Due.');
		if(regno.substring(0,1) == currentFY.substring(3) ){
			document.mainform.date_issued.value = currentDate;   
		}else{
			document.mainform.date_issued.value = first_day_of_next_year;
		}
        //return;       
    } 
		
	//set up the due date based on the date issued
	if (document.all['priority_code'].value=='A'){
         document.all['date_due'].value = todaydate1(document.all.date_issued.value,1);
    }
    if (document.all['priority_code'].value=='B'){
         document.all['date_due'].value = todaydate1(document.all.date_issued.value,7);
    }
    if (document.all['priority_code'].value=='C'){
         document.all['date_due'].value = todaydate1(document.all.date_issued.value,30);
    }
    if (document.all['priority_code'].value=='D')  {
		if(regno.substring(0,1) != currentFY.substring(3) ){
				document.all['date_due'].value = last_day_of_next_year;
		}else{
				document.all['date_due'].value = last_day_of_current_year;
		}		    
		enableField('date_due');
    }
	
	if(Date.parse(document.all['date_due'].value) > Date.parse(last_day_of_current_year)){
		if(regno.substring(0,1) == currentFY.substring(3) ){
			document.all['date_due'].value = last_day_of_current_year;
		}
	}else{
		if(regno.substring(0,1) != currentFY.substring(3) ){
			document.all['date_due'].value = last_day_of_next_year;
		}
	}
}


function localonBlur(control){

   if (control=='reg_no' && mode=='add') {

	   
	   // 2020.12.04 move this block from onREGChange() due to the selecting by keyboard 
	  	//jesse 2015.1.29
		var v_regno = document.getElementById('reg_no').value;
		var v_regno_text = document.getElementById('reg_no').options[document.getElementById('reg_no').options.selectedIndex].text;
		
		//alert('v_regno='+v_regno);alert('v_regno_text='+ v_regno_text);
		
		var fssp_enable = 'N';
		var reg_no23 = v_regno.substring(1,3);
		var objAcros=new ActiveXObject('microsoft.XMLDOM');              
   		objAcros.async=false; 
   		objAcros.load("em_work_pkg.em_get_fssp_status?pSiteID="+currentSite+"&pRegno="+document.all['reg_no'].value);
		if (objAcros.parsed) {       
      		var whatever = objAcros.selectSingleNode('fssp_enable');
			fssp_enable = objAcros.text;
		}else{
			fssp_enable = 'N';	
		}
		
	  	if(fssp_enable == 'Y' || v_regno_text.indexOf("FMO",0) != -1) {
	  		if(cust_confirm('FSSP Call','Does this work order require any FSSP calls?',3,4,1,0) != 6) {
				f_set_EmptStr_Into_FSSP();
				disableField('txt_fssp_call');
    			disableField('txt_fssp_problem');
    			disableField('txt_fssp_zip');
			}else{
				enableField('txt_fssp_call');
				enableField('txt_fssp_problem');
				enableField('txt_fssp_zip');
			}
		}else{
			f_set_EmptStr_Into_FSSP();
			disableField('txt_fssp_call');
    		disableField('txt_fssp_problem');
    		disableField('txt_fssp_zip');
		}
	   //2020.12.04

    	return;// 2020.11.19 onREGChange(); 
   }
   
   
   if ( control =='acronym' && mode == 'add') { 
   			if(currentSite == '-303'){ alert('within localblur() - control ==acronym && mode == add'); };
		
	    if (document.all['equip_no'].options.length == 1 && document.all['acronym'].selectedIndex != -1) {
              alert('There are no equipment for this acronym at subsite '+document.all['subsite_id'].value+'.');
			  disableField('equip_no');
			  
			  document.all['acronym'].selectedIndex = -1;
              document.all['acronym'].focus(); 
			  return;
		}
		
		if(ckhopms != 'N') {//the value of "ckhopms" is assigned within acroChanged().
			allopms();
		}
		
		enableField('equip_no');
		document.all['equip_no'].focus(); 
   }

   if (control=='standing_flag' && mode=='add'){

   }
   
   //2020.12.02 move this reminding message from function acroChange() 
   if ( control =='work_code' && mode == 'add') { 
   	
   		if (document.all['equip_no'].options.length==1) {
              //alert('There are no equipment for this acronym at subsite '+document.all['subsite_id'].value+'.');
			  alert('Please ensure the fields of Acronym and Equip No are properly selected.');
              document.all['acronym'].focus();
              return;
    	}
   }//2020.12.02
   
  if (control=='date_issued') { 
      if(mode=='add') { datevalidate(); }
   }
   
  if (control=='wo_no' && mode=='add') {
     var wo = document.all['wo_no'].value;
     var regno = document.all['reg_no'].value;
     var objwono=new ActiveXObject('microsoft.XMLDOM');
     objwono.async=false;
     objwono.load("em_work_pkg.em_wono9?pSiteID="+currentSite+"&p_wo="+wo+"&p_regno="+regno);
     if (objwono.parsed) {
       var result=objwono.selectSingleNode('WONO9');
       var wono = objwono.text;
     }  
	  
     if (wono !=''){           
       document.all['btnFind'].style.display="none";
       var myConfirm=newConfirm('Confirmation Request','This work order number already exists do you want to add another piece of equipment?',1,1,0);
       if (myConfirm) {
       	  mode = 'add';
          document.all['AREAMain'].style.display="block";
          document.all['BlankAREA'].style.display="none";
          enableField('btnSave');
          enableField('btnDelete');
          //enableField('btnDelete');
          disableField('wo_desc');
          disableField('subsite_id');
          disableField('standing_flag');
          disableField('work_code');      
          disableField('acronym');        
          disableField('hours_to_date');
		  
          var objFind=new ActiveXObject('microsoft.XMLDOM');
          objFind.async=false; 
          objFind.load("em_work_pkg.em_getWO?p_rowid="+URLEncode(wono)+"&pSiteID="+currentSite); 
          var nodeList = objFind.getElementsByTagName('ROW');
          if (nodeList.length==0) { alert('No matching record found in the database'); return; }  
		  
          for (i=0; i < nodeList(0).childNodes.length; i++) {
              try {      
                  document.all[nodeList(0).childNodes(i).nodeName.toLowerCase()].value=null;
                  document.all[nodeList(0).childNodes(i).nodeName.toLowerCase()].value=nodeList(0).childNodes(i).text;     

              } catch(e) { }
          }
          objFind=null;
		  
          summdisplay();    
           
          document.all['equip_no'].value='';
          document.all['equip_no_h'].value = '';
          document.all['dept_code'].value='';            
          document.all['est_hours'].value='';            
          document.all['wo_comment'].value='';
		  
          var reg_no =document.all['reg_no'].value;
          var Acro = document.all['acronym'].value;
		  
          // check if it is 4772 required. tttt
          if (document.all['req_4772'].value=='N') {        
             document.all['divEquipNo'].innerHTML='<font color=black>*  </font><b>  Equip No</b><BR><input type="text" widget name="equip_no" id="equip_no"  onKeyPress="enterequip()" TABINDEX="1" size="6" />';
             alert('Enter equipment number.');
             document.all.equip_no.focus();
          } else { 
             document.all.class_ptr.value='';  // what is this for?  why????? 
             document.all['divClass'].innerHTML='<font color=black></font><b>Class</b><BR><input type ="text" name="class_code" id="class_code" TABINDEX="1" style="width: 90">';         
             disableField('class_code'); 
             document.all['divEquipNo'].innerHTML='<font color=black>*  </font><b>  Equip No</b><BR><select widget name="equip_no" id="equip_no"  TABINDEX="1" style="width: 90" onKeyPress="typeAhead(); enterequip();" onChange="equipChange()"></select>';           
             document.all['req_4772'].value='Y';
              
             loadList('equip_no',"em_work_pkg.em_XMLEquipList_minus?psiteid="+currentSite+"&p_subsite_id="+document.all['subsite_id'].value+"&p_Acro="+Acro+"&p_equip_status=A"+"&pRegNo="+document.all['reg_no'].value+"&pWONO="+document.all['wo_no'].value,'Y'); 
             if (document.all['equip_no'].options.length < 1) {
             	alert('There are no equipment for this acronym at subsite '+document.all['subsite_id'].value+'.');
                disableField('equip_no');
                document.all['acronym'].selectedIndex=-1;
                document.all['acronym'].focus();
                return;
             } 
         } 
           
       } else {
          document.all['wo_no'].value='';
          document.all['wo_no'].focus();
          return;
       }
     }
   }
}

function enterequip() {
  var thisfield=window.event.srcElement;
  if(window.event && window.event.keyCode == 13) {
     if (document.all['req_4772'].value=='N') {
       if(isWidgetVisible('class_code')) {
        	document.all.class_code.focus();
       }else {
       		if (document.all['work_code'].value !=25) {
        		document.all['work_code'].focus(); //Tehseen:2/15/06: Trick typeahead
       		} 
       		 document.all.priority_code.focus();
       }       
     }else {
       if (document.all['work_code'].value !=25) {
        	document.all['work_code'].focus(); //Tehseen:2/15/06: Trick typeahead
       } 
        document.all.priority_code.focus();
     }   
  }
}

function yynnn(fieldnum) {
            var thisfield=window.event.srcElement;
            if(window.event && window.event.keyCode == 13) {
                  for(x=0;x<mainform.length;x++){
                if(mainform.elements[x].name=="fieldnum"){
                     if (mainform.elements[x].value.length <1) {
                       mainform.elements[x].focus();

                       return;
                     }
                }
              }
            }
}

function enterwono() {
 if(window.event && window.event.keyCode==13) {
   document.all.wo_desc.focus();
 }
}

function enterclass() {

 if(window.event && window.event.keyCode==13) {
//   document.all.priority_code.focus();
 }
}

function equipChange() {
     var objP = new ActiveXObject('microsoft.XMLDOM');
     objP.async=false;
     selAcro=document.all['acronym'].options[document.all['acronym'].selectedIndex].value;
     objP.load('em_work_pkg.em_get_class?pSiteID='+currentSite+'&pEquipPtr='+document.all['equip_no'].value);
     if(objP.parsed) {
       var loadlist = objP.selectSingleNode('CLASSCODE');
       var pCC = objP.text;
     }

//09/29/04 CM Updated so that if equip_no ALL selected for reg>70
// a drop down list will show with class codes to select from
    if(document.all['reg_no'].value.substring(1,3) > 70){
        if(document.all.equip_no.value=='ALL'){
          document.all['divClass'].innerHTML='<font color=black></font><b>Class</b><BR><select widget name="class_code" id="class_code" TABINDEX="1" onKeyPress="enterclass()" style="width: 90"></select>';
          loadList('class_code', getServerPath()+"em_work_pkg.em_XMLClassList_acro_equip?pSiteID="+currentSite+"&p_acro="+selAcro+"&p_subsite_id="+document.all['subsite_id'].value,'Y');
        } else {
          document.all['divClass'].innerHTML='<font color=black></font><b>Class</b><BR><input type ="text" name="class_code" id="class_code" TABINDEX="1" style="width: 90">';
          document.all['class_code'].value=pCC;
          disableField('class_code');
        }
     } else {
        document.all['class_code'].value=pCC;
        disableField('class_code');
     }

    objP = null;
}

function fscgo1(){ 
   if (document.all['fsccode'].value ==''){
     alert('Select an NSN to search by');
     document.all['fsccode'].focus();
     return;
   }
   else{
     loadList('nsnnom', getServerPath()+'em_work_pkg.em_xmlnsn_list?psiteid='+currentSite+'&p_fsc='+document.all['fsccode'].value, 'N'); 
   }   
}

function nomengo1(){  
   if (document.all['nomen'].value ==''){
     alert('Enter a Nomenclature to search by');
     document.all['nomen'].focus();
     return;
   }
   else{
    loadList('nsnnom', getServerPath()+'em_work_pkg.em_xmlnom_list?psiteid='+currentSite+'&p_nom='+document.all['nomen'].value, 'N'); 
   } 
}

//??????
function summdisplay(){
/*
   document.all['fy_add'].value = document.all.fy.value;
   document.all['register_add'].value=document.all.reg_no.value;
   document.all['wo_no_add'].value = document.all.wo_no.value 
   document.all['wo_desc_add'].value = document.all.wo_desc.value;
   */
}

function clearScreen(){
  for (i=0; i < mainform.elements.length; i++) {     
     try {
        if (mainform.elements[i].type != 'button') {
          if(mainform.elements[i].name != 'module_from'){
            if(mainform.elements[i].name != 'stocked'){
              if(mainform.elements[i].name != 'chkstocked'){
                if(mainform.elements[i].name != 'prtrpt'){
                 mainform.elements[i].value='';
                } 
              }  
            }    
          }      
        }
     } catch(e) { }     
   }
  if (document.all['hours_to_date'].value=='') {
    document.all['hours_to_date'].value=0;
  }
   
}

function clearMaterialScreen(){
  for (i=0; i < mainform.elements.length; i++) {     
     try {
        if (mainform.elements[i].type != 'button') {
          if(mainform.elements[i].name != 'module_from'){
            if(mainform.elements[i].name != 'stocked'){
              if(mainform.elements[i].name != 'stocked'){
                if(mainform.elements[i].name != 'reg_no'){
                  if(mainform.elements[i].name != 'wo_no'){
                    if(mainform.elements[i].name != 'wo_desc'){
                      if(mainform.elements[i].name != 'wo_ptr'){
                        if(mainform.elements[i].name != 'chkstocked'){
                          if(mainform.elements[i].name != 'prtrpt'){
                            mainform.elements[i].value='';
                          } 
                        }  
                      }
                    }  
                  }
                }                
              }  
            }    
          }      
        }
     } catch(e) { }     
   }
  if (document.all['hours_to_date'].value=='') {
    document.all['hours_to_date'].value=0;
  }  
}

function clearFindScreen(){
  for (i=0; i < mainform.elements.length; i++) { 
		 try {
			if (mainform.elements[i].type != 'button') {
			  if(mainform.elements[i].name != 'module_from'){
				if(mainform.elements[i].name != 'stocked'){
				  if(mainform.elements[i].name != 'wo_desc_key'){
					if(mainform.elements[i].name != 'reg_no_key'){
					  if(mainform.elements[i].name != 'chkstocked'){
						if(mainform.elements[i].name != 'wo_no_key'){
						  if(mainform.elements[i].name != 'prtrpt'){
							mainform.elements[i].value='';
						  } 
						}   
					  }  
					}                
				  }  
				}    
			  }      
			}
		 } catch(e) { }     
   }
   
   
	  if (document.all['hours_to_date'].value=='') {
		document.all['hours_to_date'].value=0;
	  }   
}


function clearFindScreenMWO(){
  for (i=0; i < mainform.elements.length; i++) {     
     try {
        if (mainform.elements[i].type != 'button') {
          if(mainform.elements[i].name != 'module_from'){
            if(mainform.elements[i].name != 'stocked'){
              if(mainform.elements[i].name != 'stocked'){
                if(mainform.elements[i].name != 'reg_no'){
                  if(mainform.elements[i].name != 'chkstocked'){
                    if(mainform.elements[i].name != 'acronym'){
                     if(mainform.elements[i].name != 'prtrpt'){
                       mainform.elements[i].value='';
                     } 
                    }  
                  }  
                }                
              }  
            }    
          }      
        }
     } catch(e) { }     
   }
  if (document.all['hours_to_date'].value=='') {
    document.all['hours_to_date'].value=0;
  } 
} 

function clearRegScreen(){
  for (i=0; i < mainform.elements.length; i++) {     
     try {
        if (mainform.elements[i].type != 'button') {
          if(mainform.elements[i].name != 'module_from'){
            if(mainform.elements[i].name != 'stocked'){
              if(mainform.elements[i].name != 'stocked'){
                if(mainform.elements[i].name != 'reg_no'){
                  if(mainform.elements[i].name != 'chkstocked'){
                    if(mainform.elements[i].name != 'date_issued'){
                     if(mainform.elements[i].name != 'prtrpt'){
                       mainform.elements[i].value='';
                     } 
                    }   
                  }  
                }                
              }  
            }    
          }      
        }
     } catch(e) { }     
   }
  if (document.all['hours_to_date'].value=='') {
    document.all['hours_to_date'].value=0;
  }
}

var chkradio;
function stocked_onClick() {
 for (var  i=0; i< document.mainform.stocked.length; i++){
   if ( document.mainform.stocked[i].checked){
        chkradio = i;
   } 
 }

 document.all['chkstocked'].value = (document.mainform.stocked[chkradio].value); 
 
 if (document.all['chkstocked'].value =='Y') {        
    document.all['selectdiv'].style.display="block";
    document.all['nsnnomdiv'].style.display="block";
    document.all['inpdiv'].style.display="none";
    document.all['divboh'].style.display="block";
  }else{
    document.all['selectdiv'].style.display="none";
    document.all['nsnnomdiv'].style.display="none";    
    document.all['inpdiv'].style.display="block";
    document.all['divboh'].style.display="none";
  }
}

var savedwono;
var savedequipno;
var savedregno;
var savedclass;

function cust_confirm(title,mess,icon,buts,defbut,mods) {
   retVal = makeMsgBox(title,mess,icon,buts,defbut,mods);
   //retVal = (retVal==6);
   return retVal;
}

function acbmClicked(){
        if (document.all.p_acbm_ind.checked==true) {
           document.all.work_code.value=25;
           disableField('work_code');
           document.all['p_acbm_lbl'].innerHTML='<font color=BLUE> PS Form 8210 must be approved prior to creating <br />any ACBM work orders.To access form Click </font> <input type="button" value="PS Form 8210" onClick="ps8210()">';                       
        } else {
           document.all.work_code.value="";
           document.all['p_acbm_lbl'].innerHTML='';              
           enableField('work_code');
        }
}

function workCodeChange() 
{
    if (document.all.work_code.value==25) {
        document.all.p_acbm_ind.checked=true;    
        document.all['p_acbm_lbl'].innerHTML='<font color=BLUE> PS Form 8210 must be approved prior to creating <br />any ACBM work orders.To access form Click </font> <input type="button" value="PS Form 8210" onClick="ps8210()">';                       
    } else {
        document.all.p_acbm_ind.checked=false;
        document.all['p_acbm_lbl'].innerHTML='';                        
    }
    
    //typeAhead();// this was on the swil data entry onChange
	
	//jesse 2014.9.22
	var reg_no = document.all['reg_no'].value;
	if(currentFY == '2014') { 
		if(reg_no.substring(0,1) == '5'){
			if (reg_no.substring(1,3) == '09'){ 
				if(document.all['work_code'].value != '09'){
					alert('The work code must be 09.');
					document.all['work_code'].value = '09';
				}
			}
			if (reg_no.substring(1,3) == '24'){ 
				if(document.all['work_code'].value != '24'){
					alert('The work code must be 24.');
					document.all['work_code'].value = '24';
				}
			}
		}
	}else{
			if (reg_no.substring(1,3) == '09'){ 
				if(document.all['work_code'].value != '09'){
					alert('The work code must be 09.');
					document.all['work_code'].value = '09';
				}
			}
			//if (reg_no.substring(1,3) == '24'){ // jesse 2015.9.13 register x14 replaces x24 in fy 2016
			if (reg_no.substring(1,3) == '14'){ 
				if(document.all['work_code'].value != '24'){
					alert('The work code must be 24.');
					document.all['work_code'].value = '24';
				}
			}
	}

	var v_workcode = document.getElementById('work_code').value;
	var v_prioritycode = document.getElementById('priority_code').value; 
	
	//2020.1.15 jesse only 4772-req acronyms are allowed for 08/23 work orders 
	if(v_workcode == '23' || v_workcode == '08' ) { 
		if(ckhopms == 'N') { 
			alert('The work orders with WC 08 or 23 are not allowed for this acronym.');
			document.all['work_code'].selectedIndex = -1;
			document.all['acronym'].selectedIndex = -1;
			document.all['acronym'].focus();
			
			formOnAdd();
			return;
		}
	}
	//2020.1.15 jesse only 4772-req acronyms are allowed for 08/23 work orders 
	
	//2015.8.17, per mark the priority of work codes 05/08/21/23 must be "A"
	//2015.8.25, per mark the priority of work codes 05/08/18/21/23 must be "A"
	if(v_workcode == '05' || v_workcode == '08' || v_workcode == '18' || v_workcode == '21' || v_workcode == '23'){
		//only allowed for non-standing work orders 
		if(v_prioritycode != 'A'){
			alert('The priority code has to be "A" for the work code ' + v_workcode);
			document.getElementById('priority_code').value = 'A';
			prioritychng();
			disableField('priority_code');	
		}
	}else{
		if(document.getElementById('standing_flag').value == 'N'){
			enableField('priority_code');
		}else{
			; //do nothing, priority is already set to D and field is disabled 	
		}
	}
} 

//Jesse == btnAdd_onClick() 
function formOnAdd() {   
    document.all['capDiv'].innerHTML = "";
    disableField('acronym'); 
    document.all['BlankAREA'].style.display = "none";
    enableField('reg_no');
    enableField('standing_flag');
	disableField('date_due');
	disableField('date_issued'); //jesse 2014.9.17
	
    enableField('crew_id');
    enableField('project_id');    
    enableField('dept_code');
    enableField('wo_desc');
    enableField('p_acbm_ind');
    enableField('work_code');
    enableField('est_hours');
    enableField('wo_comment');
	disableField('past_due_reason_code');
       
	document.all.AREAMat.style.display = "none";
	
  	if(mainform.elements['btnAdd'].value == 'ADD') {
  		mode='add';
		
    	savedwono = document.all['wo_no'].value;
    	savedclass = document.all.class_code.value;
    	document.all['partDiv'].innerHTML = "";
    	savedequipno = document.all['equip_no'].value;
    	savedregno = document.all['reg_no'].value;
		
    	saveElems();
    	clearFindScreen();
		
		//jesse 2015-6-15,  to use the local date as date_issued 
		currentLocalDate = currentDate;
		var objLocalDate = new ActiveXObject('microsoft.XMLDOM');
    	objLocalDate.async=false;
    	objLocalDate.load('em_time_pkg.em_get_local_date?p_site_id='+currentSite);
    	if(objLocalDate.parsed) {
        	var loadlist = objLocalDate.selectSingleNode('local_date');
        	currentLocalDate = objLocalDate.text;
    	}		
    	document.all['date_issued'].value = currentLocalDate;
		currentDate = currentLocalDate;
		//jesse 2015-6-15
		
    	//2015.8.24 standing_flag stays whatever it was. 
		//document.all['standing_flag'].value = 'N';
		
    	document.all['wo_status'].value = 'O';	
    	document.all['btnFind'].style.display = "none";
    	document.all['AREAMain'].style.display = "block";
		
    	disableField('btnMaterial');
		disableField('btnMaterialsAdd');
    	disableField('btnReport'); 
    	disableField('btnMTSC');  	 
    	disableField('btnInfo'); 
    	enableField('btnSave'); 
    	disableField('btnDelete'); 
		disableField('past_due_reason_code');
		
    	mainform.elements['btnAdd'].value = 'CANCEL ADD';
    	document.all['capDiv'].innerHTML = "Add Work Order";
    	document.all['reg_no'].focus();
		
    	loadList('reg_no', getServerPath()+'em_work_pkg.em_getXMLRegisterList?psiteid='+currentSite, 'Y');       
    	listallregno = 'N';

        var cMsg = 'Does this work involve any Asbestos Containing Building Material (ACBM)?';
        var cttl = 'Asbestos Containing Building Material (ACBM)';
        var cIcon = 3;
        var cButs = 4;
        var cDef = 1;
        var cMods = 0;
		var acbmwork = cust_confirm(cttl,cMsg,cIcon,cButs,cDef,cMods);
        if(acbmwork == 6){                     // yes, 
           document.all['work_code'].value = '25';
           disableField('work_code');
           document.mainform.p_acbm_ind.checked=true;
           document.all['p_acbm_lbl'].innerHTML='<font color=BLUE> PS Form 8210 must be approved prior to creating '
		   		+ '<br />any ACBM work orders.To access form Click </font> <input type="button" value="PS Form 8210" onClick="ps8210()">';                         
        }else if(acbmwork == 7){ //no
           document.mainform.p_acbm_ind.checked = false;
           document.all['p_acbm_lbl'].innerHTML = '';              
           //enableField('work_code');
        }
		document.mainform.p_acbm_ind.disabled = true; //jesse, checked/unchecked will changes the settings of other fields
      	 
		//jesse 2015-1-29
		disableField('txt_fssp_call');
    	disableField('txt_fssp_problem');
    	disableField('txt_fssp_zip');
	
		document.all['reg_no'].focus();
    	disableField('acronym');
		
	    if(currentSite == '-303'){ alert('line 1785 after ADD is pressed'); }; 
      
  	}else{ //who cares about the following codes if this function is called only after the ADD button is pressed??? 
			// is it called after "SAVE" is done ? maybe :)))))) 
		
		mode='edit';
		mainform.elements['btnAdd'].value = 'ADD';
        document.all['btnFind'].style.display = "block";   
        document.all['AREAMain'].style.display = "block";   
		
		/* 2020.11.18 jesse 
			loadList('reg_no', getServerPath()+'em_work_pkg.em_getXMLRegisterListAll?psiteid='+currentSite, 'Y');
			loadList('subsite_id', getServerPath()+"em_work_pkg.em_regsubsite_prc?pSiteID="+currentSite+"&pRegNo="+savedregno, 'N');
			document.all['capDiv'].innerHTML = "Work Order Details";
			disableField('standing_flag');
			disableField('date_due');
			disableField('crew_id');

			disableField('dept_code');
			disableField('hours_to_date');
			disableField('wo_no');
			disableField('subsite_id');
			disableField('acronym');
			disableField('equip_no');
			disableField('reg_no');
			disableField('wo_desc');
			disableField('priority_code');
			disableField('work_code');
			disableField('date_issued');
			disableField('est_hours');
			disableField('wo_comment');
			disableField('project_id');
			disableField('p_acbm_ind');
			document.all['wo_no'].value = savedwono;
			document.all['class_code'].value = savedclass;
			document.all['equip_no'].value = savedequipno;
			
			var objIsPastDue = new ActiveXObject('microsoft.XMLDOM');
			objIsPastDue.async = false;
			var urlStr = getServerPath() + 'emapp.em_work_pkg.em_isWOPastDue?pSiteId=' + v_currentSite + '&pWoPtr='+document.all['wo_ptr'].value;
			objIsPastDue.load(urlStr);
			if(objIsPastDue.parsed){
				var nodelist = objIsPastDue.selectSingleNode('ISPASTDUE');
				var v_isPastDue = objIsPastDue.text;
				if (v_isPastDue == 'Y'){
					enableField('past_due_reason_code');
				}else{
					disableField('past_due_reason_code');
				}
			}
			objIsPastDue = null; //jesse
      
			onChange_RegisterNo();
                  
            var objNodeList = objDoc.getElementsByTagName('widget');
            var nodename;
            for (i = 0; i < objNodeList.length; i++) {  
              try{           
                nodename = objNodeList(i).getAttribute('name');
                nodevalue = objNodeList(i).getAttribute('value'); 
                if (objNodeList(i).getAttribute('type') != 'button'){ 
                  	document.all[nodename].value = nodevalue;
                }
              }catch(e){ }  
            }     

            var objP = new ActiveXObject('microsoft.XMLDOM');
            objP.async = false;
            objP.load('em_work_pkg.em_check_parts?pSiteID=' + currentSite+'&pWOPtr=' + document.all['wo_ptr'].value);
            if(objP.parsed){
				var loadlist = objP.selectSingleNode('PARTS');
				var pExist = objP.text;
            }
            if(pExist == '4') {
				document.all['partDiv'].innerHTML = "-----Parts";     
            }else{
				document.all['partDiv'].innerHTML = "";
            }
            objP = null;

            if (document.all['wo_ptr'].value == ''){
                disableField('btnSave'); 
                disableField('btnDelete'); 
                disableField('btnReport'); 
				disableField('btnMTSC'); 
                disableField('btnMaterialsAdd'); 
                disableField('btnInfo'); 
                disableField('btnMaterial'); 
            }else{
                enableField('btnSave'); 
                enableField('btnDelete'); 
                enableField('btnReport'); 
                enableField('btnMTSC'); 
                enableField('btnMaterialsAdd'); 
                enableField('btnInfo'); 
                enableField('btnMaterial'); 
           }
	*/	   //2020.11.18 jesse 
	
			
		//2020.11.19 
		clearFindScreen(); 
		disableField('btnSave'); 
		disableField('reg_no'); 
    }
		  
	//jesse
	disableField('priority_code');
	disableField('work_code');
	//jesse

}

function closeme(reportID){
 	var myConfirm=newConfirm('Confirmation Request','Do you want to print all transactions?',1,1,0);
	
 	if (myConfirm) {
	  if (document.all['prtrpt'].value !=''){
	   		runReport(reportID);
	  }else {
			document.all['prtrpt'].value ='-1';  
	  }
 	} 
  	self.close();
}

function printwoReport(reportID){
  if (document.all['prtrpt'].value =='') {
    	document.all['prtrpt'].value ='-1';  
  }
  runReport(reportID);
  
  if (document.all['prtrpt'].value =='-1') {
    	document.all['prtrpt'].value ='';  
  }
}
    
function save(pTab, saveStatus, newValue){
   if (document.all['prtrpt'].value ==''){
     document.all['prtrpt'].value = document.all['wo_ptr'].value;
   }else{
      document.all['prtrpt'].value = document.all['prtrpt'].value +','+ document.all['wo_ptr'].value;
   } 
   
  if ((pTab != "Material")&&(mode=='add'))
  {
		document.all['divwono'].innerHTML='<font color=black>*</font><b>WO NO</b><BR><input type="text" name="wo_no" id="wo_no" size="6" onKeyPress="enterwono()" TABINDEX="1" onblur="localonBlur(this.name)" />';     
		document.all['divEquipNo'].innerHTML='<font color=black>*  </font><b>  Equip No</b><BR><input type="text" widget name="equip_no" TABINDEX="1" id="equip_no" size="6" />';
		document.all['divClass'].innerHTML='<font color=black></font><b>Class</b><BR><input type="text" name="class_code" TABINDEX="1" id="class_code" style="width: 90">';         
		disableField('class_code');
		document.all['req_4772'].value='N';
		var objFind=new ActiveXObject('microsoft.XMLDOM');
		objFind.async=false; 
		objFind.load("em_work_pkg.em_getWO?p_rowid="+URLEncode(newValue)+"&pSiteID="+currentSite); 
		var nodeList = objFind.getElementsByTagName('ROW');
		for (i=0; i < nodeList(0).childNodes.length; i++) {
		  try {      
			  document.all[nodeList(0).childNodes(i).nodeName.toLowerCase()].value=null;
			  document.all[nodeList(0).childNodes(i).nodeName.toLowerCase()].value=nodeList(0).childNodes(i).text;     
					 
		  } catch(e) { }
		}
    	objFind=null;
   }
   
    document.all['btnMaterial'].style.visibility="visible";
    document.all['btnMaterial'].style.display="block";
    enableField('btnMaterial');
    enableField('btnMaterialsAdd'); 
    enableField('btnRegistersAdd'); 
    enableField('btnReport'); 
    enableField('btnMTSC'); 
    enableField('btnInfo'); 
 
    summdisplay();    
    mainform.elements['btnAdd'].value='ADD';
    document.all['btnFind'].style.display="block";
    mode='edit';
    document.all['AREAMain'].style.display="block";
    enableField('btnSave');
    enableField('btnDelete');
    disableField('reg_no');
    disableField('subsite_id');
    disableField('acronym');
    disableField('equip_no');
    disableField('standing_flag');
    disableField('date_due');
    disableField('hours_to_date');
    disableField('wo_no');
    if (pTab == "Material"){
       material('SAVE');
    }   
}

function afterclear() {
  document.all['reg_no'].focus();
}

function nsnnomchange(){
  var nsnnom = document.all['nsnnom'].value;
  nsnnom = nsnnom.substring(0,16);
  var objBOH=new ActiveXObject('microsoft.XMLDOM');              
  objBOH.async=false;
 
  objBOH.load("em_work_pkg.em_boh_prc?p_nsnnom="+nsnnom+'&psiteid='+currentSite);
  if (objBOH.parsed) {  
    var result=objBOH.selectSingleNode('BOH');
    var aboh = objBOH.text;
    document.all['boh'].value=aboh;
  }
  objBOH=null;    
}

function material(p_save){
    disableField('reg_no');
    stocked_onClick();
    
    document.all['BlankAREA'].style.display="none";
    if (p_save=='SAVE'){
      lmode="work";
      document.all['AREAIssue'].style.display="none";
    }
   if (lmode=="work") {
      var ptr = document.all['wo_ptr'].value;
      var objAcros=new ActiveXObject('microsoft.XMLDOM');              
      objAcros.async=false;
      objAcros.load("em_work_pkg.em_materialcnt?p_wo_ptr="+ptr+'&psiteid='+currentSite);
      if (objAcros.parsed) {    
         var result=objAcros.selectSingleNode('MATERIALCOUNT');
         var arows = objAcros.text;
         grdMaterial.setNumRows(arows);
      }
      objAcros=null;    
      document.all.AREAMat.style.display="block";
      grdMaterial.setGridData(getServerPath()+'em_work_pkg.em_getMaterialList?p_wo_ptr='+ptr+'&psiteid='+currentSite);   
	  
      document.all['AREAMatGrid'].style.display="block";
      document.all['AREAMain'].style.display="none";
      document.all['AREAKey'].style.display="block";
      document.all['btnMaterial'].style.visibility="visible";
      document.all['btnMaterial'].style.display="block";
      document.all['btnMaterial'].value="Close Material";
      document.all['btnSave'].style.visibility="hidden";
      document.all['btnDelete'].style.visibility="hidden";
      document.all['btnAdd'].style.visibility="hidden";
      document.all['btnFind'].style.visibility="hidden";
      lmode="material";
   }
   else {
      document.all.AREAMat.style.display="none";
      document.all['AREAKey'].style.display="block";   
      document.all['AREAMatGrid'].style.display="none";
      document.all['AREAIssue'].style.display="none";
      document.all['AREAMain'].style.display="block";
      document.all['btnMaterial'].style.visibility="visible";
      document.all['btnMaterial'].style.display="block";
      document.all['btnMaterial'].value="Material";
      lmode="work"
      document.all['btnSave'].style.visibility="visible";
      document.all['btnDelete'].style.visibility="visible";
      document.all['btnAdd'].style.visibility="visible";
      document.all['btnFind'].style.visibility="visible";

      var objP = new ActiveXObject('microsoft.XMLDOM');
      objP.async=false;
      objP.load('em_work_pkg.em_check_parts?pSiteID='+currentSite+'&pWOPtr='+document.all['wo_ptr'].value);
      if(objP.parsed) {
        var loadlist = objP.selectSingleNode('PARTS');
        var pExist = objP.text;
      }
      if(pExist=='4') {
         document.all['partDiv'].innerHTML="-----Parts";     
      }
      else {
         document.all['partDiv'].innerHTML="";
      }
     objP = null;      
    }
}

function deleteGrid(){
   if (grdMaterial.getNumSelRows()==0) {
       alert('Please select the row you wish to delete, then try again');
       return; 
   }
   var vMsg="Are you sure you wish to delete the selected row?";  
   var vDeleteURL=""; 
   var vSelRow=grdMaterial.getSelRow();
   var rowid =grdMaterial.getCellText(7, vSelRow)+'';  
   vDeleteURL="!vm_main_pkg.save?module=EM_WORK_MAINT_001&pSiteID="+currentSite+"&pUser="+currentUser+"&prowval="+URLEncode(rowid)+"&mode=delete&tab=Material";
   var myConfirm=newConfirm('Confirmation Request','Are you sure you wish to delete the selected row?',1,1,0);
   if (myConfirm) {
      var objData=new ActiveXObject('microsoft.XMLDOM');              
      objData.async=false; 
      objData.load(vDeleteURL);
      var result=objData.getElementsByTagName('message');
      if (result(0).nodeValue != null) {
         alert(result(0).nodeValue);
      }
      material('SAVE');
   }  
   else { return } 
}

function addGrid(){
      removeAllOptions(document.all['nsnnom'].name);
      if(document.all['date_closed'].value.length >0) {
		  
       // jesse 2013.6.7  if ( (Date.parse(currentDate) - Date.parse(document.all['date_closed'].value)) >5   ) {
	     if ( new Date().getDate() - new Date(document.all['date_closed'].value).getDate()  > 5   ) {
          alert('Cannot add material to orders closed over 5 days old.');
          return;
        }
      }
      document.mainform.stocked[0].checked=true;
      stocked_onClick();
      editmode = 'N';
      if (document.all['standing_flag'].value=='Y') {
        alert('Materials cannot be added to a standing work order');
        return;
      }
      else {
        mode = 'add';
        enableField('nsn');
        enableField('nom');
        document.all['stockeddiv'].style.display="block";
        document.all['AREAMatGrid'].style.display="none";
        document.all['AREAIssue'].style.display="block";
        document.all['nsn'].value='';
        document.all['nom'].value='';
        document.all['nomen'].value='';
        document.all['qty_req'].value='';
        document.all['required_date'].value=currentDate;      
        document.all['materialrowid'].value='';
        document.all['nsnnom'].value='';
        document.all['fsccode'].value='';
        document.all['boh'].value='';
        document.all['nom'].value='';
      }
}

function editGrid(){
   if (grdMaterial.getNumSelRows()==0) {
       alert('Please select the row you wish to edit, then try again');
       return; 
   }
      editmode = 'Y';
      document.all['stockeddiv'].style.display="none";
      document.all['selectdiv'].style.display="none";
      document.all['nsnnomdiv'].style.display="none";    
      document.all['inpdiv'].style.display="block";
      disableField('nsn');
      disableField('nom');
      mode = 'edit';
      document.all['AREAMatGrid'].style.display="none";
      document.all['AREAIssue'].style.display="block";
      var vSelRow=grdMaterial.getSelRow();
      document.all['nsn'].value=grdMaterial.getCellText(3, vSelRow)+'';
      document.all['nom'].value=grdMaterial.getCellText(4, vSelRow)+'';
      document.all['qty_req'].value=grdMaterial.getCellText(5, vSelRow)+'';
      document.all['required_date'].value=grdMaterial.getCellText(6, vSelRow)+'';
      document.all['materialrowid'].value=URLEncode(grdMaterial.getCellText(7, vSelRow)+'');
      document.all['boh'].value=grdMaterial.getCellText(8, vSelRow)+'';
}

function cancelAdd(){
      mode = 'edit';
      document.all['AREAMatGrid'].style.display="block";
      document.all['AREAIssue'].style.display="none";
}

function saveMaterial(pTab){
   for (var  i=0; i< document.mainform.stocked.length; i++){
     if ( document.mainform.stocked[i].checked){
        chkradio = i;
     }   
   }
   document.all['chkstocked'].value = (document.mainform.stocked[chkradio].value); 

	  if ((document.all['chkstocked'].value =='Y')&&(editmode=='N')){
		 if (document.all['nsnnom'].selectedIndex==-1){
		   alert('Entry: NSN/Nomenclature is required. Please complete this field, then try again.');
		   document.all['nsnnom'].focus();
		   return false;  
		}
	  }

	  if (document.all['chkstocked'].value !='Y')
	  {
		if (document.all['nsn'].value.length==0){
		  alert('Entry: NSN is required. Please complete this field, then try again.');
		  document.all['nsn'].focus();
		  return false;  
		}
	 
		if (document.all['nom'].value.length==0){
		  alert('Entry: Nomenclature is required. Please complete this field, then try again.');
		  document.all['nom'].focus();
		  return false;  
		}
	  }
 
    if (document.all['required_date'].value.length==0){
      alert('Entry: Date Required is required. Please complete this field, then try again.');
      document.all['required_date'].focus();
      return false;  
    }  

    if (document.all['qty_req'].value.length==0){
      alert('Entry: Quantity Required is required. Please complete this field, then try again.');
      document.all['qty_req'].focus();
      return false;  
    }
	
 	 saveForm(pTab);
}

function acroChange(){  
	
	//jesse 2012.5.17 
	document.all['txt_fssmaid'].value = 'for FSS only'; 
	
   var objAcros=new ActiveXObject('microsoft.XMLDOM');              
   objAcros.async=false;
   selAcro=document.all['acronym'].options[document.all['acronym'].selectedIndex].value; 
   objAcros.load("em_work_pkg.em_XMLIs4772Req?pSiteID="+currentSite+"&pAcronym="+selAcro);
   
   if (objAcros.parsed) {       
      var result=objAcros.selectSingleNode('F4772REQ');
	  
	    					if(currentSite == '-303') { alert('line 2215 objAcros.text = '  + objAcros.text);}; 
		
      if (objAcros.text=='N') {        
           document.all['divEquipNo'].innerHTML='<font color=black>*  </font><b>  Equip No</b><BR><input type="text" widget name="equip_no" id="equip_no" maxlength="10" TABINDEX="1" onKeyPress="enterequip()" size="6" />';
           document.all['req_4772'].value='N';
           document.all['divClass'].innerHTML='<font color=black></font><b>Class</b><BR><select widget name="class_code" id="class_code" TABINDEX="1" onKeyPress="enterclass()" style="width: 90"></select>';         
           loadList('class_code', getServerPath()+"em_work_pkg.em_XMLClassList_acro?pSiteID="+currentSite+"&p_acro="+selAcro,'Y'); 
           enableField('equip_no');
          //2020.12.04 document.all['equip_no'].focus();  

      } else { 
	  		////document.all.wo_desc.value = document.all.wo_desc.value  || '-x'; 
			
           document.all['divClass'].innerHTML='<font color=black></font><b>Class</b><BR><input type ="text" name="class_code" id="class_code" TABINDEX="1" style="width: 90">';
           disableField('class_code'); 
		   document.all['req_4772'].value='Y';
		   
           document.all['divEquipNo'].innerHTML='<font color=black>*  </font><b>  Equip No</b><BR><select widget name="equip_no" id="equip_no"  TABINDEX="1" style="width: 90" onKeyPress="typeAhead(); enterequip();" onChange="equipChange()"></select>';
		   
           loadList('equip_no', getServerPath()+"em_work_pkg.em_XMLEquipList_acro?psiteid="+currentSite+"&p_subsite_id="+document.all['subsite_id'].value+"&p_acro="+selAcro+"&p_equip_status=A"+"&pRegNo="+document.all['reg_no'].value,'Y');
           if (document.all['equip_no'].options.length==1) {
              //alert('There are no equipment for this acronym at subsite '+document.all['subsite_id'].value+'.');
              disableField('equip_no');
              //document.all['acronym'].selectedIndex=-1;
              //document.all['acronym'].focus();
              return;
           }  
		   
		   //keyboard inputting need this field enabled. disableField('equip_no');	   
		   
       }
   } 
   ckhopms = objAcros.text;
   
 /*  jesse move this scriptlet to localblur() for the keyboard inputting
   if (objAcros.text != 'N') {  
      allopms();
   }
 */ 
 
}

//==========================================================================

function find() 
{  
// REPLACEMEN CODE starts:--------------------------------------------------
	  findWinWO = "";
      //
	  var vDialogWidth=850; //650;
      var vDialogHeight=450;
      var vRegNo=document.all.reg_no_key.value;
      var vWONo=document.all.wo_no_key.value;
      var winleft = (screen.width /2) - (vDialogWidth /2); 
      var wintop = (screen.height / 2) - (vDialogHeight / 2); 
	  //
      findWinWO=window.open('em_work_pkg.em_woSearchDialog?pSite='+currentSite+
														 '&pRegNo='+vRegNo+
														 '&pWONo='+vWONo +'&pWOStatus=E', 'workorder', 'left='+winleft+
														 ',top='+wintop+',width='+vDialogWidth+',height='+vDialogHeight+
														 ',status=yes,menubar=no,toolbar=no,resizable=yes,scrollbars=yes');
//REPLACEMENT CODE ends.-----------------------------------------------------
//
//ORIGINAL CODE starts:
//   if (findWinWO=='') 
//   {   
//      var vDialogWidth=650;
//      var vDialogHeight=450;
//      var vRegNo=document.all.reg_no_key.value;
//      var vWONo=document.all.wo_no_key.value;
//      var winleft = (screen.width / 2) - (vDialogWidth / 2); 
//      var wintop = (screen.height / 2) - (vDialogHeight / 2); 
//	  //
//      findWinWO=window.open('em_work_pkg.em_woSearchDialog?pSite='+currentSite+
//														 '&pRegNo='+vRegNo+
//														 '&pWONo='+vWONo +'&pWOStatus=E', 'workorder', 'left='+winleft+
//														 ',top='+wintop+',width='+vDialogWidth+',height='+vDialogHeight+
//														 ',status=yes,menubar=no,toolbar=no,resizable=no,scrollbars=no');
//   }
//   else 
//   {
//      try{ findWinWO.focus(); } 
//	    catch(e) { findWinWO=''; find() }
//   }
// ORIGINAL CODE ends.
}//// end of find();
//
//==========================================================================
//

/*jesse, called by finding page after "Go" is pressed AND a gridline is selected. 
	Refer to em_work_pkg.pbk at line 1908 sth like  htp.p('   opener.go1(vThisRow);') "
*/
//BTW, very bad name convention!  
function go1(rowid) {
	//jesse 2015.2.4  for reactive work order, which work code is 08, reg no is x01 (default one), react_id is not NULL
	v_react_id = 0;
	//jesse 2020.01.30 BLDGS work orders, to disable all elements and button SAVE and DELETE
	v_cm_active = 'N'; 
	
    clearFindScreenMWO();

   if (listallregno=='N') {
     loadList('reg_no', getServerPath()+"em_work_pkg.em_getXMLRegisterListAll?pSiteID="+currentSite, 'N');   
   }  
   listallregno='Y';
      
   //jesse 2020.01.16 always load the acronyms for FIND(), including some ones with cm_active =N 
   //if (document.all['acronym'].value.length ==0){ 
     loadList('acronym', getServerPath()+'em_work_pkg.em_XMLGetAcroListEquipCnt?pSiteID='+currentSite, 'Y');
   //} 
   
   document.all['divwono'].innerHTML='<font color=black>* </font><b>WO NO</b><BR><input type="text" widget name="wo_no" id="wo_no" size="6" TABINDEX="1" onKeyPress="enterwono()" TABINDEX="1" onblur="localonBlur(this.name)"  />';
   document.all['divEquipNo'].innerHTML='<font color=black>* </font><b>Equip No</b><BR><input type="text" widget name="equip_no" TABINDEX="1" id="equip_no" size="6" />';   
   document.all['req_4772'].value='N';
   document.all['divClass'].innerHTML='<font color=black></font><b>Class</b><BR><input type="text" name="class_code" id="class_code" TABINDEX="1" style="width: 90">';          
   disableField('class_code');
   
   if (document.all['subsite_id'].value.length == 0){
     	loadList('subsite_id', getServerPath()+"em_office_pkg.getXMLSubsiteList?pSiteID="+currentSite, 'N');      
   }   
   
   var objFind=new ActiveXObject('microsoft.XMLDOM');
   objFind.async=false; 
   objFind.load("em_work_pkg.em_getWO?p_rowid="+rowid+"&pSiteID="+currentSite); 
   var nodeList = objFind.getElementsByTagName('ROW');
   
   if(nodeList.length==0) 
   { 
          alert('No matching record found in the database'); 
          return;
   }   
   
   //alert('cm_active=?x');
   for(i=0; i < nodeList(0).childNodes.length; i++) 
   {
    	try{ //jesse 2015.2.4
			if(nodeList(0).childNodes(i).nodeName.toLowerCase() == 'react_id'){
				v_react_id = nodeList(0).childNodes(i).text; 
			//jesse 2015.2.4
			}else{ 
				if(nodeList(0).childNodes(i).nodeName.toLowerCase() == 'cm_active'){
					v_cm_active = nodeList(0).childNodes(i).text; 
				}else{
					//jesse document.all[nodeList(0).childNodes(i).nodeName.toLowerCase()].value=null;
					document.getElementById(nodeList(0).childNodes(i).nodeName.toLowerCase()).value=nodeList(0).childNodes(i).text; 
					
					//jesse 2015-4-6
					if(nodeList(0).childNodes(i).nodeName.toLowerCase() == 'subsite_id'){
						loadList('whep_fac_id', getServerPath()+"em_work_pkg.getXMLFacility_whep_fac_id?pSiteID="+currentSite+"&pSubsiteID="+nodeList(0).childNodes(i).text, 'Y'); 
					} 
					//jesse 2015-4-6  
				}
			}
		}catch(e){ }
    }		
     
	 ///////////////alert('cm_active=?'+v_cm_active);
	        
   	  document.all['capDiv'].innerHTML="Work Order Details"; 

       if (document.all['work_code'].value=='25') {
                document.all.p_acbm_ind.checked=true;
                //jesse 2015.2.4 disableField('work_code');
                document.all['p_acbm_lbl'].innerHTML='<font color=BLUE> PS Form 8210 must be approved prior to creating <br />any ACBM work orders.To access form Click </font> <input type="button" value="PS Form 8210" onClick="ps8210()">';                  
       }else{
            document.all.p_acbm_ind.checked=false;
            document.all['p_acbm_lbl'].innerHTML='';              
           //jesse 2015.2.4  enableField('work_code');
      }
  
  	 if (document.all.equip_no.value=='ALL') 
     {
           var objest = new ActiveXObject('microsoft.XMLDOM');
           objest.async=false;
           objest.load('em_work_pkg.em_esthours_prc?pSiteID='+currentSite+'&pWoPtr='+document.all.wo_ptr.value);
           if(objest.parsed) 
           {
             var loadest = objest.selectSingleNode('ESTHOUR');
             var lest = loadest.text;
           }
           document.all.est_hours.value=lest;
           objest=false;
   	 }   

     var objP = new ActiveXObject('microsoft.XMLDOM');
     objP.async=false;
     objP.load('em_work_pkg.em_check_parts?pSiteID='+currentSite+'&pWOPtr='+document.all['wo_ptr'].value);
 
     if(objP.parsed){
       var loadlist = objP.selectSingleNode('PARTS');
       var pExist = objP.text;
     }

     if(pExist=='4') {
        document.all['partDiv'].innerHTML="-----Parts";     
     }else{
        document.all['partDiv'].innerHTML="";
     }
 
    objP = null;
    objFind=null;

    document.all.AREAMat.style.display="none";
    document.all['AREAMain'].style.display="block";
    document.all['BlankAREA'].style.display="none";
	
    document.all['btnMaterial'].style.visibility="visible";
    document.all['btnMaterial'].style.display="block";
    enableField('btnMaterial');
    enableField('btnMaterialsAdd'); 
    enableField('btnRegistersAdd'); 
    enableField('btnReport'); 
    enableField('btnMTSC'); 
    enableField('btnInfo'); 
	//enableField('past_due_reason_code');
    disableField('standing_flag');
    disableField('date_due');
    disableField('crew_id');
    disableField('dept_code');
    disableField('hours_to_date');
    disableField('wo_no');
    disableField('reg_no');
    disableField('subsite_id');
    disableField('acronym');
    disableField('equip_no');
	disableField('priority_code');
	disableField('date_issued');

    /*jesse 
	if((document.all['wo_status'].value=='C')||(document.all['ecbm_view_only'].value=='Y')){
			disableField('wo_desc');
			disableField('work_code');
			disableField('date_issued');
			disableField('est_hours');
			disableField('wo_comment');
			disableField('project_id');
			disableField('priority_code');
			disableField('p_acbm_ind');   
			disableField('past_due_reason_code');
    }else{
			enableField('wo_desc');
			//jess enableField('work_code');
			disableField('work_code');
			//jesse enableField('date_issued');
			disableField('date_issued');
			enableField('est_hours');
			enableField('wo_comment');
			enableField('crew_id');
			enableField('dept_code');
			//jesse enableField('priority_code');
			disableField('priority_code');
			enableField('p_acbm_ind');
			enableField('project_id');
   } */
	
        //------------------------------------------------------------------------
		f_Enable_Desable_FSSP();
		
        //------------------------------------------------------------------------
        f_populateBuletinNo();
        //------------------------------------------------------------------------
		
        //jesse with(self.document.forms[0])
		var reg_no23 = document.all['reg_no'].value.substring(1,3); 
		selAcro=document.all['acronym'].options[document.all['acronym'].selectedIndex].value;

		var vWOStatusValue = document.all.wo_status.options[document.all.wo_status.options.selectedIndex].value;
		if((vWOStatusValue == 'C') || (document.all['ecbm_view_only'].value=='Y') || v_cm_active == 'N'){
        	disableField('txt_fssp_call');
            disableField('txt_fssp_problem');
            disableField('txt_fssp_zip');
			
			disableField('wo_desc');
			disableField('work_code');
			disableField('date_issued');
			disableField('est_hours');
			disableField('wo_comment');
			disableField('project_id');
			
			disableField('p_acbm_ind');   
			disableField('past_due_reason_code');
			disableField('whep_fac_id');
			disableField('crew_id');
			
			if(reg_no23 == '97' && document.getElementById('standing_flag').value == 'Y' && document.getElementById('wo_desc').value.toUpperCase().trim() == 'MISCELLANEOUS CUSTODIAN WORK') {
				//2021.02.01  BLDGS Miscellaneous custodian work order, enable comment and save 
				enableField('wo_comment');	
				enableField('btnSave');
			}else{
				//jesse 2020.01.16 no BLDGS work orders can be deleted 
				disableField('btnDelete');
				disableField('btnSave');
			}
			
		}else{ //wo is open 
			
            if (document.getElementById('txt_fssp_call').value != ''){
				enableField('txt_fssp_call');
                enableField('txt_fssp_problem');
                enableField('txt_fssp_zip');
           }else{    
               disableField('txt_fssp_call');
               disableField('txt_fssp_problem');
               disableField('txt_fssp_zip');
           }	
		   
		   if(reg_no23 > '64'){	
		   		disableField('work_code');
				disableField('dept_code');
				disableField('est_hours');
				disableField('wo_desc');
				disableField('project_id');
				disableField('whep_fac_id');
		   }else{
				enableField('dept_code');
				enableField('est_hours');
				enableField('wo_desc');
				enableField('project_id');
				enableField('whep_fac_id');
				
				//reactive work order, x01 is default register for now
				if(v_react_id > 0){ //work code should be 08, 
					enableField('work_code');
				}else{
					disableField('work_code');	
				}   		   
		   }
		      	   
		   enableField('wo_comment');
		   enableField('crew_id');
		   enableField('p_acbm_ind');
		   
		   //jesse 2020.01.16 no BLDGS work orders can be deleted 
			enableField('btnDelete');
			enableField('btnSave');
	 } //wo is open 

/* jesse 2014.9.18
		//Jesse, 2010/9/19 if reg number > 70, disable all objects except button Save, Exit and field Crew ID, Work Order Comment
		//			 starting in FY 2011 (2010.9.25)
        //if (document.all['reg_no'].value.substring(1,3) > '70') {
        if (document.all['reg_no'].value.substring(1,3) > '70' && document.all['reg_no'].value.substring(1,3) != '98'  ) {
		
		 
        	for (i=0; i < mainform.elements.length; i++) { 
				var aObj = mainform.elements[i];
				
        	  	try {
        			if (aObj.name == 'crew_id' ||  aObj.name == 'btnSave' ||  aObj.name == 'btnClose'  || aObj.name == 'btnFind'  || 
						aObj.name == 'btnMaterialsAdd' || aObj.name == 'btnRegistersAdd' ||  aObj.name == 'btnInfo' ||  aObj.name == 'btnReport'  || aObj.name == 'btnMTSC'  ||
						aObj.name == 'fssWOStart' || aObj.name == 'fseWOEn')
					{				
						enableField(aObj.name);			
					}else if(aObj.name == 'reg_no_key' || aObj.name == 'wo_no_key' ||  aObj.name == 'wo_desc_key' ||  aObj.name == 'wo_comment' ){
						enableField(aObj.name);
						aObj.readOnly = false;
        			}else{ 
						aObj.disabled = 'disabled';
						disableField(aObj.name);
					}
     		 	} catch(e) { }     
				
  		 	} //end of for
			
			if (document.all['acronym'].value == 'FSS'  || document.all['acronym'].value == 'SAMP' || document.all['acronym'].value == 'FSSOCR' || document.all['acronym'].value == 'FSSSMS')
			{ ;  }
			else{	document.all['txt_fssmaid'].value = 'for FSS only';   }
			
		}//end of if 
        //Jesse, if reg number > 70, disable all objects except button Save, Exit and field Crew ID, Work Order Comment
*/ 

		// Added by Michael E. Harris on 12/02/2011 to enable or disable past_due_reason_code dropdown
		// in the case that the current work order is past due.
		////////////////// jesse 2010.01.31 only check it on the open work orders, see procedure em_isWOPastDue(pSiteID in em_site_t.site_id%type,pWoPtr in em_wo_t.wo_ptr%type default null)  
		if (f_isPastDue() && inList(document.getElementById('wo_status').value,'O,W,L') ) { 
		
			//enableField('past_due_reason_code');
			if (document.all['past_due_reason_code'].value == '' || document.all['past_due_reason_code'].value == null) { 
				
				v_past_due_saved = false;
				enableField('past_due_reason_code'); 
				enableField('btnSave');
			}else{
				v_past_due_saved = true;
			}
		}else{
			disableField('past_due_reason_code');
			v_past_due_saved = true;
		}
 
		// Added by Michael E. Harris on 07/26/2014 for the purpose of work order certification.
		// The following code checks to see if the current user has work order certification privileges. 
		// Additionally, it checks to see if the current work order is closed and the work code is 11,12,17,18.
		// If the current work order meets these conditions, then display the "Certify" button for the work order to be certified...p
		if (document.getElementById('wo_status').value=='C' && inList(document.getElementById('work_code').value,'11,12,13,17,18') && 
			document.getElementById('modification_no').value !='' && document.getElementById('modification_no').value !=null) {
			var objWoCert = new ActiveXObject('microsoft.XMLDOM');
			objWoCert.async=false;
			objWoCert.load(getServerPath()+'em_work_pkg.isUserWOCertifier?pSiteId='+currentSite+'&pCurrentUser='+currentUser);
		 
			if (objWoCert.parsed){
				var v_result = objWoCert.selectSingleNode('RESULT');
				var v_isCertifier = objWoCert.text;
				if (v_isCertifier == 'Y') {
					document.getElementById('btnCertify').style.display = 'block';
					enableField('btnCertify');
					document.getElementById('btnCertifyAllWos').style.display='block';
					enableField('btnCertifyAllWos');
				}
				else {
					document.getElementById('btnCertify').style.display = 'none';
					document.getElementById('btnCertifyAllWos').style.display='none';
				}
			}
			
			//If the current work order has been certified, then display the status of the certification beside the title.
			var objisWoCert = new ActiveXObject('microsoft.XMLDOM');
			objisWoCert.async=false;
			objisWoCert.load(getServerPath()+'em_work_pkg.isWOCertified?pSiteId='+currentSite+'&pWO_Ptr='+document.getElementById('wo_ptr').value);
		 
			if (objisWoCert.parsed){
				var v_isCertified = objisWoCert.selectSingleNode('ISCERTIFIED');
				var v_isWOCertified = objisWoCert.text;
				if (v_isWOCertified == 'Y' ) {
					document.getElementById('capDiv').innerHTML = 'WORK ORDER DETAIL <FONT COLOR = "GREEN">(CERTIFIED)</FONT>';
					document.getElementById('btnCertify').style.display = 'none';
				}
				else {
					document.getElementById('capDiv').innerHTML = 'WORK ORDER DETAIL <FONT COLOR = "RED">(NOT CERTIFIED)</FONT>';
				}
			}
			
			//jesse 2015-1-29 cannot be edited for alll bulletin work orders
			disableField('project_id'); 
		}
		else {
				document.getElementById('capDiv').innerHTML = 'WORK ORDER DETAILS';
				document.getElementById('btnCertify').style.display = 'none';
				document.getElementById('btnCertifyAllWos').style.display='none';
		}
		// End of Add by Michael E. Harris


		// Added on 04/04/2015 for Form 4774 by Michael E. Harris
	    /*	if (document.getElementById('wo_status').value=='C' && inList(document.getElementById'work_code').value,'23,08') {
			var objIs4774 = new ActiveXObject('microsoft.XMLDOM');
			objIs4774.async=false;
			objIs4774.load(getServerPath()+'em_cact_pkg.em_get_4774_status?pSite_id='+currentSite+
							'&pWo_ptr='+document.getElementById('wo_ptr').value);
							
			if (objIs4774.parsed) {
				var v_nodeList = objisWoCert.selectSingleNode('FORMSTATUS');
				var v_formstatus = objisWoCert.text;
				
				//If the 4774 has been processed, then do not display the button
				if (v_formstatus == 'I') {
					document.getElementByid('btn4774').style.display='block';
				}
				else {
					document.getElementById('btn4774').style.display='none';
				}
			}
		}*/
				
		
	    /*jesse 2015-2-2	if (document.all['reg_no'].value.substring(1,3) == '98') {
			enableField('wo_comment');
			disableField('wo_desc');
			//disableField('txt_fssp_call');
			//disableField('txt_fssp_problem');
			//disableField('txt_fssp_zip');
			//disableField('txt_fssmaid');
			disableField('priority_code');
			disableField('work_code');
			disableField('crew_id');
			disableField('dept_code');
			disableField('date_issued');
			disableField('est_hours');
			disableField('project_id');
			disableField('whep_fac_id');
		}  */
						
        this.focus();
}

//===================================================================================
function inList(pValue, pList){
	pList = lower(pList);
	pValue = lower(pValue);
	pListArr = pList.split(',');
    for (i=0; i<pListArr.length; i++) {
	if (pListArr[i] == pValue) return true; 
	}
	return false;
	
}
//===================================================================================
function f_populateBuletinNo()
{
	with(self.document.forms[0])
	{
		var v_ModificationNo = modification_no.value;
		var v_selBulletin_3Alpha = v_ModificationNo.substr(0,3);	//nm_selBulletin_3Alpha
		var v_txtBulletin_3Num   = v_ModificationNo.substr(4,3);    //nm_txtBulletin_3Num
		var v_selBulletin_2Num 	 = v_ModificationNo.substr(8,2);	//nm_selBulletin_2Num
	
		var objDate = new Date();
		var vYearCurrent = objDate.getFullYear();
			vYearCurrent = parseInt(vYearCurrent,10); 
		var vYearMinus1  = vYearCurrent - 1;
		var vYearMinus2  = vYearCurrent - 2;
		
			vYearCurrent += ""; // This converts vYearCurrent from INTEGER to STRING
			vYearCurrent = vYearCurrent.substr(2,2); 
			vYearMinus1  += "";
			vYearMinus1  = vYearMinus1.substr(2,2); 
			vYearMinus2  += "";
			vYearMinus2  = vYearMinus2.substr(2,2); 
		
		if (v_ModificationNo == "")
		{
			f_setBulletinAlpha();
			nm_txtBulletin_3Num.value = "";
			f_setBulletinYear();
		}
		else
		{ 
				if (
					 (v_selBulletin_3Alpha != "MMO" && v_selBulletin_3Alpha != "MWO" && v_selBulletin_3Alpha != "MSB" && v_selBulletin_3Alpha != "SMO")
					  ||
					 (isNaN(parseInt(v_txtBulletin_3Num)) == true)
					 // Jesse, 2012.12.19  validation of calendar year violates the fiscal year. e.g.  the year of December of 2012 (12) !=  fiscal year (13)
					//  ||
					// (v_selBulletin_2Num != vYearCurrent && v_selBulletin_2Num != vYearMinus1 && v_selBulletin_2Num != vYearMinus2)
					)
				{
					alert("Bulletin / Modification Number (\""+v_ModificationNo+"\") was created with mistake.\nPlease call Help Desk for assistance.\n"+
														   "Use your Site Name and Bulletin / Modification Number as your referencing to the record.")
				}else{
				
					for (var i=0; i<nm_selBulletin_3Alpha.options.length; i++)
					{
						if (v_selBulletin_3Alpha == nm_selBulletin_3Alpha.options[i].value)
						{
							nm_selBulletin_3Alpha.options[i].selected = true;
							break;
						}
					}
				
					nm_txtBulletin_3Num.value = v_txtBulletin_3Num;
				
					for (var i=0; i<nm_selBulletin_2Num.options.length; i++)
					{
						if (v_selBulletin_2Num == nm_selBulletin_2Num.options[i].value)
						{
							nm_selBulletin_2Num.options[i].selected = true;
							break;
						}
					}
				}
			}
	}
}


function localClose() {
   try{ findWinWO.close() } catch(e) { }
}

// removes select options from a select list
function removeOptions(selectName,strval) {
  for (var i=document.all[selectName].options.length-1; i>=0; i--) {
      if (document.all[selectName].options[i].text == strval){
        document.all[selectName].options[i] = null;
      }  
  }
}

// removes select options from a select list
function removeOptionsVal(selectName,strval) {
  for (var i=document.all[selectName].options.length-1; i>=0; i--) {
      if (document.all[selectName].options[i].value == strval){
        document.all[selectName].options[i] = null;
      }  
  }
}

// removes select options from a select list
function removeAllOptions(selectName) {
  for (var i=document.all[selectName].options.length-1; i>=0; i--) {
        document.all[selectName].options[i] = null;
  }
}

//Jesse, this function will make a drop-down list only has a single item. 
//It is supposed to be used to replace disableField(object_name)
function oneItemizeSelect(p_obj,p_value)
{
	while  (p_obj.length > 1) { 
		if(p_obj.options[0].value != p_value){
			p_obj.remove(0); 
		}else{
			p_obj.remove(1); 
		}
	} 			
}
//Jesse

// adds a new option to the field in 'fieldname'
function addSelectOpt(fieldName,textStr,valueStr) {
   if (document.all[fieldName]) {
      var x = document.all[fieldName].length;  
      if (x == -1) {
          x = 0; 
      }
      document.all[fieldName].options[x] = new Option(); 
      document.all[fieldName].options[x].text = textStr;
      document.all[fieldName].options[x].value = valueStr;
   }
}

function allopms(){ 

	if(currentSite == '-303') {  alert('line 2847 within allopms()');   };
	
  if (document.all['reg_no'].value.substring(1,3) > 70){
		if(isWidgetVisible('equip_no')) {
		  addSelectOpt(document.all['equip_no'].name,'ALL','ALL');
		  removeOptions(document.all['equip_no'].name,'OPMS');
		}  
		else {
		  addSelectOpt(document.all['findequip_no'].name,'ALL','ALL');
		  removeOptions(document.all['findequip_no'].name,'OPMS');
		}	
  }else{	  
    if(document.all['standing_flag'].value == 'Y'){
      	allopms_called = 'Y';
		
		if(isWidgetVisible('equip_no')) {
		  addSelectOpt(document.all['equip_no'].name,'ALL','ALL');
		}else{
		  addSelectOpt(document.all['findequip_no'].name,'ALL','ALL');
		}
	
		  var objOPMS = new ActiveXObject('microsoft.XMLDOM');
		  objOPMS.async = false;
		  objOPMS.load("em_work_pkg.em_existopms?p_site_id="+currentSite+"&p_equip_cat_ptr="+document.all['acronym'].value);
		  if (objOPMS.parsed){
			var result = objOPMS.selectSingleNode('OPMS');
			var opmsexit = objOPMS.text;
		  }
		  if (opmsexit > 0){
				if(isWidgetVisible('equip_no')) {
					//2021.3.15 not needed any more, may use 109-ALL instead. 
					//addSelectOpt(document.all['equip_no'].name,'OPMS','OPMS');
					null;
				}else{
					null; //addSelectOpt(document.all['findequip_no'].name,'OPMS','OPMS');
				}        
		  }  	  
    }else{
		  if(isWidgetVisible('equip_no')) {  
			 removeOptions('equip_no','ALL');
		  }else {
			 removeOptions('findequip_no','ALL');
		  }
    } // end of standing flag
	
  } //end of reg_no > 70 
}

function register() { 
  var keyxml=genKeyXML(); 
  var v_EM_REGISTER_001=showModalDialog("!vm_main_pkg.main?pUser="+currentUser+"&pSiteID="+currentSite+"&pModule=EM_REGISTER_001&pparamxml="+keyxml+"", "", "dialogWidth:470px; dialogHeight:465px; center:yes; status:no; resizeable:no");
  try { doModalResult(v_EM_REGISTER_001) } catch(e) {}
}

function allInfo() {
   var conv=document.all.mod_user.value;
   var createUser=document.all.create_user.value;
   
   if ( document.all['mod_user'].value.indexOf('UNKNOWN') > -1 ) {
      conv='CONVERSION';
   }
   if ( document.all['create_user'].value.indexOf('UNKNOWN') > -1 ) {
         createUser='CONVERSION';
   }
   alert('Created: '+document.all['create_date'].value+' by user: '+createUser+'\n'+' Modified: '+document.all['mod_date'].value+' by user: '+conv);
}

//==================================================================================
//  SASurikov 
// nm_selBulletin_3Alpha
// nm_txtBulletin_3Num 
// nm_selBulletin_2Num
//----------------------------------------------------------------------------------
function f_setBulletinAlpha(){
	with(self.document.forms[0])
	{
		arryBulletin_Val = new Array("-1","MMO","MWO","MSB","SMO");
		arryBulletin_Txt = new Array("=== Select ===","MMO","MWO","MSB","SMO");
	
		nm_selBulletin_3Alpha.options.length = 0;
	
		for (var i=0; i<arryBulletin_Val.length; i++)
		{
			optionX = new Option;
			optionX.value = arryBulletin_Val[i];
			optionX.text  = arryBulletin_Txt[i];
			nm_selBulletin_3Alpha.options[i] = optionX;
		}
		nm_selBulletin_3Alpha.options[0].selected = true;
	
	}
}

//==============================================================================
function f_setBulletinYear()
{
	with(self.document.forms[0])
	{
		var objDate = new Date();
		var vYearCurrent = objDate.getFullYear();
		vYearCurrent = parseInt(vYearCurrent,10); 
		var vYearMinus1  = vYearCurrent - 1;
		var vYearMinus2  = vYearCurrent - 2;
	
		vYearCurrent += ""; // This converts vYearCurrent from INTEGER to STRING
		vYearCurrent = vYearCurrent.substr(2,2); 
		vYearMinus1  += "";
		vYearMinus1  = vYearMinus1.substr(2,2); 
		vYearMinus2  += "";
		vYearMinus2  = vYearMinus2.substr(2,2); 
	
		arryYear_Val = new Array("-1",        vYearCurrent,vYearMinus1,vYearMinus2);
		arryYear_Txt = new Array("= Select =",vYearCurrent,vYearMinus1,vYearMinus2);
	
		for (var i=0; i<arryYear_Val.length; i++)
		{
			optionX = new Option;
			optionX.value = arryYear_Val[i];
			optionX.text  = arryYear_Txt[i];
			nm_selBulletin_2Num.options[i] = optionX;
		}
	
		nm_selBulletin_2Num.options[0].selected = true;
	}
}

//==============================================================================

function onKeyUp_Ctrl(arg_CtrlName)
{
	with(self.document.forms[0])
	{	
		var vValue = eval(arg_CtrlName+".value");
		var vValueLength = eval(arg_CtrlName+".value.length");
		var vValueToBeValidated = "";
		var vPass = "YES";
	
		for (var i=0; i<vValueLength; i++)
		{
			vValueToBeValidated = vValue.substr(i,1);
	
			if ( isNaN(vValueToBeValidated) == true )
			{
				vValue = vValue.substr(0, (vValueLength - 1));
				if ( arg_CtrlName == "nm_txtBulletin_3Num")
				{
					nm_txtBulletin_3Num.value = vValue;
					vPass = "NO";
					alert("Sorry, this field accepts only 3 numeric values.");
				}
				break;
			}
		}
	
		if ( vPass == "YES" && nm_txtBulletin_3Num.value.length == 3)
		{
			nm_selBulletin_2Num.focus();
		}
		
	}
}

//==============================================================================
f_setBulletinAlpha();
self.document.forms[0].nm_txtBulletin_3Num.value = "";
f_setBulletinYear();
//f_Enable_Desable_BulletinNumber();

//==============================================================================

//jesse 2012.5.24  get rid of the "/" and    "\"  from work order comment , which displays work orders improperly
function wo_onChange()
{
	var wo_str = document.getElementById('wo_comment').value;
	
		if( wo_str.indexOf("/")  >  -1  ||  wo_str.indexOf("\\")  >  -1)
		{  
         	alert(" The character /  or  \\  is not allowed.");
			wo_str = wo_str.replace(/\//g," ");
			document.all.wo_comment.value = wo_str;
			document.all['wo_comment'].focus(); 
        }
		
		if( wo_str.indexOf("<")  >  -1  ||  wo_str.indexOf(">")  >  -1)
		{  
         	alert(" The character <  or  >  is not allowed.");
			wo_str = wo_str.replace(/[&\/\\#,+()$~%.'":*?<>{}]/g,' ');
			document.all.wo_comment.value = wo_str;
			document.all['wo_comment'].focus(); 
        }
}

//jesse 2015.4.7 validate the input of est. hours. 
function esthours_onChange(){
	var est_value = null;
	
	est_value = document.getElementById("est_hours").value;
	if(est_value == '') {
		alert('The Est Hrs cannot be left empty.'); 
		document.getElementById("est_hours").value = '0.1';
		document.getElementById("est_hours").focus(); 
	}else{
		//debug jesse, to be continued ...
		//alert('est hours is : ' + est_value); 
		if(isNaN(est_value)){
			alert('Please enter a valid number.');	
			document.getElementById("est_hours").value = '0.1';
			document.getElementById("est_hours").focus(); 	
		}else if(est_value <= 0){
			alert('Please enter a POSITIVE number.');		
		}else{
			if(est_value > 24 && est_value <= 9999){
				alert('Oh, it takes more than a day?');	
			}else if(est_value > 9999){ //a year =8760
				alert('Wow, please ensure "' + est_value + '" is the correct estimated hour, which is more than 9999 hours. ');
				document.getElementById("est_hours").value = '9999';
				document.getElementById("est_hours").focus();
			}
		}
	}
}

//what the heck is this?
function go(){;}

