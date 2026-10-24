var serverDate;
var serverTime;
var timeOffSet;
var roleName;
hideSplash();   
id = setTimeout("login();",4100);

var oNewP = document.createElement("p");
oNewP.value='N';

// hirad6 to stop the run change =1 to =2.
var v_runit=2; if (v_runit==1) {var uinfo2=setTimeout('check_mtsc_sc()',20000);clearTimeout(uinfo2);if (!uinfo2) {uinfo2=setInterval("uinfo2=null;check_mtsc_sc();",20000); } else {clearInterval(uinfo2);uinfo2=setInterval("uinfo2=null;check_mtsc_sc();",20000);}}

// variables needed for modal URL if change password required:
var vUserLogon='';
var vFedStrip='';
document.all['statusMsg'].innerHTML='<b>not logged in</b>';

//////////////////////////////////////////////////////////////
// TWA 09/17/2010
// Extracts the UserID in the URL that was passed in
// from EM_USER_LOGIN.js and VM_HOMPAGE_PRC
//////////////////////////////////////////////////////////////
function delineate(str)
{
return (str.substring(str.indexOf("=") + 1, str.lastIndexOf("&")));
}


//////////////////////////////////////////////////////////////
// TWA 09/17/2010
// Extracts the SiteID in the URL that was passed in
// from EM_USER_LOGIN.js and VM_HOMPAGE_PRC
//////////////////////////////////////////////////////////////
function delineate2(str)
{
point = str.lastIndexOf("=");
return(str.substring(point+1,str.length));
}

// Get UserID
var locate = window.location.toString();
var currentUser = delineate(locate);

// Get SiteID
var locate = window.location.toString();
var currentSite = delineate2(locate);


function splitString(retVal,delim) {
	var retValAR=retVal.split(delim);
	vUserLogon=retValAR[0];
	vFedStrip=retValAR[1];
}

function login(){ 
	var usrStr;   
	if (currentUser=="") {

		if (v_Result=="CANCEL" || v_Result == "FAIL") {
			window.close();
		} 
		else if (v_Result.indexOf("/")>0) {
			splitString(v_Result,"/");
			urlStr="em_user_login_pkg.change_pass_dialog?p_user_logon="+em_encrypt(vUserLogon)+"&p_fedstrip="+vFedStrip;  // Michael E. Harris - Placed in response to CISO Pen Test Issues
			var v_Pass=window.open(urlStr, "", "dialogWidth: 352px; dialogHeight: 390px; center: yes; status: no; resizeable: no; scroll: no; edge: sunken; help: 0;");
			if (v_Pass=="LOGIN") {
			    login();
			} else {
			   alert("There was a problem with your login information.\nPlease try again.");
			}
		} else {
			currentUser=v_Result;
			currentUser=v_Result.substring(0, v_Result.indexOf('-'));
			currentSite=v_Result.substring(v_Result.indexOf('-')+1);
			loginOK();
		}
   } else {


      if (currentSite=="") {
         login();
      } else {
         loginOK();
      }
	  
   }
   loginAudit(1);
}

function moveSysAdTb() {
	try {
		var leftPx=document.body.clientWidth/2-96;
		var sysad=document.all.sysadmintoolbar;
		sysad.style.left=leftPx+"px";
	} 
	catch(e) {


		alert(e.message);
	}
}

function loginOK() {
	if (currentSite=="") {
		login();
	} 
	else {
		getInnerHTML('em_main_btns','em_user_login_pkg.em_main_menu?p_user_id='+currentUser+'&p_site_id='+currentSite);
		getInnerHTML('admintoolbar','em_user_login_pkg.em_admin_menu?p_user_id='+currentUser+'&p_site_id='+currentSite);
		getInnerHTML('invtoolbar','em_user_login_pkg.em_inv_menu?p_user_id='+currentUser+'&p_site_id='+currentSite);

		if (currentUser!="0"){
			getInnerHTML('sysadmintoolbar','em_user_login_pkg.em_sysadmin_menu?p_user_id='+currentUser+'&p_site_id='+currentSite);
			moveSysAdTb();
		}	
		if (currentUser!="") { 
			var uinfo=setTimeout('getUserInfo()',2000);
		}
	}


	// Added by Michael E. Harris on 09/11/2014
	// If the current user is the maintenance manager, 
	// then notify him/her of the number of work orders that require certification.
 	var objCert = new ActiveXObject('microsoft.XMLDOM');
	objCert.async = false;
	objCert.load('emapp.em_work_pkg.isUserWoCertifier?pSiteId='+currentSite+'&pCurrentUser='+currentUser);
	if (objCert.parsed) {
		var v_result = objCert.selectSingleNode('ISCERTIFIED');
		var v_isCertifier = objCert.text;
		if (v_isCertifier == 'Y') {
		   var vDialogWidth=700;
			var vDialogHeight=700;     
			var winleft = (screen.width / 2) - (vDialogWidth / 2); 
			var wintop = (screen.height / 2) - (vDialogHeight / 2); 
			
			var v_result=window.showModalDialog("emapp.em_work_pkg.em_certification_list_dialog?pSiteID=" + currentSite + "&pCurrentUser="+currentUser+"", "", "dialogWidth:" + vDialogWidth + "px; dialogHeight:" + vDialogHeight + "px; center:yes; status:no; resizeable:no");
		}

	}
			
			/*	var objCert = new ActiveXObject('microsoft.XMLDOM');
	objCert.async = false;
	objCert.load('emapp.em_work_pkg.isUserWoCertifier?pSiteId='+currentSite+'&pCurrentUser='+currentUser);
	if (objCert.parsed) {
		var v_result = objCert.selectSingleNode('ISCERTIFIED');
		var v_isCertifier = objCert.text;
		if (v_isCertifier == 'Y') {
			var objCertCount = new ActiveXObject('microsoft.XMLDOM');
			objCertCount.async = false;
			objCertCount.load('emapp.em_work_pkg.em_certification_count?pSiteId='+currentSite);
			if (objCertCount.parsed) {
				var v_cert_count = objCertCount.selectSingleNode('CERTIFYCOUNT');
				var v_certification_count = objCertCount.text;
				alert('There are '+v_certification_count+' modification work order(s) that require certification.');
			}
		}
	}*/
	// End of addition by Michael E. Harris


	var objOPMS = new ActiveXObject('microsoft.XMLDOM');
	objOPMS.async = false;
	objOPMS.load("EM_COMPOSE_NOTES_001_PKG.checknewnotes?pUserID="+currentUser);
	if (objOPMS.parsed){
		var result = objOPMS.selectSingleNode('NOTEEXIST');
		var opmsexit = objOPMS.text;
	}
	if (opmsexit == "Y") {
		alert("New Notes have been posted. You may view them in the notes module.");
		return;
	}

	
	
	var objRPT = new ActiveXObject('microsoft.XMLDOM');
	objRPT.async = false;
	objRPT.load("EM_AD_EMAIL_RPT_004_PKG.em_ckh_name?pSiteID="+currentSite);
	if (objRPT.parsed){
		var result = objRPT.selectSingleNode('/NOTENTERED/NAMEEXIST');
		var nameexists = result.text;
		result = objRPT.selectSingleNode('/NOTENTERED/EMAILEXIST');
		var emailexists = result.text;
	}
	
	if (emailexists == "0") {
		alert("Please enter the "+nameexists+" email \n  address in the ADMIN - eMail Rpts module.");
	return;
	}
	// initalize the session timer
	// commented out until better solution can be found
	//initTimer();
}

function hmtrim (str) {
	var	str = str.replace(/^\s\s*/, ''),ws = /\s/,i = str.length;while (ws.test(str.charAt(--i)));return str.slice(0, i + 1);
}

function check_mtsc_sc(){
  objmtscSC='';
  var objmtscSC=new ActiveXObject("microsoft.XMLHTTP");
   objmtscSC.open('POST', getServerPath()+'em_inv_issue_pkg.checkPickList_sc?pSiteID='+currentSite, false);
   objmtscSC.send();
  var mtscSCresult=objmtscSC.responseText;
  objmtscSC=null;
   if (mtscSCresult.substring(0, 1) == '0') {
		null;oNewP.value='N';
   } else {
      //document.all.btnPick.style.color='RED';
      //document.all.btnPick.value='PROCESS PICK LIST ('+objmtscSC.responseText+' ITEMS WAITING)';
      //alert("The pick list contains a balance of items to be issued. Please proceed to the pick list to issue these items.");     
		alert("The pick list contains "+hmtrim(mtscSCresult)+" items \n to be issued from MTSC SHOPPING CART. \n Please proceed to the PICK LIST to ISSUE these items.");
		oNewP.value='Y';	  
   }
//  document.body.style.cursor='default';

}

function logout() {
   loginAudit(2);
   self.close();
}

function getUserInfo() {
	if (isNaN(eval(currentUser))!=true) {
		var objUserInfo=new ActiveXObject('microsoft.XMLDOM');              
		objUserInfo.async=false;
		objUserInfo.load(getServerPath()+"em_user_login_pkg.em_get_user_info?p_user_id="+currentUser+"&p_site_id="+currentSite);
		objUserProfile=objUserInfo.getElementsByTagName("login")
		try { 
			var childList=objUserProfile(0).childNodes;
			var divHTML='<table cellspacing="0" cellpadding"0" border="0">\n<tr>\n   ';
			for (i=0; i < childList.length; i++) { 
            if (childList(i).nodeName!='USER_ID') {
               divHTML=divHTML+'<td style="padding-right: 20px"><b>'+childList(i).nodeName.replace('_',' ')+':</b> '+childList(i).text+'</td>\n   ';
            }

            if (childList(i).nodeName=='DATE') {
               var serverDate=childList(i).text;
            }
            
            if (childList(i).nodeName=='TIME') {
               var serverTime=childList(i).text;
            }
            
            if (childList(i).nodeName=='ROLE') {
               var roleName=childList(i).text;
            }

				if (childList(i).nodeName=='CONVERTED') {
					var converted=childList(i).text;
					if (converted=='N') {
						alert('Data for this site is not yet available in eMars. Only the Notes and, depending upon your system permissions, Admin/Security may be accessed at this time.');
					}}
				}



				divHTML=divHTML+'</tr>\n</table>';
				document.all['statusMsg'].innerHTML=divHTML;
		} catch(e) { 
			alert('error: '+e);
			return;
		}
	}
}

function getInnerHTML(tagID,vURL) {
	document.all[tagID].innerHTML=null; 
	var objHTTP=new ActiveXObject("microsoft.XMLHTTP");
	objHTTP.open('POST', getServerPath()+vURL, false);
	objHTTP.send();
	document.all[tagID].innerHTML=objHTTP.responseText;
	objHTTP=null;
}

// Application timeout variables and functions below here...

var timer;
var tmove=0;

function checkTimer() {
	tmove="0";
	var so=setInterval("timeHasElapsed()",30000)
}

function timeHasElapsed() {
	if(tmove=="0") {
		poppedUp = true;
		alert("Page Timed out.\nThis window will now close.");
		var appURL=window.location.href;
		window.location.href=appURL;
		closeChildWindows();
	}
}

function initTimer() {
	//alert('New application timeout function is running.');
	document.onmousemove = function () {
	tmove=tmove+1;
	if (!timer) { 
		timer=setInterval("timer=null;checkTimer();",600000); 
	} else {
		clearInterval(timer);
		timer=setInterval("timer=null;checkTimer();",600000);
	}
	}
}

//
//===================================================================
//===================================================================
//
function mcloseChildWindows() 
{
    alert("All eMARS open Windows are going to close!...Sorry..");
	//
    for (var i=0,len=childWindows.length;i<len;i++) 
	{
        if (childWindows[i] && !childWindows[i].closed) childWindows[i].close();
    }
}
//
//===================================================================
//===================================================================
//
function amccc(p_user_id){
	//
	///////alert(p_user_id)
    //
	mcloseChildWindows();  
	//
	var currentUser=p_user_id;    
    var selObj = document.getElementById("p_fedstrip");
    var selIndex = selObj.selectedIndex;
    var p_site_id=selObj.options[selIndex].value;
    var currentSite=p_site_id;
	//
	
	
    if (currentSite=="") {login();} 
    else 
	{
        getInnerHTML('em_main_btns',   'em_user_login_pkg.em_main_menu?p_user_id='+    currentUser+'&p_site_id='+currentSite);
       	getInnerHTML('admintoolbar',   'em_user_login_pkg.em_admin_menu?p_user_id='+   currentUser+'&p_site_id='+currentSite);
		getInnerHTML('invtoolbar',     'em_user_login_pkg.em_inv_menu?p_user_id='+     currentUser+'&p_site_id='+currentSite);
		getInnerHTML('sysadmintoolbar','em_user_login_pkg.em_sysadmin_menu?p_user_id='+currentUser+'&p_site_id='+currentSite);
		moveSysAdTb();
		///////alert(5);
    }
	//
	if (isNaN(eval(currentUser))!=true) 
	{//1
        var objUserInfo=new ActiveXObject('microsoft.XMLDOM');              
        objUserInfo.async=false;
        objUserInfo.load(getServerPath()+"em_user_login_pkg.em_get_user_info?p_user_id="+currentUser+"&p_site_id="+currentSite);
        objUserProfile=objUserInfo.getElementsByTagName("login")
        try 
		{//1.1 
            var childList=objUserProfile(0).childNodes;
            var divHTML='<table cellspacing="0" cellpadding"0" border="0">\n<tr>\n   ';
            for (i=0; i < childList.length; i++) 
			{//1.1.1 
                if (childList(i).nodeName!='USER_ID') 
				{
					divHTML=divHTML+'<td style="padding-right: 20px"><b>'+childList(i).nodeName.replace('_',' ')+':</b> '+childList(i).text+'</td>\n   ';
				}
                if (childList(i).nodeName=='DATE') {var serverDate=childList(i).text;}
                if (childList(i).nodeName=='TIME') {var serverTime=childList(i).text;}
                if (childList(i).nodeName=='ROLE') {var roleName=childList(i).text;}
                if (childList(i).nodeName=='CONVERTED') 
				{//1.1.1.1
					var converted=childList(i).text;
					if (converted=='N') 
					{
						alert('Data for this site is not yet available in eMars. '+
							  'Only the Notes and, depending upon your system permissions, '+
							  'Admin/Security may be accessed at this time.');
					}
				}//1.1.1.1
        	}//1.1.1
        divHTML=divHTML+'</tr>\n</table>';
        document.all['statusMsg'].innerHTML=divHTML;
        } //1.1
        catch(e) { alert('error: '+e);return;}
    }//1

}
//
//===================================================================
//===================================================================
//
function loginAudit(pValue){
   var objHTTP=new ActiveXObject("microsoft.XMLHTTP");   
   objHTTP.open('POST', getServerPath()+'emapp.em_login_audit_pkg.CheckLogin_InOut?p_user_id='+currentUser+'&p_value='+pValue, false);
   objHTTP.send();
   objHTTP=null;   
}
//
//===================================================================
//===================================================================
//
function em_encrypt(p_string) {
	var encoded_text;
	encoded_text = convert_string_to_hash(p_string);
	//alert(encoded_text);
	var v_hash_t = '';
	var v_hash = 0;
	var v_encoded_text_t = encoded_text+'';
	v_hash = Math.random() *10000;
	v_hash =  parseInt(v_hash, 16);
	v_hash_t = v_hash+'';
	//alert('v_random= '+v_hash_t);
	var v_hash_len = v_hash_t.length+'';
	//alert('v_random_length: '+v_random_len);
	//alert('encoded password: '+v_random_len + v_random_t + v_encoded_text_t);
	return v_hash_len + v_hash_t + v_encoded_text_t;
	//return;
}
//
//===================================================================
//===================================================================
//
function convert_string_to_hash(p_input_string){

	//var originalString = "Hello, World!";

	// Convert string to hexadecimal
	var encString = '';
	for (var i = 0; i < p_input_string.length; i++) {
		encString += p_input_string.charCodeAt(i).toString(16);
	}
	return encString;
}
