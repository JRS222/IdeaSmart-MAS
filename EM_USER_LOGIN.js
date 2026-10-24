// Last changed Jan  17 by TA to implement HQ Manager role
var tries=0;
var adminUser=false;
vUserID='';
vUserName='';  // 10/2010 - TWA (for https)
var d=document.all;
d.p_fedstrip.selectedIndex = -1;

function bodyOnLoad() {
   document.all['p_user_logon'].disabled=false;
   document.all['p_user_password'].disabled=false;
   document.all.p_user_logon.focus();
   document.body.style.cursor='default';
}

function onFocusField(control) {
   control.style.background="YELLOW";
}

function onBlurField(control) {
   control.style.background="WHITE";
}

//functionality to submit form on enter press (CP)
function submitForm()
{
	var keyvalue = window.event.keyCode;
	if(keyvalue == 13) {
		chkValue(0);
	}
}

//=============================================================================================================================
// 10/2010
// TWA (https) - Opens the vm_hompage_prc screen
//------------------------------------------------
function openAppWin2(siteId)
{
   var objName="emarsMain";
   var winName="emarsSTAG";
   var h=screen.height-56;
   var w=screen.width-16;
   var winProps="top=0,left=0,width="+ w + ",height="+ h + ",directories=0,fullscreen=0,location=0,menubar=0,resizable=0,scrollbars=0,status=0,titlebar=1,toolbar=0";
   var v_Result;
   var fedstrip=document.all['p_fedstrip'].value;
   fedstrip=siteId;
   try  {
	  v_Result=objName+".focus()";
	  eval(v_Result);
   } 
   catch(e) {
	   
	   v_Result=window.open(getServerPathNonSSL()+"vm_hompage_prc?p_user="+vUserID+
																"&p_site="+fedstrip, 
																"", "top=0,left=0,width="+ w + ",height="+ h+
																",directories=0,fullscreen=0,location=0,menubar=0,resizable=0,scrollbars=0,status=0,titlebar=1,toolbar=0");
	   eval(v_Result);
   }
}

function getServerPath(){
  var thisURL=document.location.hostname+':'+document.location.port+'/'+document.location.pathname;
  thisURL=thisURL.substring(0, thisURL.lastIndexOf('/'));
  if (thisURL.indexOf('//') > 0) {
     thisURL=document.location.hostname+':'+document.location.port+document.location.pathname;
     thisURL=thisURL.substring(0, thisURL.lastIndexOf('/'));
  }
  return 'http://'+thisURL+'/';
}
//====================================================================================================================
//------------------------------------------------
// 10/2010
// TWA (https) - Sets non SSL path
//-------------------------------------------------
function getServerPathNonSSL(){
  var thisURL=document.location.hostname+':80/'+document.location.pathname;
  thisURL=thisURL.substring(0, thisURL.lastIndexOf('/'));
  if (thisURL.indexOf('//') > 0) {
     thisURL=document.location.hostname+':80'+document.location.pathname;
     thisURL=thisURL.substring(0, thisURL.lastIndexOf('/'));
  }
  return 'http://'+thisURL+'/';
}

//==================================================================================================================
// Revised code.  Original is saved below.
function chkValue(winVal) {


   var objAcros=new ActiveXObject('microsoft.XMLDOM');              
       objAcros.async=false;
   var username=document.all['p_user_logon'].value
   var password=document.all['p_user_password'].value;
   var guest=document.all['chkGuest'].checked;
   var fedstrip;

   

   if (fedstrip==''||password==''||username=='') {
      // Original alert message:
	  //alert('You must complete all three fields in order to log in.');
	  alert("Please type-in \"ACE Logon ID\" and \"Password\".");
      return;
   }

   if (document.all.siteDiv.style.visibility=='visible') {
      if (document.all['p_fedstrip'].selectedIndex < 1) {
         alert('You must select a site in order to log in.');
         return;
      }
   }

   try { 
      fedstrip=document.all['p_fedstrip'].value;
	  //alert(fedstrip);
   } catch(e) { }

  
//hirad5 112514  
/*	  
   objAcros.load(getServerPath()+"em_user_login_pkg.em_save_user_login?p_Guest="+guest+
							   										 "&p_user_logon="+username+
																	 "&p_user_password="+password+
																	 "&p_fedstrip="+fedstrip);
																	 
*/
																	 
  objAcros.load(getServerPath()+"em_user_login_pkg.em_save_user_login?p_Guest="+guest+
							   										 "&p_user_logon="+em_encrypt(username)+
																	 "&p_user_password="+em_encrypt(password)+
																	 "&p_fedstrip="+fedstrip);

															 
																	 
   var objList=objAcros.selectSingleNode('login');
   var area=objList.getAttribute('area');
   vUserID=objList.getAttribute('user'); 		// 10/2010 - TWA (for https)
   vUserName=objList.getAttribute('username'); 	// 10/2010 - TWA (for https)
   fedstrip=objList.getAttribute('fedstrip'); 	// 10/2010 - TWA (for https)


//alert("SAS2  "+objList.getAttribute('result')+"   fedstrip="+fedstrip)


   if (objList.getAttribute('result')=='FAIL') {
      tries++;
      if (3-tries > 0) {       
         alert('Login Failed. You have '+(3-tries)+' attempt(s) remaining.');
         d.p_user_password.value="";
         d.p_user_logon.focus();
         d.p_user_logon.select();
         window.returnValue="FAIL";
      }
      else {
         alert('You have exceeded the maximum allowable number of failed login attempts.\n\n' + 
               'If this problem persists, contact your system administrator for assistance.');
         window.returnValue="FAIL";
         self.close();
      }  
	  
   } else if (objList.getAttribute('result')=='INACTIVE') {
       alert('This account is currently inactive.  Please contact your system administrator.');
       window.returnValue="FAIL";
   } else if (objList.getAttribute('result')=='DELETED') {
       alert('This account is no longer valid.  Please contact your system administrator.');
       window.returnValue="FAIL";
   } else if (objList.getAttribute('result')=='ADMINISTRATOR') {       
	   alert('This is an Administrator user account. Please select the site you wish to access.');
       loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_options?pRole=AD', 'Y');
	   document.all.siteDiv.style.visibility="visible";
       pickSite();
	   return;
   } else if (objList.getAttribute('result')=='LOCALADMIN') {       
       alert('This is a Local Administrator user account. Please select the site you wish to access.');
       loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_options?pRole=LA', 'Y');
       document.all.siteDiv.style.visibility="visible";
       pickSite();
       return;
   } else if (objList.getAttribute('result')=='AREAMGR') {
       alert('This is an Area Manager user account. '+
			 'Please select the site you wish to access. You will be granted read-only access to the site you select.');
       loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_by_area?pAreaCd='+area, 'Y');
       document.all.siteDiv.style.visibility="visible";
       pickSite();  
       return;
   }
//-----------------------------------------------------------------------------------------------------------
    else if (objList.getAttribute('result')=='DISTRICT_USER')
	{
		loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_optionsDM?pRole=DM&pUserID='+em_encrypt(vUserID), 'Y');
		document.all.siteDiv.style.visibility="visible";
		pickSite();
		return;
	}
//-----------------------------------------------------------------------------------------------------------
 // Vijaya added for eMARS new Org Structure on 09/10/2020
 else if (objList.getAttribute('result')=='MULTIPLE_DIVUSER')
	{
		loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_optionsMD?pRole=MD&pUserID='+em_encrypt(vUserID), 'Y');
		document.all.siteDiv.style.visibility="visible";
		pickSite();
		return;
	}  
//-----------------------------------------------------------------------------------------------------------
 // Vijaya added for eMARS new Org Structure on 09/10/2020
 else if (objList.getAttribute('result')=='MULTIPLE_SITEUSER')
	{
		loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_optionsMS?pRole=MS&pUserID='+em_encrypt(vUserID), 'Y');
		document.all.siteDiv.style.visibility="visible";
		pickSite();
		return;
	}  


    else if (objList.getAttribute('result')=='HQMANAGER') {
       alert('This is a Headquarters Manager user account. Please select the site you wish to access. You will be granted read-only access to the site you select.');
       loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_options?pRole=HM', 'Y');
       document.all.siteDiv.style.visibility="visible";
       pickSite();  
       return;
   } else if (objList.getAttribute('result')=='OK') {    
       if (fedstrip=='') { fedstrip=objList.getAttribute('fedstrip'); } 
       window.returnValue=objList.getAttribute('user')+'-'+fedstrip;
       self.close();
   } else if (objList.getAttribute('result')=='PASSW') {
       alert('This is either your first time accessing EMARS or \nyour password has expired or been reset. \n\nYou will need to change it now');
       //window.returnValue='PASSWORD'; Changed to keep user info so user doesn't have to retype username - ds, 03/26/2003
       window.returnValue=objList.getAttribute('username')+'/'+objList.getAttribute('fedstrip');
  // TWA (https) - New call to change password 
	   var urlStr=getServerPathNonSSL()+"em_user_login_pkg.change_pass_dialog?p_user_logon="+em_encrypt(username)+"&p_fedstrip=1";
	   var v_Pass=showModalDialog(urlStr, "", "dialogWidth: 352px; dialogHeight: 390px; center: yes; status: no; resizeable: no; scroll: no; edge: sunken; help: 0;");
       self.close();
   }

  //--------------------------------------------
  // 10/2010
  // TWA (https) - If not password change 
  // And not fail - Call to main screen function
  //--------------------------------------------
  if (objList.getAttribute('result')!='PASSW') {
     if (window.returnValue!="FAIL") {
		openAppWin2(fedstrip);	//TWA
		}}
}
//========================================================================================================
// Original code of chkValue
/*
function chkValue(winVal) {
   var objAcros=new ActiveXObject('microsoft.XMLDOM');              
       objAcros.async=false;
   var username=document.all['p_user_logon'].value
   var password=document.all['p_user_password'].value;
   var guest=document.all['chkGuest'].checked;
   var fedstrip;


   if (fedstrip==''||password==''||username=='') {
      // Original alert message:
	  //alert('You must complete all three fields in order to log in.');
	  alert("Please provide \"ACE Logon ID\" and \"Password\", and select needed site when system requires.");
      return;
   }

   if (document.all.siteDiv.style.visibility=='visible') {
      if (document.all['p_fedstrip'].selectedIndex < 1) {
         alert('You must select a site in order to log in.');
         return;
      }
   }

   try { 
      fedstrip=document.all['p_fedstrip'].value;
   } catch(e) { }

   objAcros.load(getServerPath()+"em_user_login_pkg.em_save_user_login?p_Guest="+guest+
							   										 "&p_user_logon="+username+
																	 "&p_user_password="+password+
																	 "&p_fedstrip="+fedstrip);
   var objList=objAcros.selectSingleNode('login');
   var area=objList.getAttribute('area');
   vUserID=objList.getAttribute('user'); 		// 10/2010 - TWA (for https)
   vUserName=objList.getAttribute('username'); 	// 10/2010 - TWA (for https)
   fedstrip=objList.getAttribute('fedstrip'); 	// 10/2010 - TWA (for https)


   if (objList.getAttribute('result')=='FAIL') {
      tries++;
      if (3-tries > 0) {       
         alert('Login Failed. You have '+(3-tries)+' attempt(s) remaining.');
         d.p_user_password.value="";
         d.p_user_logon.focus();
         d.p_user_logon.select();
         window.returnValue="FAIL";
      }
      else {
         alert('You have exceeded the maximum allowable number of failed login attempts.\n\n' + 
               'If this problem persists, contact your system administrator for assistance.');
         window.returnValue="FAIL";
         self.close();
      }  
	  
   } else if (objList.getAttribute('result')=='INACTIVE') {
       alert('This account is currently inactive.  Please contact your system administrator.');
       window.returnValue="FAIL";
   } else if (objList.getAttribute('result')=='DELETED') {
       alert('This account is no longer valid.  Please contact your system administrator.');
       window.returnValue="FAIL";
   } else if (objList.getAttribute('result')=='ADMINISTRATOR') {       
	   alert('This is an Administrator user account. Please select the site you wish to access.');
       loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_options?pRole=AD', 'Y');
	   document.all.siteDiv.style.visibility="visible";
       pickSite();
	   return;
   } else if (objList.getAttribute('result')=='LOCALADMIN') {       
       alert('This is a Local Administrator user account. Please select the site you wish to access.');
       loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_options?pRole=LA', 'Y');
       document.all.siteDiv.style.visibility="visible";
       pickSite();
       return;
   } else if (objList.getAttribute('result')=='AREAMGR') {
       alert('This is an Area Manager user account. '+
			 'Please select the site you wish to access. You will be granted read-only access to the site you select.');
       loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_by_area?pAreaCd='+area, 'Y');
       document.all.siteDiv.style.visibility="visible";
       pickSite();  
       return;
   }
//-----------------------------------------------------------------------------------------------------------
//    else if (objList.getAttribute('result')=='DISTRICT_USER')
//	{
		//loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_optionsDM?pRole=DM&pUserID='+vUserID, 'Y');
		//document.all.siteDiv.style.visibility="visible";
		//alert("SAS em_user_login.js; line 176");
        //pickSite();
		//return;
//	}
//-----------------------------------------------------------------------------------------------------------
    else if (objList.getAttribute('result')=='HQMANAGER') {
       alert('This is a Headquarters Manager user account. Please select the site you wish to access. You will be granted read-only access to the site you select.');
       loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_options?pRole=HM', 'Y');
       document.all.siteDiv.style.visibility="visible";
       pickSite();  
       return;
   } else if (objList.getAttribute('result')=='OK') {    
       if (fedstrip=='') { fedstrip=objList.getAttribute('fedstrip'); } 
       window.returnValue=objList.getAttribute('user')+'-'+fedstrip;
       self.close();
   } else if (objList.getAttribute('result')=='PASSW') {
       alert('This is either your first time accessing EMARS or \nyour password has expired or been reset. \n\nYou will need to change it now');
       //window.returnValue='PASSWORD'; Changed to keep user info so user doesn't have to retype username - ds, 03/26/2003
       window.returnValue=objList.getAttribute('username')+'/'+objList.getAttribute('fedstrip');
  // TWA (https) - New call to change password 
	   var urlStr=getServerPathNonSSL()+"em_user_login_pkg.change_pass_dialog?p_user_logon="+username+"&p_fedstrip=1";
	   var v_Pass=showModalDialog(urlStr, "", "dialogWidth: 352px; dialogHeight: 390px; center: yes; status: no; resizeable: no; scroll: no; edge: sunken; help: 0;");
       self.close();
   }

  //--------------------------------------------
  // 10/2010
  // TWA (https) - If not password change 
  // And not fail - Call to main screen function
  //--------------------------------------------
  if (objList.getAttribute('result')!='PASSW') {
     if (window.returnValue!="FAIL") {
		openAppWin2(fedstrip);	//TWA
		}}
}
*/
//========================================================================================================
function pickSite() {
       document.all.p_fedstrip.focus();
       document.all['p_user_logon'].disabled=true;
       document.all['p_user_password'].disabled=true;
} 

function guestLogin() {
   if (document.all.chkGuest.checked) {
      document.all.p_user_logon.value='GUEST';
      document.all.p_user_password.value='GUEST';
      document.all.p_user_logon.disabled=true;
      document.all.p_user_password.disabled=true;
      document.all.p_user_logon.style.background='#DCDCDC';
      document.all.p_user_password.style.background='#DCDCDC';
      document.all.siteDiv.style.visibility="visible";
      alert('Please select the site you wish to access. As a guest, you will only be able to view the catalog for the site you select.');
      loadList('p_fedstrip', getServerPath()+'em_user_login_pkg.em_fedstrip_options', 'Y');
      document.all.p_fedstrip.focus();
   } else {
      document.all.siteDiv.style.visibility="hidden";
      document.all.p_user_logon.value='';
      document.all.p_user_password.value='';
      document.all.p_user_logon.style.background='WHITE';
      document.all.p_user_logon.disabled=false;
      document.all.p_user_password.disabled=false;
      document.all.p_user_password.style.background='WHITE';
      document.all.p_user_logon.focus();
   }
}

function validate_password() {
   if (document.all['p_old_password']) {
      var p_old_password=document.all['p_old_password'].value;
      var p_user_password=document.all['p_user_password'].value;
      var verify_password=document.all['verify_password'].value;
      var username=document.all['p_user_logon'].value
      var fedstrip=document.all['p_fedstrip'].value
      
      var objAcros=new ActiveXObject('microsoft.XMLDOM');              
      objAcros.async=false;

	  //hirad5 112514
      //objAcros.load(getServerPath()+"em_user_login_pkg.em_save_user_login?p_user_logon="+username+"&p_user_password="+p_old_password+'&p_fedstrip='+fedstrip);
      //objAcros.load(getServerPath()+"em_user_login_pkg.em_save_user_login?p_user_logon="+username+"&p_user_password="+encodeURIComponent(p_old_password)+'&p_fedstrip='+fedstrip); --MEH
	objAcros.load(getServerPath()+"em_user_login_pkg.em_save_user_login?p_user_logon="+em_encrypt(username)+"&p_user_password="+em_encrypt(p_old_password)+'&p_fedstrip='+fedstrip);

      objList=objAcros.selectSingleNode('login');
      vUserID=objList.getAttribute('user');

      if (fedstrip==''||username=='') {
         alert('You must complete all fields in order to log in.');
         document.all['p_user_logon'].focus();
         return false;
      }
      
      if (objList.getAttribute('result')!='PASSW') {
         if (objList.getAttribute('result')!='OK') {
            alert('Your Old Password is incorrect.  Please re-enter your Old Password.');
            document.all['p_old_password'].focus();
            return false;      
         }
      }
      
      if (p_old_password.length==0) {
         alert('Please enter your old password.');
         document.all['p_old_password'].focus();
         return false;
      }

      if (p_user_password.length<minPwdLen) {
         alert('Your New Password must be at least '+minPwdLen+' characters long.');
         document.all['p_user_password'].focus();
         return false;
      }

      if (p_user_password.match(/((\/+|\\+)|(\s+))|(\?+)/)!=null) {
         alert('Blank space and the following characters are not\npermitted for use in passwords:\n\n      /   \\   =   ?');
         document.all['p_user_password'].value='';
         document.all['p_user_password'].focus();
         return false;
      }

      if (chkRequiredChars(p_user_password)) {
         alert('Your password must be at least ' + minPwdLen + ' characters in length and ' + 
               '\ninclude at least one digit from each of the following groups:\n' +
               '\n            1. UPPERCASE letters (A-Z)' +
               '\n            2. lowercase letters (a-z)' +
               '\n            3. numbers (0-9)');
         document.all['p_user_password'].focus();
         document.all['p_user_password'].select();
         return false;
      }
      
      if (verify_password=='') {
         alert('Please verify your new password.');
         document.all['verify_password'].focus();
         return false;
      }
      
      if (p_user_password!=verify_password) {
         alert('Your New Password does not match the Verify Password.\nPlease re-enter your New Password and Verify again.');
         document.all['p_user_password'].focus();
         return false;
      }      
   } else {
      alert('There was an error.  Please re-enter\nyour Login information.\nIf this problem persists, contact your\nsystem adiministrator for assistance.');
      return false;
   }
   
   return true;
}

function chkRequiredChars(fieldStr) {
   var ok = false;
   if (fieldStr.match(/[A-Z]+/g)==null) {
     ok = true;
   }
   if (fieldStr.match(/[a-z]+/g)==null) {
     ok = true;
   }
   if (fieldStr.match(/[0-9]+/g)==null) {
     ok = true;
   }
   return ok;
}

function onClickChangePwd() {

   if(validate_password()) {
      if(vUserID!=''){
         alert('Thank you.  You must now login again\nwith your new password.');
         var objAcros=new ActiveXObject('microsoft.XMLDOM');              
         objAcros.async=false;

		//hirad5 112614
         //objAcros.load(getServerPath()+"em_user_login_pkg.em_update_password?p_user_id="+vUserID+"&p_user_password="+document.all['p_user_password'].value);
           objAcros.load(getServerPath()+"em_user_login_pkg.em_update_password?p_user_id="+em_encrypt(vUserID)+"&p_user_password="+em_encrypt(document.all['p_user_password'].value));
          var objList=objAcros.selectSingleNode('login');

         if(objList.getAttribute('result').toLowerCase()=='error') {
            alert('Error: '+objList.getAttribute('message'));
            window.returnValue='FAIL';   
            tries=0;
         }
         window.returnValue='LOGIN';
      } else {
         window.returnValue='FAIL';
         alert('There was an error.  Please re-enter\nyour Login information.\nIf this problem persists, contact your\nsystem adiministrator for assistance.');
      }
      self.close();
   }
    
}

function closePrompt() {
    if (!window.returnValue) {
       alert('You are not currently logged in.');
       window.returnValue='CANCEL';
    }
}

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

function convert_string_to_hash(p_input_string){

//var originalString = "Hello, World!";

// Convert string to hexadecimal
	var encString = '';
	for (var i = 0; i < p_input_string.length; i++) {
		encString += p_input_string.charCodeAt(i).toString(16);
	}
	return encString;
}

