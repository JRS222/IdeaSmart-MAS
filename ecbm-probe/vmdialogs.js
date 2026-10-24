var currentButton;
var currentUser="";

var imgName;
var imgSrcName;
var imgSrc;
var alertStr='';
var thisImg;
var objImageList; 

function getServerPath(){
  var thisURL=document.location.hostname+':'+document.location.port+'/'+document.location.pathname;
  thisURL=thisURL.substring(0, thisURL.lastIndexOf('/'));
  if (thisURL.indexOf('//') > 0) {
     thisURL=document.location.hostname+':'+document.location.port+document.location.pathname;
     thisURL=thisURL.substring(0, thisURL.lastIndexOf('/'));
  }
  return 'http://'+thisURL+'/';
}


function main(pWindow, pDialogWidth, pDialogHeight, pbutton, pcurrentUser, pcurrentSite) {
   var winleft = (screen.width / 2) - (pDialogWidth / 2);
   var wintop = (screen.height / 2) - (pDialogHeight / 2); 

   module=pWindow+'Win';
  
   currentUser= pcurrentUser;
   currentSite=pcurrentSite;
   
   if (isNaN(eval(currentUser))!=true) {
      if (pbutton.name.indexOf('_off') >= -1) {
         clickButton(pbutton);
         if (pbutton.style.border.indexOf('inset') > -1) {   
            currentButton=window.open('!vm_main_pkg.main?pSiteID='+currentSite+'&pUser='+currentUser+'&pModule='+pWindow, pWindow, 'left='+winleft+',top='+wintop+',width='+pDialogWidth+',height='+pDialogHeight+',status=no,menubar=no,toolbar=no,resizable=no,scrollbars=no');
            childWindows[childWindows.length]=currentButton;
         } else {
           currentButton.close();
         }
      }
   } else if (pWindow=='EM_FINDER_001') {
      clickButton(pbutton);
      if (pbutton.style.border.indexOf('inset') > -1) { 
         currentButton=window.open('!vm_main_pkg.main?pUser=9999&pModule='+pWindow, pWindow, 'left='+winleft+',top='+wintop+',width='+pDialogWidth+',height='+pDialogHeight+',status=no,menubar=no,toolbar=no,resizable=no,scrollbars=no');
         childWindows[childWindows.length]=currentButton;
      } else {
        currentButton.close();
      }
   }
}

function scheduler(pbutton,pWidth,pHeight,pWindow) {
   if (currentUser != "") {
      clickButton(pbutton);
      if (pbutton.style.border.indexOf('inset') > -1) { 
         currentButton=window.open('/pls/emarsdev/EM_AP_PLANNER_PKG.view_sched?pUser='+currentUser, pWindow, 'width='+pWidth+',height='+pHeight+',status=no,menubar=no,toolbar=no,resizable=no,scrollbars=no');
/*         currentButton=window.open('/emarsdev/EM_AP_PLANNER_PKG.view_sched?pUser='+currentUser, pWindow, 'width='+pWidth+',height='+pHeight+',status=no,menubar=no,toolbar=no,resizable=no,scrollbars=no'); */
/*         currentButton=window.open('/emarsdev/EM_SCHED_001_PKG.mars_view_ap_sched?pUser='+currentUser, 'EM_SCHED_001', 'width=600,height=400,status=no,menubar=no,toolbar=no,resizable=no,scrollbars=no'); */
      } else {
         currentButton.close();
      }
   }
}

function showInvToolbar(pbutton) {
   if (currentUser != "") {
     //test for transfer items
     var objTransfer = new ActiveXObject("microsoft.XMLDOM");
     objTransfer.async=false;
     objTransfer.load(getServerPath()+'em_inv_transfers_pkg.em_chk_transfer_items_prc?pSiteID='+currentSite);
     if (objTransfer.parsed) {
       var transfer = objTransfer.selectSingleNode('TRANSITEMS');
       objTransfer='';
     }
     clickButton(pbutton);
     if (document.all.invtoolbar.style.display=='none') {
        document.all.invtoolbar.style.display='block';
        pbutton.style.border='2px inset';
        //Transfer Items Message CM
        if (transfer.text == 'Y') {
          alert('There are items to transfer');
        }
     } else {
        document.all.invtoolbar.style.display='none';
        try {
           pbutton.style.border='2px outset';
        } catch(e) { }
     }
   }
   
   if (pbutton.style.border.indexOf('outset') < 1) { 
      var objchk = new ActiveXObject('microsoft.XMLDOM');
      objchk.async=false;   
      objchk.load ('em_inv_update_pkg.em_replaced_exit_prc?pSiteID='+currentSite);
      var qty;
      if (objchk.parsed) {
        var nodelist = objchk.selectSingleNode('QTYREPLACED');
        qty=objchk.text;
      }
      
      if (qty > 0) {
        window.open('em_inv_update_pkg.display_replacement_data?pSiteID='+currentSite, 'prefwin', 'status=yes,menubar=yes,toolbar=yes,resizable=yes,scrollbars=yes') 
        return;
      }   
     
   }   
     
}

function showAdminToolbar(pbutton) {
   if (currentUser != "") {
      clickButton(pbutton);
      if (document.all.admintoolbar.style.display=='none') {
         document.all.admintoolbar.style.display='block';
         pbutton.style.border='2px inset';
      } else {
         document.all.admintoolbar.style.display='none';
         pbutton.style.border='2px outset';
      }
   }
}

function clickButton(pbutton) {
   if (pbutton.style.border.indexOf('outset') > -1) {
      pbutton.style.border='2px inset';
   } else {
      pbutton.style.border='2px outset';
   }
}


function disableButton(pButton) {
  document.all[pButton].src="";
  
}

function enableButton(pButton) {
  document.all[pButton].src="";
}


function getNextElement(field) {
   var fieldFound = false;
   var form = field.form;

   for (var e = 0; e < form.elements.length; e++) {
      if (fieldFound && form.elements[e].type != 'hidden' && form.elements[e].disabled != true) {
         break;
      }
      if (field == form.elements[e]) {
         fieldFound = true;
      }
   }
   return form.elements[e % form.elements.length];
}


function tabOnEnter(field,evt) {
   var keyCode = document.layers ? evt.which : document.all ? evt.keyCode : evt.keyCode;
                                
   if (keyCode != 13) {
      return true;
   } else {
      getNextElement(field).focus();
      //getNextElement(field).select();
      return false;
   }
}


/*--------------------------------------------------------------------
   Purpose: Returns the current time in the format HH24:MI
  
   MODIFICATION HISTORY
   Person           Date        Comments
   Dan Silk         3/14/2003   
   
----------------------------------------------------------------------*/
 function getCurTime() {
   var dNow = new Date();
   var numHours = dNow.getHours();
   var numMinutes = dNow.getMinutes();
   var strHours;
   var strMinutes;

   if (numHours < 10) {
      strHours = '0'+numHours;
   } else {
      strHours = numHours;
   }

   if (numMinutes < 10) {
      strMinutes = '0'+numMinutes;
   } else {
      strMinutes = numMinutes;
   }

   var timeStr = strHours+':'+strMinutes;

   return timeStr;
 }


/*--------------------------------------------------------------------
   Purpose: Returns the current date in the format: MM/DD/YYYY
  
   MODIFICATION HISTORY
   Person           Date        Comments
   Dan Silk         3/14/2003   
   
----------------------------------------------------------------------*/
function getCurDate() {
   var dNow = new Date();
   var numMonth = dNow.getMonth()+1;
   var numDay = dNow.getDate();
   var numYear = dNow.getYear();

   var numMOnth;
   var numDay;


   if (numMonth < 10) {
      strMonth = '0'+numMonth;
   } else {
      strMonth = numMonth;
   }

   if (numDay < 10) {
      strDay = '0'+numDay;
   } else {
      strDay = numDay;
   }

   var dateStr = strMonth+'/'+strDay+'/'+numYear;

   return dateStr;
}

function handOn(curObj) {
   curObj.style.cursor='hand';
}

function handOff(curObj) {
   curObj.style.cursor='default';
}
