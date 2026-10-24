/*********************************
   MODIFICATION HISTORY
   Person           Date        Comments
-- Dan Silk         10/2/2003   * fixed object error in getWOno() function which 
                                  gets the workorder number for a piece of equipment.
-- Dan Silk         10/27/2003  * removed commented (old) code
-- 
*********************************/
var f_react_id=document.all['react_id'];                 
var cntEmpDetails=0;            // counter to set reminder for saving employee info handled by enableSaveDetailsBtn()
var findWin='';
var pVal=0;
var reactCallIFrame;
/*
if (f_react_id.value!=null) {   // the react_id value
   pVal=f_react_id.value;
}
*/
var fastMode=false;
var doc=document.all;
var dateNow='';
var timeNow='';
var myXMLobj=new Object();
var alertCounter=15;
var lastWidget;
var eqWin=new Object();

function bodyOnLoad() {
   //show_displayScreenCode();
   objAvailableList=objDoc.getElementsByTagName("ROW");
   disableScreenFields();                    // Initially, all fields are disabled.
   document.body.style.cursor = "default";

   //this adds the HTML for the left column React Calls
   getInnerHTML("callFrameDiv","em_react_002_pkg.em_show_iframe_column?pSiteID="+currentSite+"&p_user_id="+currentUser);  
   toggleMode();
   doc.btnAdd.focus();
}

function onAddCall() {
  
   mode = 'add';
   document.all['call_status'].value='NEW';
   
   swapAddCancelBtns();
   enableMainFields();
   
   reloadEmpLists();
}

function reloadEmpLists() {   
 // get options for lstAvailable select list
   removeOptions('lstAvailable');
   var aURL=getServerPath()+'EM_REACT_PKG.em_XMLAvailableEmps?pSiteID='+currentSite;
   loadNewAvailableList('lstAvailable', aURL);

 // get options for lstDispatched select list   
   removeOptions('lstDispatched');
   var dURL=getServerPath()+'EM_REACT_PKG.em_XMLDispatchedEmps?pSiteID='+currentSite+'&p_react_id='+pVal;
   loadDispatchedList('lstDispatched', dURL);
 }
 
function resetScreen() {
   document.body.style.cursor='wait';
   loadNewAvailableList(getServerPath()+'EM_REACT_PKG.em_XMLAvailableEmps?pSiteID='+currentSite);
   resetAllFields();
   disableScreenFields();
   if (document.all['btnAdd'].disabled==true) {
      swapAddCancelBtns();
   }
   document.all['call_status'].value='';
   call_status_frame.location.href="EM_REACT_002_PKG.em_show_call_status?pSiteID="+currentSite+"&p_user_id="+currentUser;
   document.all['btnMode'].disabled=false;
   document.body.style.cursor='default';
   pVal=0;
   doc.currentUser.value=currentUser;
   doc.currentSite.value=currentSite;
   doc.btnAdd.focus();
}


function validate_form() {
   p_react_id=document.all['p_react_id'];
   date_reported=document.all['date_reported'];
   time_reported=document.all['time_reported'];
   date_closed=document.all['date_closed'];
   time_closed=document.all['time_closed'];
   acronym=document.all['acronym'];
   equip_no=document.all['equip_no'];
   reg_no=document.all['reg_no'];
   wo_ptr=document.all['wo_ptr'];
   react_code=document.all['react_code'];
   noun_code=document.all['noun_code'];
   verb_code=document.all['verb_code'];
   var dateDiff=compareDates('','',date_reported.value,'');
  
   if (dateDiff<0) {
      alert('The Date Reported cannot be greater than today\'s date.');
      date_reported.focus();
      return true;
   }

   if (document.all['react_id'].length>0) {
      if (Math.round(dateDiff)>4) {
         alert('The Date Reported cannot be prior to 4 days ago.');
         date_reported.focus();
         return true;
      }
   }
   if (date_reported.value==null||date_reported.value=='') { 
      alert('Please enter a value for \"Reported On Date\".');
      date_reported.focus();
      return true;
   }
   if (time_reported.value==null || time_reported.value=='') { 
      alert('Please enter a value for \"Reported On Time\".');
      time_reported.focus();
      return true;
   }
   if (acronym.value.length==0) { 
      alert('Please enter an \"Acronym\".');
      acronym.focus();
      return true;
   }
   if (equip_no.value.length==0) { 
      alert('Please enter an \"Equipment No.\"');
      doc.equip_no.focus();
      return true;
   }

   if (react_code.selectedIndex<1) { 
      alert('Please select an \"Trouble\" code.');
      react_code.focus();
      return true;
   }
   if (mode=='close') {
      if (date_closed.length > 0) {
         dateDiff=compareDates(date_closed.value,time_closed.value,date_reported.value,time_reported.value);
         if (dateDiff<0) {
            alert('The Date and Time Closed cannot be prior to the Date and Time Reported.');
            date_closed.focus();
            return true;
         }
      }
      if (noun_code.selectedIndex<1) { 
         alert('Please select a value for \"Action Noun\".');
         noun_code.focus();
         return true;
      }
      if (verb_code.selectedIndex<1) { 
         alert('Please select a value for \"Action Verb\".');
         verb_code.focus();
         return true;
      }
   }
   if (!validateDispatchXML()) {
      return true;
   }
   return false;
}

// This function simply alerts required widgets and their values
function showWidgets() {
   var alertTxt='';
   var widgets=objDoc.getElementsByTagName("widget");
   var widgetType;
   var widgetName;
   for (x=0; x<widgets.length; x++) {
      widgetType=widgets[x].getAttribute("type");
      widgetName=widgets[x].getAttribute("name");
      switch (widgetType) {
         case "select": { alertTxt=alertTxt+widgetName+': '+doc[widgetName].value+'\n'; break }
         case "input": { alertTxt=alertTxt+widgetName+': '+doc[widgetName].value+'\n'; break }
         case "hidden": { alertTxt=alertTxt+widgetName+': '+doc[widgetName].value+'\n'; break }
      }
   }

   alert(alertTxt);
}

function saveOnClick() {
   //showWidgets();
   document.body.style.cursor='wait';
   
  // enable the 'Close Call' button
   enableField('btnCall');
   if (validate_form()) {
      document.body.style.cursor='default';
      return;
   } else {
      saveForm(mode);
      resetScreen();
   }
}

   
/*************************************************************************************
 This function is called by the saveForm() function in em_dialog.js.  React call info 
 is added to em_react_t via saveForm() so now we need to save the dispatched employee 
 info to em_react_emp_t.  This info is in a separate XML obj and is updated in the 
 Dispatch Details area of the react screen. 
 *************************************************************************************/
function save(pTab, saveStatus, newValue) {
   pVal=newValue;
   document.all['react_id'].value=pVal;
   // Build URL to send XML data to Oracle procedure

   var urlStr='em_react_pkg.saveReactEmpDetails?pUserID='+currentUser+'&pSiteID='+currentSite+'&p_react_id='+newValue+'&pMode='+mode+'&pXMLDoc='+objDispatchDoc.xml+'&p_emp_count='+objDispatchedList.length;
   objEmps=getXMLresults(urlStr);
   var msgList=objEmps.getElementsByTagName("message")
   var status=msgList(0).getAttribute("status");
   var msg=msgList(0).text;
   if (status.toLowerCase()!='ok') {
      alert(msg);
   }  
   pVal=0;
   document.body.style.cursor='default';  
}

function onCloseCall() {
   /* fill 'Completed On' Date and Time fields and update the EM_REACT_T record
      and each employee record in EM_REACT_EMP_T for the current react_id */
   mode='close';
   enableCloseFields();
   enableEmpDetails();
   disableField('btnCall');
}

function saveClosedCall() {
   document.body.style.cursor='wait';
   if (document.all['lstDispatched'].length>0) {
      var p_noun_code=document.all['noun_code'].value;
      var p_verb_code=document.all['verb_code'].value;
      var p_closed_date=URLEncode(document.all.date_closed.value.replace('/','-'));
      var p_closed_time=URLEncode(document.all.time_closed.value.replace('/','-'));
      var urlStr='EM_REACT_PKG.NEW_EM_CLOSE_REACT_PRC?pUserID='+currentUser+'&pSiteID='+currentSite+'&p_react_id='+pVal+'&p_noun_code='+p_noun_code+'&p_verb_code='+p_verb_code+'&p_closed_date='+p_closed_date+'&p_closed_time='+p_closed_time;
      objCloseDoc=getXMLresults(urlStr);
      return true;
      swapAddCancelBtns();
   }
   else {
      document.all['lstAvailable'].selectedIndex=0;
      document.all['lstAvailable'].focus();
      alert('No one is currently dispatched for this call.  Please dispatch at least one employee before closing this call.');
      return false;
   }
   document.body.style.cursor='default';
   
}

// when "Dispatch>>" button is clicked
function onDispatchEmps() {
   setVarDateTime();
   document.all['remove'].disabled = false;
   enableEmpDetails();
   var a = document.all['lstAvailable'];   
   var aIndex = a.selectedIndex;
   if (aIndex > -1) {
      var aText =  a.options[aIndex].text;
      var aValue = a.options[aIndex].value;
      var d = document.all['lstDispatched']
      var dLen;
      addSelectOpt(d.name,aText,aValue);            // add the selected available emp to dispatched

   // NEED TO UPDATE XML DOM FOR DISPATCHED EMP HERE AND SET THE DATE AND TIME FIELDS ON 
   // THE SCREEN TO THE VALUES SET IN THE XML DOM.

   // FIRST FIND THE EMP IN THE AVAILABLE XML DOC, THEN ADD VALUES TO DISPATCHED XML DOC
   var thisNodeLength = objAvailableList.length;   
   var thisRow;
   var valueNode;
   var dispatchedNode;
   var addThisNode;
   var laborClassNode;
   var startDateNode;
   var startTimeNode;
   var endDateNode;
   var endTimeNode;
   
      for (x=0; x<thisNodeLength; x++) {                           // loop through the nodes to find the employee info
         thisRow        = objAvailableList(x);                          // Set node object variables, check for matches in 
         valueNode      = thisRow.selectSingleNode('VALUE');       // the list box and update the data in the 
         textNode       = thisRow.selectSingleNode('TEXT');        // DISPATCHED node accordingly.
         dispatchedNode = thisRow.selectSingleNode('DISPATCHED');
         laborClassNode = thisRow.selectSingleNode('LABOR_CLASS');
         startDateText  = dateNow;
         startTimeText  = timeNow;
         if (fastMode==true && mode=='closed') {
            endDateNode    = dateNow;
            endTimeNode    = timeNow;
         } else {
            endDateNode    = '';
            endTimeNode    = '';
         }
         if (valueNode.text == aValue) {         // find the right employee
            dispatchedNode.text = 'Y';

         // ADD THIS PERSON TO THE DISPATCHED XML DOC
         // GET THE ROOT ELEMENT IN THE XML DOC
            var objDispDocRoot = objDispatchDoc.documentElement;

         // ADD A NEW "ROW" NODE
            var addEmp = objDispatchDoc.createElement("ROW");
            addEmp = objDispDocRoot.appendChild(addEmp);
            addEmp.setAttribute("id","0");
   
         // ADD THE CHILD NODES AND VALUES TO "ROW"
            addNewElement(objDispatchDoc,addEmp,"VALUE",valueNode.text);
            addNewElement(objDispatchDoc,addEmp,"TEXT",textNode.text);
            addNewElement(objDispatchDoc,addEmp,"LABOR_CLASS",laborClassNode.text);
            addNewElement(objDispatchDoc,addEmp,"START_DATE",startDateText);
            addNewElement(objDispatchDoc,addEmp,"START_TIME",startTimeText);
            addNewElement(objDispatchDoc,addEmp,"END_DATE","");
            addNewElement(objDispatchDoc,addEmp,"END_TIME","");
            addNewElement(objDispatchDoc,addEmp,"COMMENTS","");
            addNewElement(objDispatchDoc,addEmp,"DISPATCHED",'Y');
            addNewElement(objDispatchDoc,addEmp,"EMP_STATUS",'');
        }
      }
      removeSelectOpt(a.name);                     // remove emp from available list
      a.selectedIndex = -1;
      dLen = d.length;
      d.selectedIndex = dLen-1;
      d.focus();
      refreshDispatchList();
      getEmpDetails(d.name);

      var i;
      try{ 
         for (i=0; i < objAvailableList.length; i++) { 
            var childList=objAvailableList(i).childNodes;     
            if (childList(8).text=='N') {
               var objEntry = document.createElement("OPTION");     
               objEntry.value=childList(0).text;
               objEntry.text=childList(1).text;     
               document.all[pList].options.add(objEntry, i); 
            }
         }
      } catch(e) { 
        return;
      }  
   } 
   getNextElement('dispatch');
}

// add a node to an XML doc by parent node 
function addNewElement(xmlDoc,pNode,newElemName,newElemText) {
   var newElem = xmlDoc.createElement(newElemName);
   newElem = pNode.appendChild(newElem);
   var newElemText = xmlDoc.createTextNode(newElemText);
   newElem.appendChild(newElemText);
}

// with onDispatchEmps functions, adds emp data to the objDispatchedList XML doc for dispatched emps
function addEmpToDList(empID) {
   var i;
   try{ 
      for (i=0; i < objDispatchedList.length; i++) { 
         objDispDocRoot = objDispatchDoc.documentElement;
               
         var childList=objDispatchedList(i).childNodes;     
         var objEntry = document.createElement("OPTION");     
         objEntry.value=childList(0).text;
         objEntry.text=childList(1).text;     
         document.all[pList].options.add(objEntry, i); 
      }
   } catch(e) { 
      return;
   }
}

// retrieves selected employee details and displays then in the employee detail area on screen
function getEmpDetails(listBox) {
   resetEmpDetails();

   var thisBox = document.all[listBox];
   var i = thisBox.selectedIndex;
   document.all['lstAvailable'].selectedIndex = -1;      
   document.all['lstDispatched'].selectedIndex = -1;
   
   var objThisList;
   var thisNodeLength;

   if (listBox == 'lstAvailable') {
      objThisList = objAvailableList;    
   } else if (listBox == 'lstDispatched') {
      objThisList = objDispatchedList;
   } else {
      return;
   }

   thisBox.selectedIndex = i;
   
   var thisNodeLength = objThisList.length;   
   var thisRow;
   var valueNode;
   var dispatchedNode;
   var addThisNode;
   var laborClassNode;
   var startDateNode;
   var startTimeNode;
   var endDateNode;
   var endTimeNode;

   if (i!=-1) {                                                    // if an employee is selected. . .
      for (x=0; x<thisNodeLength; x++) {                           // loop through the nodes to find the employee info
      
         thisRow        = objThisList(x);                          // Set node object variables, check for matches in 
         valueNode      = thisRow.selectSingleNode('VALUE');       // the list box and update the data in the 
         textNode       = thisRow.selectSingleNode('TEXT');        // DISPATCHED node accordingly.
         dispatchedNode = thisRow.selectSingleNode('DISPATCHED');
         laborClassNode = thisRow.selectSingleNode('LABOR_CLASS');
         startDateNode  = thisRow.selectSingleNode('START_DATE');
         startTimeNode  = thisRow.selectSingleNode('START_TIME');
         endDateNode    = thisRow.selectSingleNode('END_DATE');
         endTimeNode    = thisRow.selectSingleNode('END_TIME');
         commentsNode   = thisRow.selectSingleNode('COMMENTS');

// find the right employee
         if (objThisList(x).childNodes(0).text == document.all[listBox].options[i].value) {         
             document.all['empTitle'].value = laborClassNode.text;      // Save details for later use
             document.all['emp_start_date'].value = startDateNode.text;                
             document.all['emp_start_time'].value = startTimeNode.text;
             document.all['emp_end_date'].value   = endDateNode.text;
             document.all['emp_end_time'].value   = endTimeNode.text;
             document.all['react_comments'].value   = commentsNode.text;
         }
      }
   }
}

// disable employee details
function disableEmpDetails() {
    disableField('emp_start_date');
    disableField('emp_start_time');
    document.all['emp_start_date_btn'].disabled = true;
    disableField('emp_end_date');
    disableField('emp_end_time');
    document.all['emp_end_date_btn'].disabled   = true;
    disableField('react_comments');
    document.all['save_emp_info'].disabled      = true;
}

// enable employee details
function enableEmpDetails() {    
   if (doc.lstDispatched.selectedIndex==-1) { 
      enableField('emp_start_date');
      enableField('emp_start_time');
      document.all['emp_start_date_btn'].disabled = false;
      enableField('emp_end_date');
      enableField('emp_end_time');
      document.all['emp_end_date_btn'].disabled   = false;
      enableField('react_comments');         
   }
}

// when an acronym is selected this builds the 'equip no' list
function acronymOnChange() {
   document.all['wo_ptr'].value='';
   var urlStr='em_react_pkg.em_XMLAbbrevEquipList?pSiteID='+currentSite+'&p_equip_cat_ptr='+document.all['acronym'].value;
   myXMLobj=getXMLresults(urlStr);
   //loadList('equipment_ptr', urlStr, 'N');
   //enableField('equipment_ptr');
}

// when an 'equip no' is select, gets value for 'wo no'
function getWOno() {
   document.all['wo_ptr'].value = '';
   var urlStr='em_react_002_pkg.em_GetWOptr?pSiteID='+currentSite+'&p_acronym='+acro.value.toUpperCase()+'&p_equip_no='+eq.value.toUpperCase();
   var eqObj=getXMLresults(urlStr);
   var msgNode=eqObj.getElementsByTagName("message");
   var msgStr=msgNode(0).text;
   var msgStatus=msgNode(0).getAttribute("status");
   if (msgStatus=='ok') {
      return true;
   } else {
      var myConfirm=newConfirm('Equipment No. Error','The equipment number you entered does not exist.\n\nWould you like to search for one?',1,1,0);
      if (myConfirm) {
         //open the acronym finder... same as clicking the find button next to equip/wo section
         alert('open finder');
      } else {
         doc.equip_no.focus();
      }
      return false;
   }
   for (x=0;x<equipList.length;x++) {
      equipPtr=equipList(x);
      if (equipPtr.text==selected_text) {
         w=x-1;
      }
   }
   document.all['wo_ptr'].value = woPtrList(w).text;
   document.all['wo_no'].value = woNoList(w).text;
   
}

// removes the selected option from a select list (used on 'lstDispatched' and 'lstAvailable')
function removeSelectOpt(fieldName) {
   var sIndex = document.all[fieldName].selectedIndex;
   document.all[fieldName].options[sIndex] = null;
}

// adds a new option to the field in 'fieldname'
function addSelectOpt(fieldName,textStr,valueStr) {
   if (document.all[fieldName]) {
      var x = document.all[fieldName].length;  
      if (x == -1) {
         x = 0; 
      }
      var addNewOption = new Option();
      document.all[fieldName].options[x] = addNewOption;   
      document.all[fieldName].options[x].text = textStr;
      document.all[fieldName].options[x].value = valueStr;
   }
}

// removes ALL options from a select list
function removeOptions(selectName) {
  for (var i=document.all[selectName].options.length-1; i>=0; i--) {
      document.all[selectName].options[i] = null;
  }
}

// Refreshes the Dispatched emp list
function refreshDispatchList() {
   objDispatchedList=objDispatchDoc.getElementsByTagName("ROW");
}

// Loads list identified by pList using pQuery. 
function loadDispatchedList(pList, pQuery) {
   objDispatchDoc = null;
   objDispatchDoc = new ActiveXObject("Microsoft.XMLDOM");   // holder for dispatched employees XML doc (see loadDispatchedList function below)
   objDispatchDoc.async=false;
   objDispatchDoc.load(pQuery);

   if (objDispatchDoc.parsed) {
      refreshDispatchList();
   } 
   else { return; }
   
   var i;
//   document.all[pList].options.length=0;
   try{ 
      for (i=0; i < objDispatchedList.length; i++) { 
         var childList=objDispatchedList(i).childNodes;     
         var objEntry = document.createElement("OPTION");     
         objEntry.value=childList(0).text;
         objEntry.text=childList(1).text;     
         document.all[pList].options.add(objEntry, i); 
      }
   } catch(e) { 
      return;
   }
   if (mode!='add') {
      if (doc.lstDispatched.length>0 && doc.lstDispatched.selectedIndex==-1) {
         doc.lstDispatched.selectedIndex=0;
         doc.lstDispatched.focus();
         doc.lstDispatched.onclick();
      } else {
         doc.lstAvailable.focus();
      }
   }
}

// Loads list identified by pList using pQuery. 
function loadNewAvailableList(pList, pQuery) {
   objNewDoc = null;
   objNewDoc = new ActiveXObject("Microsoft.XMLDOM");
   objNewDoc.async=false;
   objNewDoc.load(pQuery);

//   if (objNewList.parsed) {
      objAvailableList=objNewDoc.getElementsByTagName("ROW");
//   } 
//   else { return; }

   setLstAvailableOpts(pList);
}
 
function setLstAvailableOpts(pList) {
   var i;

   try{ 
      for (i=0; i < objAvailableList.length; i++) { 
         var childList=objAvailableList(i).childNodes;     
         if (childList(8).text=='N') {
            var objEntry = document.createElement("OPTION");     
            objEntry.value=childList(0).text;
            objEntry.text=childList(1).text;     
            document.all[pList].options.add(objEntry, i); 
         }
      }
   } catch(e) { 
      return;
   }  
}

function setAvailableNodeValue(nodeStr,valStr,empid) {
   for (i=0; i < objAvailableList.length; i++) { 
      var childList=objAvailableList(i).childNodes;
      if (childList(0).text==empid) {
         childList(8).text=valStr;
      }
   }
}

function dispatchEmps(listBox,dispatchedValue) {
  setVarDateTime();
  if (listBox == 'lstDispatched') {
      document.all['lstAvailable'].selectedIndex = -1;
   }
   if (listBox == 'lstAvailable') {
      document.all['lstDispatched'].selectedIndex = -1;
   }
   var count=0;
   var thisNodeLength = objAvailableList.length;
   var thisBox = document.all[listBox];
   var thisRow;
   var valueNode;
   var dispatchedNode;
   var addThisNode;
   var laborClassNode;
   var startDateNode;
   var startTimeNode;
   var endDateNode;
   var endTimeNode;
   var i = thisBox.selectedIndex;
   var thisEmpTitle='';
   var thisEmpStartDate='';
   var thisEmpStartTime='';
   var thisEmpEndDate='';
   var thisEmpEndTime='';
   
//   alert(i);
   
   if (i!=-1) {                                                    // if an employee is selected. . .

      for (x=0; x<thisNodeLength; x++) {                           // loop through the nodes to find the employee info
      
         thisRow        = objAvailableList(x);                          // Set node object variables, check for matches in 
         valueNode      = thisRow.selectSingleNode('VALUE');       // the list box and update the data in the 
         textNode       = thisRow.selectSingleNode('TEXT');        // DISPATCHED node accordingly.
         dispatchedNode = thisRow.selectSingleNode('DISPATCHED');
         laborClassNode = thisRow.selectSingleNode('LABOR_CLASS');
         startDateNode  = thisRow.selectSingleNode('START_DATE');
         startTimeNode  = thisRow.selectSingleNode('START_TIME');
         endDateNode    = thisRow.selectSingleNode('END_DATE');
         endTimeNode    = thisRow.selectSingleNode('END_TIME');

         if (valueNode.text == thisBox.options[i].value) {         // find the right employee
             thisEmpTitle     = laborClassNode.text;               // Save details for later use
             thisEmpStartDate = startDateNode.text;                
             thisEmpStartTime = startTimeNode.text;
             thisEmpEndDate   = endDateNode.text;
             thisEmpEndTime   = endTimeNode.text;
                                                                   // check if dispatching employee
                                                                   // then confirm dispatch if dispatching employee
             if (dispatchedValue == 'Y' || listBox == 'lstDispatched') {  
               if(confirm("DISPATCH EMPLOYEE? \n \n Are you sure you want to dispatch this employee:\n \n "+textNode.text+"?")) {
                  dispatchedNode.text=dispatchedValue;          // set dispatched status to 'Y'
                  if (startDateNode.text==''){                  // Set start date in XML for current employee
                     startDateNode.text=dateNow;
                  }
                  if (startTimeNode.text==''){                  // Set start time in XML for current employee
                     startTimeNode.text=timeNow;
                  }
                  thisEmpStartDate = startDateNode.text;           // Save details for later use
                  thisEmpStartTime = startTimeNode.text;
                  thisEmpEndDate   = endDateNode.text;
                  thisEmpEndTime   = endTimeNode.text;
			   } else { 
                  return;
			   }
            }
         }
         thisRow        = NULL;
         valueNode      = NULL;
         textNode       = NULL;
         dispatchedNode = NULL;
         laborClassNode = NULL;
         startDateNode  = NULL;
         startTimeNode  = NULL;
         endDateNode    = NULL;
         endTimeNode    = NULL;
         thisEmpTitle     = NULL;
         thisEmpStartDate = NULL;
         thisEmpStartTime = NULL;
         thisEmpEndDate   = NULL;
         thisEmpEndTime   = NULL;
      }

// set employee details area title to default
/*
      if (listBox == 'lstDispatched' || dispatchedValue == 'Y') {
         document.all['empDetailsHd'].style.background = '#f7a6c6';
      } else {
         if (listBox == 'lstAvailable') {
            document.all['empDetailsHd'].style.background = '#9cf7f7';
         } else {
            document.all['empDetailsHd'].style.background = '#BBBBBB';
         }
      }
*/
      document.all['empTitle'].value = thisEmpTitle;
      document.all['emp_start_date'].value = thisEmpStartDate;
      document.all['emp_start_time'].value = thisEmpStartTime;
      document.all['emp_end_date'].value   = thisEmpEndDate;
      document.all['emp_end_time'].value   = thisEmpEndTime;
   } else {
      resetEmpDetails();
   }
   
   if (dispatchedValue != '' && thisBox.selectedIndex != -1) {
      doEmpLists();
      thisBox.selectedIndex = i;
      for (j=0; j<document.all['lstDispatched'].length; j++) {
         if (document.all['lstDispatched'].options[j].value = valueNode.text) {
            document.all['lstDispatched'].focus();
            document.all['lstDispatched'].selectedIndex = j;
         }
      }
   }
//alert(i);
}

function resetEmpDetails() {
   document.all['empTitle'].value = '';
   document.all['emp_start_date'].value = '';
   document.all['emp_start_time'].value = '';
   document.all['emp_end_date'].value = '';
   document.all['emp_end_time'].value = '';
   document.all['react_comments'].value= '';
       
 //set employee details area title to default
   //document.all['empDetailsHd'].style.background = '#BBBBBB';
}

function checkRemove() {
   d = document.all['lstDispatched'];
   dIndex = d.selectedIndex;
   
   if (dIndex>-1) {
      if (d.options[dIndex].text.indexOf('(')==-1) {
         enableField('remove');
      } else {
         disableField('remove');
      }
   } else {
      disableField('remove');
   }
} 

function removeFromDispatchList() {
   d        = document.all['lstDispatched'];
   dIndex   = d.selectedIndex;
   
   if (d.options[dIndex].value.indexOf('(')>-1) {
      if (dIndex>-1 ) {
         theEmpID = d.options[dIndex].value
         removeFromDispatchXML(theEmpID)
         removeSelectOpt(d.name);
//      doEmpLists();
         setAvailableNodeValue('DISPATCHED','N',theEmpID)      
         removeOptions('lstAvailable');
         setLstAvailableOpts('lstAvailable');
      }
      resetEmpDetails();
      disableEmpDetails();
      checkRemove();
   }
}

function onFocusAvailable(control) {
//   if (control.getAttribute('UNSELECTABLE')=="off") {
      control.style.background='Yellow'; 
      disableEmpDetails();
//      enableEmpDetails();
//   }
}

function onBlurAvailable(control) {
//   if (control.getAttribute('UNSELECTABLE')=="off") {
      control.style.background='White'; 
//   }
}

function onEmpStartDateBtn() {
   highlightOnChange('emp_start_date');
   highlightOnChange('emp_start_time');
   enableSaveDetailsBtn();
   calproc('mainform.emp_start_date'); 
}

function onEmpEndDateBtn() {
   highlightOnChange('emp_end_date');
   highlightOnChange('emp_end_time');
   enableSaveDetailsBtn();
   calproc('mainform.emp_end_date'); 
   document.all['emp_end_time'].value=timeNow; 
}

function dispatchOnBlur(control){
//   if (control.getAttribute('UNSELECTABLE')=="off") {
      control.style.background='White'; 
      //document.all['save_emp_info'].disabled=true;  
//   }
}

function dispatchOnFocus(control){
//   if (control.getAttribute('UNSELECTABLE')=="off") {
      control.style.background='Yellow'; 
      document.all['save_emp_info'].disabled=false;  
      //if (document.all['call_status'].value=='OPEN'||document.all['call_status'].value=='NEW') {
         enableEmpDetails();
      //}
//   }
}

function dispatchOnClick(control){
   checkRemove(); 
   getEmpDetails(control.name);

   var d = document.all['lstDispatched'];
   var dSel = d.selectedIndex;
   if (dSel>-1) {
      if (objDispatchedList(dSel).childNodes(9).text=='C') {
         disableEmpDetails();
      } else {
         enableEmpDetails();
      }
   }
}

function enableSaveDetailsBtn() {
   d=document.all['lstDispatched'];
   dIndex=d.selectedIndex;
   if (cntEmpDetails==0||cntEmpDetails==6) {
      alert('Be sure to click \"Save Employee Details\" for each \ndispatched employee to commit any changes in the \nfollowing \"Employee Details\" fields: \n\n  Start Date \n  Start Time \n  End Date \n  End Time \n  Comments. \n\nWhen all of the updates are complete, click the \"Save\" \nbutton at the bottom of the screen to save all changes \nin the current record to the database.');
   }
   cntEmpDetails++;
   if (cntEmpDetails>11) {
      cntEmpDetails=0;
   }
   if (dIndex>-1) {
      //document.all['save_emp_info'].disabled=false;  
      enableField('save_emp_info');
   }
}

// remove employee from dispatch XML
function removeFromDispatchXML(empID){
   var objRemoveNode;
   var thisNodeLength = objDispatchedList.length;

   for (x=0; x<thisNodeLength; x++) {
   
      objParentNode=objDispatchedList(x).parentNode;
      objRemoveNode=objDispatchedList(x);
       
      if (objDispatchedList(x).childNodes(0).text==empID) {
         objParentNode.removeChild(objRemoveNode);
      }
      refreshDispatchList();
      thisNodeLength = objDispatchedList.length
      if (x==thisNodeLength) return;
   }   
}

function onSaveEmpDetails(){
   removeHighlight('emp_start_date');
   removeHighlight('emp_start_time');
   removeHighlight('emp_end_date');
   removeHighlight('emp_end_time');
   removeHighlight('react_comments');

   cntEmpDetails=1;
 
   var d        = document.all['lstDispatched'];
   var dIndex   = d.selectedIndex;
   var theEmpID = d.options[dIndex].value;
   var startDate = document.all['emp_start_date'].value;
   var startTime = document.all['emp_start_time'].value;
   var endDate   = document.all['emp_end_date'].value;
   var endTime   = document.all['emp_end_time'].value;
   var comments  = document.all['react_comments'].value;
   var childListLength;
   var thisNodeLength = objDispatchedList.length;
   var alertStr='';

   for (x=0; x<thisNodeLength; x++) {
      //if (objDispatchedList(x).childNodes(0).text==theEmpID) {
      if (x==dIndex) {
         empName=objDispatchedList(x).selectSingleNode('TEXT').text;
         childListLength = objDispatchedList(x).childNodes.length;
         for (y=0; y<childListLength; y++) {
            if (objDispatchedList(x).childNodes(y).nodeName=='START_DATE') {
               if (objDispatchedList(x).childNodes(y).text!=startDate) {
                  alertStr=alertStr+'\nStart Date: '+startDate;
               }
               objDispatchedList(x).childNodes(y).text=startDate;
            }
            if (objDispatchedList(x).childNodes(y).nodeName=='START_TIME') {
               if (objDispatchedList(x).childNodes(y).text!=startTime) {
                  alertStr=alertStr+'\nStart Time: '+startTime;
               }
               objDispatchedList(x).childNodes(y).text=startTime;
            }
            if (objDispatchedList(x).childNodes(y).nodeName=='END_DATE') {
               if (objDispatchedList(x).childNodes(y).text!=endDate) {
                  alertStr=alertStr+'\nEnd Date: '+endDate;
               }
               objDispatchedList(x).childNodes(y).text=endDate;
            }
            if (objDispatchedList(x).childNodes(y).nodeName=='END_TIME') {
               if (objDispatchedList(x).childNodes(y).text!=endTime) {
                  alertStr=alertStr+'\nEnd Time: '+endTime;
               }
               objDispatchedList(x).childNodes(y).text=endTime;
            }
            if (objDispatchedList(x).childNodes(y).nodeName=='COMMENTS') {
               if (objDispatchedList(x).childNodes(y).text!=comments) {
                  alertStr=alertStr+'\nComments: '+comments;
               }
               objDispatchedList(x).childNodes(y).text=comments;
            }
         }
      }
      refreshDispatchList();
      thisNodeLength = objDispatchedList.length
      if (x==thisNodeLength) return;
   } 
   if (alertStr.length>0) {
      alertStr = 'Updated the following information \nfor '+empName+': \n'+alertStr;
      alert(alertStr);
   } else {
      alert('Employee data unchanged.');
   }
   getNextElement('save_emp_info');
}

// get DISPATCHED XML into string
function getXMLStr() {
   var thisNodeLength = objDispatchedList.length;
   var alertText='<ROWSET>';
   
   for (x=0; x<thisNodeLength; x++) {
      
      alertText=alertText+'<'+objDispatchedList(x).nodeName+'>';
      
      childListLength = objDispatchedList(x).childNodes.length;
         
      for (y=0; y<childListLength; y++) {
         alertText=alertText+'<'+objDispatchedList(x).childNodes(y).nodeName+'>';
         alertText=alertText+objDispatchedList(x).childNodes(y).text;
         alertText=alertText+'</'+objDispatchedList(x).childNodes(y).nodeName+'>';
      }     
      alertText=alertText+'</'+objDispatchedList(x).nodeName+'>';
   }   
   alertText=alertText+'</ROWSET>'
   return alertText;
}

function go(rowid) {
   mode='edit';  
   if (!document.all['btnAdd'].disabled) {
      swapAddCancelBtns();
   }
   enableMainFields();
   var urlStr='em_react_pkg.getReactInfo?p_rowid='+URLEncode(rowid);
   var objReactXML = new ActiveXObject("Microsoft.XMLDOM");
   objReactXML.async = false;
   objReactXML.load(urlStr);
   objRecordDetails = objReactXML.getElementsByTagName("ROW");
   for (i=0; i<objRecordDetails(0).childNodes.length; i++) {
      fieldName=objRecordDetails(0).childNodes[i].nodeName.toLowerCase();
      fieldValue=objRecordDetails(0).childNodes[i].text;
      document.all[fieldName].value=fieldValue;
      document.all[fieldName].disabled=false;
      if (fieldName=='date_closed') {
         enableCloseFields();
      } else if (fieldName=='time_closed') {
         enableCloseFields();
      } else if (fieldName=='react_id') {
         pVal=fieldValue;
      }
   }  
   reloadEmpLists();   
   enableField('btnCall');
   if (document.all['date_closed'].value.length>0) {
      document.all['call_status'].value='CLOSED';
      disableScreenFields();
      disableEmpDetails();
      enableField('lstDispatched');
      focusOnFirstTab();
   } else {
      try {
         doc.btnCall.focus();
      } catch(e) {}
   }
   objRecordDetails=null;
   objReactXML=null;
}

function localClose() {
   try{ findWin.close() } catch(e) { }
}


function resetAllFields() {
   mode='add';
   pVal=0;
   document.mainform.reset();
   removeOptions('lstAvailable');
   removeOptions('lstDispatched');
}

function disableScreenFields() {
   disableField('date_reported');
   disableField('time_reported');
   disableField('date_closed');
   disableField('time_closed');
   disableField('acronym');
   disableField('equip_no');
   disableField('equipFinder');
   disableField('react_code');
   disableField('noun_code');
   disableField('verb_code');
   disableField('lstAvailable');
   disableField('lstDispatched');
   disableField('empTitle');
   disableField('dispatch');
   disableField('remove');
   disableField('save_emp_info');
   disableField('btnCall');
   disableField('btnSave');
   disableEmpDetails();
}

function swapAddCancelBtns() {
   var a=document.all['btnAdd'];
   var c=document.all['btnCancel'];
   
   if (a.style.visibility=='hidden') {
      a.style.visibility = 'visible';
      a.disabled=false;
      c.style.visibility = 'hidden';
      c.disabled=true;
   } else {
      a.style.visibility = 'hidden';
      a.disabled=true;
      c.style.visibility = 'visible';
      c.disabled=false;
   }
}

function enableMainFields() {
   enableField('btnMode');
   document.all.btnMode.tabIndex=-1;
   enableField('date_reported');
   enableField('time_reported');
   if (fastMode) {
      disableField('date_closed');
      disableField('time_closed');
   } else {
      enableField('date_closed');
      enableField('time_closed');
   }
   enableField('acronym');
   enableField('equip_no');
   enableField('equipFinder');
   doc.equipFinder.tabIndex=-1;
   enableField('react_code');
   enableField('lstAvailable');
   enableField('lstDispatched');
   enableField('dispatch');
   enableField('btnSave');
   document.all['empTitle'].readOnly            = true;
   document.all['empTitle'].style.background       = '#FFFFFF'; 

   if (fastMode) {
      setObjDateTime('date_reported','time_reported');
      if (mode=='add') {
         document.all['acronym'].focus();
      }
   } else {
      document.all['date_reported'].focus();
   }
}

function enableCloseFields() {
   mode='close';
   document.all['empTitle'].readOnly = true;
   enableField('date_closed');
   enableField('time_closed');
   setObjDateTime('date_closed','time_closed');
   enableField('noun_code');
   enableField('verb_code');
   document.all['noun_code'].focus();
}

// this used to be the script for the 'Edit Available Employees' button in the swil
function launchEM_REACT_EMP() { 
   //var keyxml=genKeyXML();
   var keyxml="";  
   var vWidth=540;
   var vHeight=475;
   var winleft = (screen.width / 2) - (vWidth / 2); 
   var wintop = (screen.height / 2) - (vHeight / 2);  
   var urlStr="!vm_main_pkg.main?pSiteID="+currentSite+"&pUser="+currentUser+"&pModule=EM_REACT_EMP";
   var v_EM_REACT_EMP=showModalDialog(urlStr, "", "dialogWidth:"+vWidth+"px; dialogHeight:"+vHeight+"px; center:yes; status:no; resizeable:no; scroll: no;"); 
   try { 
      doModalResult(v_EM_REACT_EMP) 
   } catch(e) { }
}

function find() {  
   if (findWin=='') {   
      document.all.module_from.value='EM_REACT_MAINT_001';
    //  var keyxml=genKeyXML();
      var vWidth=600;
      var vHeight=450;     
      var winleft = (screen.width / 2) - (vWidth / 2); 
      var wintop = (screen.height / 2) - (vHeight / 2); 
      var winProps = "dialogWidth:"+vWidth+"px; dialogHeight:"+vHeight+"px; center:yes; status:no; resizeable:no; scroll:no;";
      var urlStr="!vm_main_pkg.main?pSiteID="+currentSite+"&pUser="+currentUser+"&pModule=EM_REACT_FINDER";
      //prompt('finder',urlStr);
      findWin=showModalDialog(urlStr, "", winProps); 
      if (findWin.length>0) {
         if(document.all['btnCancel'].disabled==false) {
            resetScreen();
         }
         go(findWin);
      }
   }
   else {
      try{ findWin.focus(); } catch(e) { findWin=''; find(); }
   }
}   

function launchReactCallHelp() {

}

function iFrameOpenCall(rowid) {
   //parent.focus();
   resetScreen();
   go(rowid);
}

function toggleMode() {
 if (doc.btnMode) {
   if (!fastMode) {
       fastMode=true;
       document.all.btnMode.innerHTML='On';
   } else {
       fastMode=false;
       document.all.btnMode.innerHTML='Off';
   }
 }
}

function validateDispatchXML(){
   if (timeNow=='') {
      setVarDateTime();
   }
   var d = document.all['lstDispatched'];
 //  var objDispatchedNode;   
   var thisNodeLength = objDispatchedList.length;
   var dateDiff=0;
   var x=0;
   objEmpValue = new ActiveXObject("Microsoft.XMLDOM");
   objEmpValue.async=false;
   objEmpValue=objDispatchDoc.getElementsByTagName("VALUE");
   objStartDate = new ActiveXObject("Microsoft.XMLDOM");
   objStartDate.async=false;
   objStartDate=objDispatchDoc.getElementsByTagName("START_DATE");
   objStartTime = new ActiveXObject("Microsoft.XMLDOM");
   objStartTime.async=false;
   objStartTime=objDispatchDoc.getElementsByTagName("START_TIME");
   objEndDate = new ActiveXObject("Microsoft.XMLDOM");
   objEndDate.async=false;
   objEndDate=objDispatchDoc.getElementsByTagName("END_DATE");
   objEndTime = new ActiveXObject("Microsoft.XMLDOM");
   objEndTime.async=false;
   objEndTime=objDispatchDoc.getElementsByTagName("END_TIME");
   for (x=0; x<thisNodeLength; x++) {
//alert(x+' '+d.options[0].value);
      if (objEmpValue(x).text==d.options[x].value) {
         if (objEndDate(x).text!='') {
            dateDiff=compareDates(objEndDate(x).text,objEndTime(x).text,objStartDate(x).text,objStartTime(x).text);
            if (dateDiff<0) {
               alert('Employee End Date and Time cannot be prior to Employee Start Date and Start Time.');
               document.all['lstDispatched'].selectedIndex=x;
               document.all['lstDispatched'].focus();
              return false;
            }
         }
         dateDiff=compareDates(objStartDate(x).text,objStartTime(x).text,document.all['date_reported'].value,document.all['time_reported'].value);      
         if (dateDiff<0) {
            alert('Employee Start Date and Time cannot be prior to Reported Date and Time.');
            document.all['lstDispatched'].selectedIndex=x;
            document.all['lstDispatched'].focus();
            return false;
         }
      }      
      dateDiff=0;
   }
   return true;
}

function getDocAll() {
   var docMax=document.all.length;
   var x=0;
   var alertTxt='';
   for (x=0; x<docMax; x++) {
      
   }
}

/***************************************************************
  This gets data & time from the server 
***************************************************************/
function setObjDateTime(dateField,timeField) {
   var urlStr='EM_TIME_PKG.em_XMLdatetime_prc?p_site_id='+currentSite;
   myXmlObj=getXMLresults(urlStr);
   if (myXmlObj.parsed) {
      if (dateField != null) {
         var dateVal=myXmlObj.getElementsByTagName("DATE");
         doc[dateField].value=dateVal(0).text;
      }
      if (timeField != null) {
         var timeVal=myXmlObj.getElementsByTagName("TIME");
         doc[timeField].value=timeVal(0).text;
      }
   }   
}

function setVarDateTime() {
   var urlStr='EM_TIME_PKG.em_XMLdatetime_prc?p_site_id='+currentSite;
   myXmlObj=getXMLresults(urlStr);
   if (myXmlObj.parsed) {
      var dateVal=myXmlObj.getElementsByTagName("DATE");
      dateNow=dateVal(0).text;
      var timeVal=myXmlObj.getElementsByTagName("TIME");
      timeNow=timeVal(0).text;
   }   
}

function localonBlur(widget) {
   lastWidget=widget;
   if (widget=='verb_code') {
      if (mode=='close') {
         doc.btnSave.focus();
      }
   } 
   var field=doc[widget]; 
   var widgetLabel=getWidgetAttrib(fieldName,"title");
   var widget_lowercase = widget.toLowerCase();

   if (widget_lowercase.indexOf("date")) {
      var time_widget='';
      try {
         time_widget=widgetreplace("Date","Time")
         d[time_widget].focus();
      } catch(e) {
         try {
            time_widget=widgetreplace("date","time")
            d[time_widget].focus();
         } catch(e) {}
      }
   }
}


function localonFocus(widget) {
   if (lastWidget=="equipment_ptr") {
      doc.react_code.focus();
   }
}

function validateAcro() {
   var acro=doc.acronym;
   if(acro.value!='?') {   
      if(acro.value.length>0) {
         var urlStr='em_react_002_pkg.em_validate_acro_prc?pSiteID='+currentSite+'&p_acronym='+acro.value.toUpperCase();
         var acroObj=getXMLresults(urlStr);
         var msgNode=acroObj.getElementsByTagName("message");
         var msgStr=msgNode(0).text;
         var msgStatus=msgNode(0).getAttribute("status");
         if (msgStatus=='ok') {
            return true;
         } else {
            var myConfirm=newConfirm('Acronym Error','The acronym you entered does not exist.\n\nWould you like to search for one?',1,1,0);
            if (myConfirm) {
               openEqFinder();
            } else {
              acro.focus();
            }
            return false;
         }
      } else {
         return true;
      }
   } else {
      openEqFinder();
   }
}

function validateEq() {
   var eq=doc.equip_no;
   var acro=doc.acronym;
   
   if(eq.value!='?') {   
      if(eq.value.length>0) {
         var urlStr='em_react_002_pkg.em_validate_equip_prc?pSiteID='+currentSite+'&p_acronym='+acro.value.toUpperCase()+'&p_equip_no='+eq.value.toUpperCase();
         
//prompt('url',urlStr);
         
         var eqObj=getXMLresults(urlStr);
         var msgNode=eqObj.getElementsByTagName("message");
         var msgStr=msgNode(0).text;
         var msgStatus=msgNode(0).getAttribute("status");
         for (i=0; i<msgNode(0).childNodes.length; i++) {
            fieldName=msgNode(0).childNodes[i].nodeName.toLowerCase();
            fieldValue=msgNode(0).childNodes[i].text;
            document.all[fieldName].value=fieldValue;
         }
         if (msgStatus=='ok') {
            return true;
         } else {
            var myConfirm=newConfirm('Equipment No. Error','The equipment number you entered does not exist.\n\nWould you like to search for one?',1,1,0);
            if (myConfirm) {
               openEqFinder();
            } else {
               eq.focus();
            }
            return false;
         }
      } else {
         return true;
      }
   } else {
      openEqFinder();
   }
}

function openEqFinder() {
   openSwilWin('EM_REACT_EQ_FINDER');
   try {
      window.EM_REACT_EQ_FINDER.document.all.acro_filter.focus();
      window.EM_REACT_EQ_FINDER.document.all.acro_filter.value=doc.acronym.value;
      window.EM_REACT_EQ_FINDER.findAcro();
   } catch(e) {}
}


//**********************************************************************
function saveEmpDetails() {
   var thisBox = document.all['lstDispatched'];
   var i = thisBox.selectedIndex;
   var objThisList;
   var thisNodeLength;

   objThisList = objDispatchedList;
   thisBox.selectedIndex = i;
   
   var thisNodeLength = objThisList.length;   
   var thisRow;
   var valueNode;
   var dispatchedNode;
   var addThisNode;
   var laborClassNode;
   var startDateNode;
   var startTimeNode;
   var endDateNode;
   var endTimeNode;

   if (i!=-1) {                                                    // if an employee is selected. . .
      for (x=0; x<thisNodeLength; x++) {                           // loop through the nodes to find the employee info
      
         thisRow        = objThisList(x);                          // Set node object variables, check for matches in 
         valueNode      = thisRow.selectSingleNode('VALUE');       // the list box and update the data in the 
         textNode       = thisRow.selectSingleNode('TEXT');        // DISPATCHED node accordingly.
         dispatchedNode = thisRow.selectSingleNode('DISPATCHED');
         laborClassNode = thisRow.selectSingleNode('LABOR_CLASS');
         startDateNode  = thisRow.selectSingleNode('START_DATE');
         startTimeNode  = thisRow.selectSingleNode('START_TIME');
         endDateNode    = thisRow.selectSingleNode('END_DATE');
         endTimeNode    = thisRow.selectSingleNode('END_TIME');
         commentsNode   = thisRow.selectSingleNode('COMMENTS');

// find the right employee
         if (objThisList(x).childNodes(0).text == document.all[listBox].options[i].value) {         
             document.all['empTitle'].value = laborClassNode.text;      // Save details for later use
             document.all['emp_start_date'].value = startDateNode.text;                
             document.all['emp_start_time'].value = startTimeNode.text;
             document.all['emp_end_date'].value   = endDateNode.text;
             document.all['emp_end_time'].value   = endTimeNode.text;
             document.all['react_comments'].value   = commentsNode.text;
         }
      }
   }
}


function cust_confirm(title,mess,icon,buts,defbut,mods) {
   retVal = makeMsgBox(title,mess,icon,buts,defbut,mods);
   //retVal = (retVal==6);
   return retVal;
}
