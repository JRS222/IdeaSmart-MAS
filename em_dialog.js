
//  --$Header: /EMARS/Development/Web Contents/Script/em_dialog.js 16    9/05/19 9:31a Jesse $

//  Purpose: This is the main javascript file for all emars modules.
//  10/8/2004 - dsilk, converted to text/javascript file \
//              "em_scripts.js".  Previously, this script file was
//              kept in EM_SCRIPTS_PKG.PBK in the EMARS database.

var objDoc;
var objStyle;
var debugMode;
var dataDirtyflag = false;
var mode='edit';
var currentUser;
var currentSite;
var hasHTTPDivs=false;
var gmessagesOn=true;
var fieldtypes=new Array("input","select","text","select-one","select-multiple","textarea","checkbox","radio","button","image");
var childWindows=[];
var currentDate="";
var currenTime="";

function setDateTime() {
   var urlStr="EM_TIME_PKG.em_XMLdatetime_prc?p_site_id="+currentSite;
   myXmlObj=getXMLresults(urlStr);
   if (myXmlObj.parsed) {
      var dateVal=myXmlObj.getElementsByTagName("DATE");
      dateNow=dateVal(0).text;
      var timeVal=myXmlObj.getElementsByTagName("TIME");
      timeNow=timeVal(0).text;
   }
}

function initSwilScreen() {
   try {
      bodyOnLoad();
   } catch(e) { }
}

function closeChildWindows() {
   for (var i=0,len=childWindows.length;i<len;i++) {
      if (childWindows[i] && !window[childWindows[i]].closed) window[childWindows[i]].close();
   }
}

String.prototype.trim = function()
{
  return( this.replace(/^\s*([\s\S]*\S+)\s*$|^\s*$/,"$1") );
}

function refreshApplet(gridName){
  //alert("RF2 Gridname=" + gridName);
  var h=eval(document.all[gridName].height);
  document.all[gridName].height=eval(h+1);
  //alert("Old=" + h + " New=" + eval(h+1));
 }

function clickTab(tab) {
   var iNode=0;
   var tabName=tab.name;
   tabName=tabName.substring(0, 4);
   var objNodeList = objDoc.getElementsByTagName("tab");
   for (iNode=0; iNode < objNodeList.length; iNode++) {
     if (tabName.indexOf(objNodeList(iNode).getAttribute("name")) > -1) {
       objNodeList(iNode).setAttribute("status", "active");
      }
     else {

       objNodeList(iNode).setAttribute("status", "inactive");
     }
   }
   var showbuttons=showTabButtons(tab);
   saveElems(objDoc);
   screendiv.outerHTML=objDoc.transformNode(objStyle);
   myclickTab(tabName);
}

function trim(s) {
  while (s.substring(0,1) == ' ') {
     s = s.substring(1, s.length);
  }
  while (s.substring(s.length-1, s.length) == ' ') {
     s = s.substring(0, s.length-1);
  }
  return s;
}

function getServerPath(){
  var thisURL=document.location.hostname+":"+document.location.port+"/"+document.location.pathname;
  thisURL=thisURL.substring(0, thisURL.lastIndexOf("/"));
  if (thisURL.indexOf("//") > 0) {
     thisURL=document.location.hostname+":"+document.location.port+document.location.pathname;
     thisURL=thisURL.substring(0, thisURL.lastIndexOf("/"));
  }
  return "http://"+thisURL+"/";
  //return "";
}


function doSelects(pModuleID) {
   var oOptCount;
   objDoc = new ActiveXObject("Microsoft.XMLDOM");
   objStyle = new ActiveXObject("Microsoft.XMLDOM");
   objDoc.async = false;
   objStyle.async = false;
   objDoc.loadXML('<?xml version="1.0"?>'+mainpagexml.xml);
   var xslsheet=objDoc.getElementsByTagName('xsl');
   var sitelist=objDoc.getElementsByTagName('site');
   currentSite=sitelist(0).text;
   var userlist=objDoc.getElementsByTagName('user');
   currentUser=userlist(0).text;
   objStyle.load(xslsheet(0).text);

   var objNodeList = objDoc.getElementsByTagName('widget');

   for (iNode=0; iNode < objNodeList.length; iNode++) {
      try {
         if ((objNodeList(iNode).getAttribute('type')=='select')&&(objNodeList(iNode).getAttribute('widget_menu_sql').indexOf(0,1) != '$')){
            var widgetName=objNodeList(iNode).getAttribute('name');
            var procname=objNodeList(iNode).getAttribute('widget_menu_sql');
            if (procname.indexOf('pSiteID')>0) {
               procname=procname+currentSite;
            }
            if (procname.length > 1) {
               objSelect = new ActiveXObject("Microsoft.XMLDOM");
               objSelect.async=false;
               objSelect.load(procname);

               //The following line commented out by DSilk, 12/5/2003:
               //objNodeList(iNode).setAttribute('widget_menu_sql', 'aaa');

               var lstSelect=objSelect.getElementsByTagName('ROW')
               for (iSel=0; iSel < lstSelect.length; iSel++) {
                  newElem=objDoc.createElement('option');
                  newElem.text=lstSelect(iSel).selectSingleNode('TEXT').text;
                  var sel=lstSelect(iSel).selectSingleNode('SELECTED');
                  newElem.setAttribute('value', lstSelect(iSel).selectSingleNode('VALUE').text);
             var x=objNodeList(iNode).appendChild(newElem);
               }
            }
         }
         if (objNodeList(iNode).getAttribute('type').indexOf('vis')>0) {
            if (objNodeList(iNode).getAttribute('widget_menu_sql') != '') {
              hasHTTPDivs=true;
            }
         }
      } catch(e) { }
   }
   document.all['mainpagediv'].innerHTML=objDoc.transformNode(objStyle);

   if (hasHTTPDivs==true) {
      var objNodeList = objDoc.getElementsByTagName("widget[@type='division']");
      for (iNode=0; iNode < objNodeList.length; iNode++) {
        if (objNodeList(iNode).getAttribute('widget_menu_sql') != '') {
           n=objNodeList(iNode).getAttribute('name');
           x=objNodeList(iNode).getAttribute('widget_menu_sql');
           if (x.indexOf('pSiteID')>0 || x.toLowerCase().indexOf('p_site_id')>0) {
              x=x+currentSite;
           }
           x=getServerPath()+x;
           var objHTTP=new ActiveXObject("microsoft.XMLHTTP");
           objHTTP.open('POST', x, false);
           objHTTP.send();
           document.all[n].innerHTML=objHTTP.responseText;
           objHTTP=null;
        }
      }
   }


   if (debugMode=='on') {
      alert(objDoc.xml);
      alert(document.all['mainpagediv'].innerHTML);
   }

   //doIFRAMES(objNodeList);
   document.all['currentUser'].value=currentUser;
   document.all['currentSite'].value=currentSite;

   var so = setTimeout("mouseOff()",1000);  // turns the mouse cursor to default pointer
   try {
      startTimer();
   } catch(e) { }
}


function removeListOptions(selectName) {
  for (var i=document.all[selectName].options.length-1; i>=0; i--) {
      document.all[selectName].options[i] = null;
  }
}

/* Loads list identified by pList using pQuery. Set pBlank=Y to insert blank row */
function loadList(pList, pQuery, pBlank) {
   objLDoc = new ActiveXObject("Microsoft.XMLDOM");
   objLDoc.async=false;
   objLDoc.load(pQuery);
   if (objLDoc.parsed) {
      var objDocList=objLDoc.getElementsByTagName("ROW");
   }
   else {
      return;
   }
   var i;
   var selectedIdx=-1;

   document.all[pList].options.length=0;
   try{
      for (i=0; i < objDocList.length; i++) {
         var childList=objDocList(i).childNodes;
         var objEntry = document.createElement("OPTION");
         objEntry.value=childList(0).text;
         objEntry.text=childList(1).text;
         document.all[pList].options.add(objEntry, i);
         try {
            if (childList(0).getAttribute("SELECTED")=="SELECTED") {
               selectedIdx=i;
            }
         } catch(e) {}
      }
   } catch(e) {
      return;
   }
   if (pBlank=='Y') {
     addBlankItem(pList)
     if (selectedIdx>-1) {
        selectedIdx=selectedIdx+1;
     }
   }
  if (selectedIdx>-1) {
     document.all[pList].options[selectedIdx].selected=true;
   }
   objLDoc=null;
}

// Write screen values back to xml-dom
function saveElems() {
   var objNodeList=objDoc.getElementsByTagName('widget');
   var widgettype;
   var nodename;
   var nodevalue;
   for (i=0; i < objNodeList.length; i++) {
     widgettype=objNodeList(i).getAttribute('type');
     try {
        nodename=document.all[objNodeList(i).getAttribute('name')].name;
        nodevalue=document.all[objNodeList(i).getAttribute('name').toLowerCase()].value;
        objNodeList(i).setAttribute('value', nodevalue);
     } catch(e) { null }
   }
}

function addBlankItem(listname) {
   var objEntry = document.createElement("OPTION");
   objEntry.value="-1";
   objEntry.text="";
   objEntry.selected=true;
   document.all[listname].options.add(objEntry, 0);
}

function closeWin() {
   try {
    /* changed by DSILK 1/26/2004
      onerror=errOut;
      var parentRef='opener.document.'+parent.name+'.style.border="2px outset"';
      eval(parentRef);
      parent.close();
    */
      var moduleNode=objDoc.getElementsByTagName('module');
      var moduleName=moduleNode(0).text;
      var parentRef='opener.document.' + moduleName + '.style.border="2px outset"';
      eval(parentRef);
      self.close();

      try{ localClose() } catch(e) { }
   }
   catch(e) { }
}

/* Pass me a reference to a HTML list and a string to search for. */
/* I'll find the string in the list and set the selectedIndex.    */
function setListItem(objList, strItem) {
  var listlength=objList.options.length -1;
  var i;
  for (i=0; i < listlength; i++) {
    if (objList.options(i).value==strItem) {
      objList.options(i).selected=true;
      break;
    }
  }
}


/* Clear all form fields in preparation for adding a new item */
/* Add this call to your Add button */
function addMode(tabname){
  var clearscreen=true;
  if (mode=='edit') {
     if ((dataDirtyflag==true)&&(currentPerm=='Y')) {
        var response = newConfirm('Please Confirm','Would you like to save your changes before adding a new item?',1,0,0);
       if (response) {

          try {
             var saveresult=localSaveProc();
          } catch(e) {}
          dataDirtyflag=false;
       }
       else { return; }
     }

     saveElems();

     mode='add';

     /* Implement locally if needed */
     try {
       clearscreen=oktoclear();
     } catch(e) { }

     if (clearscreen) {
        for (i=0; i < mainform.elements.length; i++) {
           try {
              if (mainform.elements[i].type != 'button') {
                 mainform.elements[i].value='';
              }
           } catch(e) { }
        }
     }
     try {
        formOnAdd(tabname)
     }
     catch(e) { }

     /* Implement locally if needed */
     try {
       afterclear();
     } catch(e) { }

   }
   else if (mode=='add') {
     if (dataDirtyflag==true) {
       var response = newConfirm('Please Confirm','You have not yet saved the item you were adding. Are you sure you want to go back to View/Edit mode?',1,0,0);
      if (response) { dataDirtyflag=false; }
      else { return; }
     }
      mode='edit';
      nodelist=objDoc.getElementsByTagName('widget');
      for (i=0; i < nodelist.length; i++) {
        if (nodelist[i].getAttribute('type') != 'button') {
           try {
        mainform.elements[nodelist[i].getAttribute('name').toLowerCase()].value = nodelist[i].getAttribute('value');
           } catch(e) { }
        }
      }
     formOnEdit(tabname);
   }
}


function closePrompt() {
   try {
      closeChildWindows();
   } catch(e) {}
   closeWin();
   if (dataDirtyflag==true) {
      event.returnValue = "Your changes have not been saved.";
   }
   try {
      if (vWinCode) {
        vWinCode.close();
      }
   } catch(e) {}
   try {
      parent.clickButton(parent.document.all[window.name]);
   } catch(e) {
      return;
   }
}

//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
// disable buttons separately for default style
//Old code commented by hirad5
function xxdisableField(control) {
   formItem=document.all[control];
   formItem.tabIndex=-1;
   if (formItem.type.toLowerCase()!="button" &&
       formItem.type.toLowerCase()!="checkbox" &&
        formItem.type.toLowerCase()!="radio" &&
         formItem.type.toLowerCase()!="image") {
         formItem.style.background="#eaeaea";
      if (formItem.type.toLowerCase()!="select-one" && formItem.type.toLowerCase()!="select-multiple") {
         formItem.readOnly=true;
      }
   } else {
      formItem.disabled=true;
      if (formItem.type.toLowerCase()=="image") {
         formItem.style.cursor="default";
      }
   }
}

//New Changed by Michael E. Harris on 09/20/2010
function disableField(control) {
   formItem=document.all[control];
   formItem.tabIndex=-1;
   if (formItem.type.toLowerCase()!="button" &&
       formItem.type.toLowerCase()!="checkbox" &&
        formItem.type.toLowerCase()!="radio" &&
         formItem.type.toLowerCase()!="image") {
         formItem.style.background="#eaeaea";
     if (formItem.type.toLowerCase()!="select-one" && formItem.type.toLowerCase()!="select-multiple") {
         formItem.readonly=true;
      }
   } else {
      if (formItem.type.toLowerCase()=="image") {
         formItem.style.cursor="default";
      }
   }
    formItem.disabled=true;
}

//Old code commented by hirad5
function xxenableField(control) { 
   try {
     formItem=document.all[control];
     addToTabIndex(formItem.name);
     if (formItem.type.toLowerCase()!="button" &&
          formItem.type.toLowerCase()!="checkbox" &&
           formItem.type.toLowerCase()!="radio" &&
             formItem.type.toLowerCase()!="image") {
        formItem.style.background="#ffffff";
        if (formItem.type.toLowerCase()!="select-one" &&
            formItem.type.toLowerCase()!="select-multiple") {
           formItem.readOnly=false;
        }
     } else {
        formItem.disabled=false;
         if (formItem.type.toLowerCase()=="image") {
            formItem.style.cursor="hand";
         }
     }
   } catch(e) {}
}

//////////////////////////////////////////////////////////////////////////////////
//New Changed by Michael E. Harris on 09/20/2010
function enableField(control) {
   try {
     formItem=document.all[control];
     addToTabIndex(formItem.name);
     if (formItem.type.toLowerCase()!="button" &&
          formItem.type.toLowerCase()!="checkbox" &&
           formItem.type.toLowerCase()!="radio" &&
             formItem.type.toLowerCase()!="image") {
        formItem.style.background="#ffffff";
        if (formItem.type.toLowerCase()!="select-one" &&
            formItem.type.toLowerCase()!="select-multiple") {
           formItem.readOnly=false;
        }
     } else {
        if (formItem.type.toLowerCase()=="image") {
            formItem.style.cursor="hand";
         }
     }
	formItem.disabled=false;
  } catch(e) {}
}

//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////


function onFocus(pControl) {
   formItem=document.all[pControl];
   try {
     if (formItem.type.toLowerCase()!="button") {
        if (formItem.type.toLowerCase()!="checkbox" &&
           formItem.type.toLowerCase()!="radio") {
           if (formItem.tabIndex==-1 || formItem.tabIndex=="-1") {
              //focusOnFirstTab();
              if (formItem.style.background=="#eaeaea") {
                 getNextElement(pControl);;
              }
           } else {
               document.all[pControl].style.backgroundColor="YELLOW";
           }
        } else {
            if (formItem.length>0) {
               for (f=0; f<formItem.length; f++) {
                  if (formItem[f].focus) {
                     formItem[f].style.background="yellow";
                  }
               }
            } else {
                formItem.style.background="yellow";
            }
        }
     }
   } catch(e) {}
   try{
      localonFocus(pControl);
   } catch(e) {}
}

function onBlur(pControl) {
   formItem=document.all[pControl];
   if (formItem.tabIndex!=-1 && formItem.tabIndex!="-1") {
      if (formItem.type.toLowerCase()!="button") {
         if (formItem.type.toLowerCase()!="checkbox" &&
               formItem.type.toLowerCase()!="radio") {
            document.all[pControl].style.backgroundColor="WHITE";
         } else {
            if (formItem.length>0) {
               for (b=0; b<formItem.length; b++) {
                 formItem[b].style.background="";
               }
            } else {
               formItem.style.background="";
            }
         }
      }
   }
   try{
      localonBlur(pControl);
   } catch(e) {}
   try{
      tString="";
   } catch(e) {}
}


function eDivOnFocus(pControl) {
   eDiv=document.all[pControl];
   eDiv.style.backgroundColor="YELLOW";
   try{
      localonFocus(pControl);
   } catch(e) {}
}

function eDivOnBlur(pControl) {
   eDiv=document.all[pControl];
   eDiv.style.backgroundColor="WHITE";
   try{
      localonBlur(pControl);
   } catch(e) {}
   try{
      tString="";
   } catch(e) {}
}

/*
   The variable array, fieldtypes, and the function focusOnFirstTab() are used in the onFocus
   and onBlur functions.  I added these to help control the tabIndex property of an object in
   the mainform collection.  - Dan Silk 05/21/2003
*/

function focusOnFirstTab(){
   objForm=document.mainform;
   var thisTab;
   for (x=0;x<objForm.length;x++) {
      thisTab=objForm(x).tabIndex;
      thisType=objForm(x).type;
      if (objForm(x).tabIndex>-1) {
         for (i=0;i<fieldtypes.length;i++) {
            if (thisType==fieldtypes[i]) {
               if (thisTab==0) {
                  if (isWidgetVisible(objForm(x).name)) {
                     objForm(x).focus();
                     //alert(objForm(x).name);
                     return;
                  }
               }
            }
         }
      }
   }
}



function setFieldIndex(fieldSet) {
   var formItems=document.mainform;
   for (x=0; x<formItems.length; x++) {
      for (i=0; i<fieldtypes.length; i++) {
         if (formItems(x).type==fieldtypes[i]) {
            disableField(formItems(x).name);
         }
      }
   }
   var thisField="";

      for (y=0; y<fieldSet.length; y++) {
         thisField=fieldSet[y];
         enableField(thisField);
      }
   try {
      document.all[fieldSet[0]].focus();
      if (getWidgetAttrib(fieldSet[0],"type")=="input") {
         var wi=fieldSet[0];
         document.all[wi].select();
      }
   } catch(e) {}
   try{
      localFieldIndex();
   } catch(e) {}
}

function getNextTabIndex(){
   objForm=document.mainform;
   var maxTab=-2;
   var thisTab;
   for (x=0;x<objForm.length;x++) {
      thisTab=objForm(x).tabIndex;
      if (thisTab>maxTab) {
         maxTab=thisTab;
      }
   }
   if (maxTab>0) {
      maxTab=maxTab+1;
   }
   if (maxTab<0) {
      maxTab=0;
   }
   return maxTab;
}

function addToTabIndex(formObj) {
   thisObj=document.all[formObj];
   //thisObj.tabIndex=getNextTabIndex();
   thisObj.tabIndex=1;
}

function onHelp(pControl) {
  alert('help')
}

function onPickWin(pControl) {
  alert('help')
}

function onKeypress(pControl) {
  if (document.all[pControl].value=='?') {
    alert('help');
  }
}

//Function so show only the buttons that belong on a tag
function showTabButtons(tab){
   var iNode=0;
   var tabName=tab.name;
   onerror=errOut;
   tabName=tabName.substring(0, 4);
   var objNodeList = objDoc.getElementsByTagName('widget');
   for (iNode=0; iNode < objNodeList.length; iNode++) {

     if ( objNodeList(iNode).getAttribute('type')=='button' &&
         objNodeList(iNode).getAttribute('tab_name')!='ALL' &&
         objNodeList(iNode).getAttribute('tab_name')!=null)
     {
       if(objNodeList(iNode).getAttribute('tab_name')==tabName)
       {
         //alert(objNodeList(iNode).getAttribute('style'));
         objNodeList(iNode).setAttribute('style','display="inline"');
        }
        else {
     objNodeList(iNode).setAttribute('style','display="none"');
        }
      }
   }
}

function errOut(errMsg,location,lineNum){
   alert('error is :'+ errMsg);
   return(false);
}

//button click function same as in vm_dialogs for closing buttons
function clickButton(pbutton) {
  if (pbutton.style.border.indexOf('outset') > -1) {
    pbutton.style.border='2px inset';
  }
  else {
    pbutton.style.border='2px outset';
  }
}


// Generates an XML document containing db key elements.
function genKeyXML() {
   var xmlDoc='<KEYLIST>';
   var widgettype;    //GSM ADDED
   var j;         //GSM ADDED
   var keyLoopVal;    //GSM ADDED
   objRep = new ActiveXObject("Microsoft.XMLDOM");
   objRep.async=false;
   objRep.loadXML('<?xml version="1.0"?>'+document.all['mainpagexml'].xml);
   if (objDoc.parsed) {
      var objKeyList=objRep.getElementsByTagName("widget");
   }
   else {
      alert('em_dialog.runReport failed to parse query for report: '+pReportNo);
      return;
   }

   for (i=0; i < objKeyList.length; i++) {
      widgettype=objKeyList(i).getAttribute('type');
      checked=objKeyList(i).getAttribute('checked');
      //GSM ADDED
      multipleSel=objKeyList(i).getAttribute('multiple');
      selectedVal=objKeyList(i).getAttribute('selected');

      if (objKeyList(i).getAttribute('keyval')=='Y') {
                //GSM ADDED SECTION - RADIO BUTTONS
      if (widgettype == 'radio') {
         if (keyLoopVal != objKeyList(i).getAttribute('name')){
                        keyLoopVal = objKeyList(i).getAttribute('name');
                        j = 0;
         }

          if (document.all[objKeyList(i).getAttribute('name').toLowerCase()][j].checked) {
            xmlDoc = xmlDoc + '<KEY><'+objKeyList(i).getAttribute('name').toUpperCase() + '>';
                 xmlDoc = xmlDoc + document.all[objKeyList(i).getAttribute('name').toLowerCase()][j].value +
            '</'+objKeyList(i).getAttribute('name').toUpperCase()+'></KEY>';
                        }
                        j ++;
        }

        //GSM ADDED SECTION - MULIPLE SELECTS
        else if (widgettype == 'select' && multipleSel == 'multiple') {

         if (keyLoopVal != objKeyList(i).getAttribute('name')){
                        keyLoopVal = objKeyList(i).getAttribute('name');
         }

            xmlDoc = xmlDoc + '<KEY><'+objKeyList(i).getAttribute('name').toUpperCase() + '>';

            for (t=0; t < document.all[keyLoopVal].length; t++) {
               if (t==0) {
                  xmlDoc = xmlDoc + document.all[objKeyList(i).getAttribute('name').toLowerCase()][t].value;
               }
                else{
                   xmlDoc = xmlDoc + "," + document.all[objKeyList(i).getAttribute('name').toLowerCase()][t].value;
                }
                }

                xmlDoc = xmlDoc + '</'+objKeyList(i).getAttribute('name').toUpperCase()+'></KEY>';
        }

        // THIS SECTION IS CALLED FOR ALL OTHER WIDGETS
        else {

         //try {
           xmlDoc = xmlDoc + '<KEY><'+objKeyList(i).getAttribute('name').toUpperCase() + '>';
           xmlDoc = xmlDoc + document.all[objKeyList(i).getAttribute('name').toLowerCase()].value +
         '</'+objKeyList(i).getAttribute('name').toUpperCase()+'></KEY>';
         /*} catch(e) {
           alert(objKeyList(i).getAttribute('name'));
         }*/
      }
      }
   }

   xmlDoc = xmlDoc + '</KEYLIST>';
   //alert(xmlDoc);
   objRep=null;
   return xmlDoc;
}

// Generates an XML document containing db key elements.
function genKeySubsetXML() {
   var xmlDoc='<KEYLIST>';  
   var widgettype;    //GSM ADDED
   var j;         //GSM ADDED
   var keyLoopVal;    //GSM ADDED
   objRep = new ActiveXObject("Microsoft.XMLDOM");
   objRep.async=false;

   var thisList = new Array();
   var counter;
   var currList;

   for (i=0; i < swil_all_areas.length; i++) {
      try {
         if (document.all[swil_all_areas[i]].style.display != "none") {
            currList = eval ('swil_' + swil_all_areas[i] + '_keylist');

            if (currList.length > 0) {
               counter = thisList.length;
               for (y=0; y < currList.length; y++ ) {
                  thisList[counter] = currList[y];
                  counter++;
               }
            }
         }
      } catch(e){}
   }

   objRep.loadXML('<?xml version="1.0"?>'+document.all['mainpagexml'].xml);
   if (objRep.parsed) {
      var objKeyList=objRep.getElementsByTagName("widget");
   }
   else {
      alert('em_dialog.runReport failed to parse query for report: '+pReportNo);
      return;
   }

   for (i=0; i < objKeyList.length; i++) {

   	if(checkKeyWidget(thisList, objKeyList(i).getAttribute('name'))) {

      widgettype=objKeyList(i).getAttribute('type');
      checked=objKeyList(i).getAttribute('checked');
      //GSM ADDED
      multipleSel=objKeyList(i).getAttribute('multiple');
      selectedVal=objKeyList(i).getAttribute('selected');

      if (objKeyList(i).getAttribute('keyval')=='Y') {
                //GSM ADDED SECTION - RADIO BUTTONS
      	if (widgettype == 'radio') {
			 if (keyLoopVal != objKeyList(i).getAttribute('name')){
							keyLoopVal = objKeyList(i).getAttribute('name');
							j = 0;
			 }

			 if (document.all[objKeyList(i).getAttribute('name').toLowerCase()][j].checked) {
				xmlDoc = xmlDoc + '<KEY><'+objKeyList(i).getAttribute('name').toUpperCase() + '>';
				xmlDoc = xmlDoc + document.all[objKeyList(i).getAttribute('name').toLowerCase()][j].value +
				'</'+objKeyList(i).getAttribute('name').toUpperCase()+'></KEY>';
			 }
			 j ++;
        }
        //GSM ADDED SECTION - MULIPLE SELECTS
        else if (widgettype == 'select' && multipleSel == 'multiple') {

			 if (keyLoopVal != objKeyList(i).getAttribute('name')){
					keyLoopVal = objKeyList(i).getAttribute('name');
			 }

            xmlDoc = xmlDoc + '<KEY><'+objKeyList(i).getAttribute('name').toUpperCase() + '>';

			counter = 0; //jesse 
            for (t=0; t < document.all[keyLoopVal].length; t++) {
				if(document.all[objKeyList(i).getAttribute('name').toLowerCase()][t].selected){ 
					//jesse 2019.04.19 only add the items selected 
				    
				   if (counter == 0){ //first item selected 
					  xmlDoc = xmlDoc + document.all[objKeyList(i).getAttribute('name').toLowerCase()][t].value;
				   }else{
					   xmlDoc = xmlDoc + "," + document.all[objKeyList(i).getAttribute('name').toLowerCase()][t].value;
				   }
				   counter++; 
				   /*  jesse 2019.04.19 
				   if (t==0) {
					  xmlDoc = xmlDoc + document.all[objKeyList(i).getAttribute('name').toLowerCase()][t].value;
				   }else{
					   xmlDoc = xmlDoc + "," + document.all[objKeyList(i).getAttribute('name').toLowerCase()][t].value;
				   }*/
				}
            }

            xmlDoc = xmlDoc + '</'+objKeyList(i).getAttribute('name').toUpperCase()+'></KEY>';
			
        } else {
		// THIS SECTION IS CALLED FOR ALL OTHER WIDGETS
           xmlDoc = xmlDoc + '<KEY><'+objKeyList(i).getAttribute('name').toUpperCase() + '>';
           xmlDoc = xmlDoc + document.all[objKeyList(i).getAttribute('name').toLowerCase()].value +
         '</'+objKeyList(i).getAttribute('name').toUpperCase()+'></KEY>';
      	}
      } // end of if (objKeyList(i).getAttribute('keyval')=='Y') 
	} //end of if(checkKeyWidget(thisList, objKeyList(i).getAttribute('name')))

   }//end of for loop 

   xmlDoc = xmlDoc + '</KEYLIST>';
   //alert(xmlDoc);
   objRep=null;
   return xmlDoc;
}

function checkKeyWidget(thisList, widgetName) {

    for (x=0; x<thisList.length ;x++) {

            if (thisList[x] == widgetName) {

            return(true);
            }
    }
    return(false);
}

function runReport(pReportID) {
   var xmlKeyString=genKeySubsetXML();
   //	alert('xmlKeyString='+xmlKeyString); 
	
    /* alert('em_show_report?pReportID='+pReportID+'&pParamXML='+xmlDoc);  */
   //showReport('em_show_report?pReportID='+pReportID+'&pParamXML='+xmlKeyString);
   showReport('em_show_report?pReportID='+pReportID+'&pUser='+currentUser+'&pSite='+currentSite+'&pParamXML='+xmlKeyString);
}


/* Launches the report evoked from the report selection menu into a new window */
function showReport(pReportURL) {
  var maxX = 280
  var maxY = 150
  var scrollbartext = ""
  //var newWindow = null
   var winleft = (screen.width / 2) - (maxX / 2);
   var wintop = (screen.height / 2) - (maxY / 2);

  try {
     if (newWindow) {
        newWindow.close();
     }
  } catch(e) {}

  if (screen.width < maxX)
     maxX = screen.width-20
  if (screen.height < maxY)
  {
     maxY = screen.height-50
     scrollbartext = ",scrollbars=yes"
  }
  //window.open(pReportURL, 'prefwin', 'resizeable=1,status=no,width=' + maxX + ',height=' + maxY + scrollbartext)
  newWindow = window.open(pReportURL, 'prefwin', 'resizeable=1,status=no,width=' + maxX + ',height=' + maxY + scrollbartext + ',top=' + wintop + ',left=' + winleft);
  //newWindow = window.open(pReportURL);
  newWindow.focus();
}

/*http://emarss58.eng.usps.gov/vmdev1/!em_run_report?pApplication=EMARS&pReportType=REPORT60&p_report=sch004.rdf
  &p_W_WEEK_IND_DD=this&p_W_DAILY_TOURLY=C&p_W_SENIOR=C&p_W_CREW_RAD=A&p_EXEC_DATE=05/09/201913:57:35
  &p_W_ASSIGN_DAY=13-MAY-2019,14-MAY-2019,15-MAY-2019&p_W_ASSIGN_CREW=001,002,003&p_SEPPAGES=Y&p_WHOLECREW=Y&p_site_id=490&p_user_id=3748&email=N&desformat=PDF
 
  http://emarss58.eng.usps.gov/vmdev1/!em_run_report?pApplication=EMARS&pReportType=REPORT60&p_report=sch003.rdf
  &p_EXEC_DATE=05/09/201914:37:51&p_SEPPAGES=Y&p_WHOLECREW=N&p_SCHED_DATE=05/09/2019&p_EMP_ID=47339&p_CREW_ID=001
  &p_EMP_NAME=ADAMS,*space*W*space*W................0005..ON........*space*X&p_MOVEMODE=move&p_site_id=490&p_user_id=3748&email=N&desformat=PDF

*/


function calproc(widgetname){
   var elem=event.srcElement;
   if (eval("document."+widgetname+".readOnly") || eval("document."+widgetname+".disabled")) {
      elem.style.cursor="default";
   } else {
      elem.style.cursor="hand";
      show_calendar(widgetname);
   }
}

/* Call this for all SAVE buttons */
function saveForm(pTab) {
   if (validateRequired(pTab)==-1) {
      return;
   }

   try {
     var module=objDoc.getElementsByTagName('module');
     var mod=module(0).text;
   }

   catch(e) { alert('Failed to get ModuleID. Data was not saved'); return; }

   var vURL='module='+mod+'&mode='+mode+'&tab='+pTab+'&';

   try {     beforeSave(pTab);
   } catch(e) { }

   var widgetNodes=objDoc.getElementsByTagName("widget[@type='select'||@type='input'||@type='checkbox'||type='textarea'||@type='hidden'||@type='radio'||@type='label'||@type='division']");
   var widget_name="";
   var widget_type="";
   var valText;
   var formElem;
   var bDontPassMe=false;

   for (i=0; i < widgetNodes.length; i++) {
      widget_name=widgetNodes[i].getAttribute("name");
      widget_type=widgetNodes[i].getAttribute("type");
      formElem=document.all[widget_name];                // added by ds for radio loop below
      bDontPassMe=false;
      try {
         if (widget_type=="radio") {
             valText=document.all[widget_name][0].checked;

         } else if (widget_type=="division") {
            if (widgetNodes[i].getAttribute("editable")=="Y") {
               valText=URLEncode(document.all[widget_name].innerHTML);
            } else {
              bDontPassMe=true;
            }
         } else {
            valText=URLEncode(document.all[widget_name].value);
         }
         if (bDontPassMe==false) {
            vURL=vURL+widget_name+'='+valText+'&';
         }

      } catch(e) {
         if (debugMode=="on") {
            alert("saveForm error: " + e.message);
         }
      }
   }
   vURL=vURL.substring(0, vURL.length -1);
   vURL="!vm_main_pkg.save?pSiteID="+currentSite+"&pUser="+currentUser+"&"+vURL;

   if (debugMode=='on') {
      prompt("debugMode is on: ",vURL);
   }

   objSave=getXMLresults(vURL);
   if (objSave.parsed) {
      message=objSave.getElementsByTagName('message');
      dataDirtyflag = false;
      var saveStatus=message(0).getAttribute('status');
      var newValue=message(0).getAttribute('newvalue');
      var auxText=message(0).getAttribute('auxtext');
      if (gmessagesOn) {
         alert(message(0).text);
      }
   }
   try {
      save(pTab, saveStatus, newValue, auxText);
   } catch(e) { }
}


function info() {
   var conv=document.all.mod_user.value;
   if ( document.all['mod_user'].value.indexOf('UNKNOWN') > -1 ) {
      conv='CONVERSION';
   }
   alert('last modified: '+document.all['mod_date'].value+' by user: '+conv)
}


/* Validates form then passes control to local save() proc */
function validateRequired(tag){
   if ((tag=="")||(tag==null)) {
      reqlist=objDoc.getElementsByTagName('widget[@required="Y"]');
   } else {
      reqlist=objDoc.getElementsByTagName('widget[@required="Y"&&@tag="'+tag+'"]');
   }

   for (iReq=0; iReq < reqlist.length; iReq++){
      var widgetname=reqlist(iReq).getAttribute('name').toLowerCase();
      var widgetlabel=reqlist(iReq).getAttribute('label');
      var widgetmode=reqlist(iReq).getAttribute('addmode');
      var tagname=reqlist(iReq).getAttribute('tag');
      var req=reqlist(iReq).getAttribute('required');
      if (mainform.elements[widgetname].type.indexOf('select') > -1) {
          if ((document.all[widgetname].value==-1)||(document.all[widgetname].value=="")) {
             if ((widgetmode=='Y')&&(mode.toLowerCase()=='edit')) {
           null;
             } else {
                if (isWidgetVisible(widgetname)) {
                   alert('Entry: '+widgetlabel+' is required. Please complete this field, then try again.');
                   document.all[widgetname].focus();
                   return -1;
                }
             }
           }
       }
       else if(mainform.elements[widgetname].type="text") {
     if (document.all[widgetname].value=="") {
             if ((widgetmode=='Y')&&(mode.toLowerCase()=='edit')) {
           null;
             } else {
                if (isWidgetVisible(widgetname)) {
                   alert('Entry: '+widgetlabel+' is required. Please complete this field, then try again.');
                   document.all[widgetname].focus();
                   return -1;
                }
             }
           }
       }
    }
}


function isWidgetVisible(widgetname) {
   try {
      if (document.all[widgetname].tabIndex!=-1) {
         document.all[widgetname].focus();
         return true;
      } else {
         return false;
      }
   }
   catch(e) { return false; }
}


function leadZero(string_in) {
           if (string_in.length < 2) {
             return "0" + string_in
           }
           else {
             return string_in
           }
         }

function sessionDate(date_in) {
           var v_MM  = date_in.getMonth()+1
           var v_DD  = date_in.getDate()
           var v_Hrs = date_in.getHours()
           var v_Min = date_in.getMinutes()
      var v_Sec = date_in.getSeconds()
           return leadZero(v_MM.toString()) + "/" + leadZero(v_DD.toString()) + "/" + date_in.getYear()
                  + leadZero(v_Hrs.toString()) + ":" + leadZero(v_Min.toString())+ ":" + leadZero(v_Sec.toString())
         }


function getNextElement(fieldName) {
   var formItems=document.mainform;
   var thisIndex=0;
   var field = formItems[fieldName];

   for (formInt=0;formInt<formItems.length;formInt++) {
      try{
         if (formItems.elements[formInt].name==field.name){
            thisIndex=formInt+1;
         }
      } catch(e) { }
   }
   if (thisIndex>formItems.length-1) {
      thisIndex=0;
   }
   for (focusInt=thisIndex;focusInt<formItems.length;focusInt++) {
      if (checkFieldType(formItems.elements[focusInt].name)) {
         if (isWidgetVisible(formItems.elements[focusInt].name)){
            break;
         } else {
            if (focusInt==formItems.length-1) {
               field.focus();
            }
         }
      }
   }
}

function checkFieldType(widget) {
   var thisType=getWidgetAttrib(widget,"type");
   if (thisType=="division") {
      if (getWidgetAttrib(widget,"editable")=="Y") {
         return true;
      }
   }
   var typeAR=new Array("button","checkbox","imagebutton","input","radio","select");
   for (x=0; x<typeAR.length; x++) {
      if (thisType==typeAR[x]) {
         return true;
      }
   }
   var formTypesAR=new Array("button","checkbox","input","radio","select-one","select-multiple");
   for (x=0; x<formTypesAR.length; x++) {
      if (document.all[widget].type==formTypesAR[x]) {
         return true;
      }
   }
   return false;
}

function tabOnEnter(field,evt) {
   var keyCode = document.layers ? evt.which : document.all ? evt.keyCode : evt.keyCode;

   if (keyCode != 13) {
      return true;
   } else {
      try { getNextElement(field.name); } catch(e) {}
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


function highlightOnChange(control) {
    document.all[control].style.background  = '#f7f78c';
    document.all[control].style.fontWeight = 'bold';
}

function removeHighlight(control) {
    document.all[control].style.background  = '#ffffff';
    document.all[control].style.fontWeight = 'normal';
}



function dummy() { }

function mouseOff() {
   document.body.style.cursor = 'default';
}

function mouseOn() {
   document.body.style.cursor = 'wait';
}

function whatever() {
//;askldjfa;slkdjf

var so = setTimeout("mouseOff()",500);

}


// ------------------------------------------------------
// URL encoding  next 5 functions
// ------------------------------------------------------
var hexVals = new Array("0","1","2","3","4","5","6","7","8","9","A","B","C","D","E","F");
var unsafeString = "&'/+-,";

function URLDecode(val)
{
   var returnstr=unescape(val);
   return returnstr;
        // while coding i found that IE had problem writing '<form>' to innerhtml. </form> was ok. 'unknown runtime error' IE5.5.
}

function URLEncode(val)
{
   lStr = new String(val); //Robert Dinion  changed to a use a String as some browsers wouldn't get val.length 
   var state   = 'urlenc';
   var len     = lStr.length;
   var backlen = len;
   var newStr  = "";
   var frag    = "";
   var encval  = "";

   for (u=0;u<len;u++) {
      frag=lStr.substring(u,u+1);
      if (isURLok(frag)) {                          // comment this line to encode everything
         newStr = newStr + frag;                    // comment this line to encode everything
      } else {                                      // comment this line to encode everything
         encval=decToHex(frag.charCodeAt(0),16);
         newStr = newStr + "%" + encval;
      }                                             // comment this line to encode everything
   }

   if (debugMode=="on") {
      alert("full value is: " + val + "\nvalue length: " + len + "\nnew val is: " + newStr);
   }

   return newStr;
}

function decToHex(num, radix) // part of URL Encode
{
   var hexString = "";
   while (num >= radix) {
      temp = num % radix;
      num = Math.floor(num / radix);
      hexString += hexVals[temp];
   }
   hexString += hexVals[num];
   return reversal(hexString);
}

function reversal(s) // part of URL Encode
{
   var len = s.length;
   var trans = "";
   for (rev=0; rev<len; rev++) {
      trans = trans + s.substring(len-rev-1, len-rev);
   }
   s = trans;
   return s;
}

function isURLok(compareChar) // part of URL Encode
{
   var rePattern = new RegExp("([0-9]|[a-z]|[A-Z])","g");   //regular expression, matches only numbers or letters

//   if (unsafeString.indexOf(compareChar) == -1 && compareChar.charCodeAt(0) > 32 && compareChar.charCodeAt(0) < 123)
   var matches=compareChar.match(rePattern);

   if (matches!=null) {
      return true;
   } else {
      return false;
   }
}

// ------------------------------------------------------
// end of URL encode and decode functions
// ------------------------------------------------------
// ------------------------------------------------------


// This function uses XMLhttpRequest to get a blob of HTML from an Oracle pkg.procedure
// Be sure to add any parameters to the procedure like "pkg.procedure?param1=this&param2=that"
function getInnerHTML(tagID,vURL) {
   mouseOn();
     document.all[tagID].innerHTML=null;
     var objHTTP=new ActiveXObject("microsoft.XMLHTTP");
     objHTTP.open('POST', getServerPath()+vURL, false);
     objHTTP.send();
     document.all[tagID].innerHTML=objHTTP.responseText;
     objHTTP=null;

   mouseOff();
}

function getMyInnerHTML(tagID,HTMLProc) {
mouseOn();
var xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
xmlhttp.open("GET", getServerPath()+HTMLProc);

var myXML='';

 xmlhttp.onreadystatechange=function() {
  if (xmlhttp.readyState==4) {
    myXML=xmlhttp.responseText;
    if (myXML!=null) {
       document.all[tagID].innerHTML=myXML;
    }
    mouseOff();
  }
 }
 xmlhttp.send(null)
 //alert(myXML);
}


function execHTTP(vURL) {
   var objHTTP=new ActiveXObject("microsoft.XMLHTTP");
   var objDoc = new ActiveXObject("Microsoft.XMLDOM");
   objHTTP.open('POST', getServerPath()+vURL, false);
   objHTTP.send();
   var result = objHTTP.responseText;
   if ((result==null)||(result.indexOf('No data returned')) > -1) { return }
   objDoc.loadXML(result);
   if (objDoc.parseError) { return result }
   var msg=objDoc.selectSingleNode('message');
   if (msg.getAttribute('status')=='ERROR') {
      if (msg.text.length == 0) {
         alert('An error occurred which was not caused by you. The error number is '+msg.getAttribute("errNo")
                + '.  System administrators have been notified.');
      } else {
         alert('HERE'+msg.text);
      }
      return;
   }
   return result;
}

function handOn(curObj) {
   curObj.style.cursor='hand';
}

function handOff(curObj) {
   curObj.style.cursor='default';
}

function doIFRAMES() {
   iFrames=objDoc.getElementsByTagName("widget[@type='iframe']");
   iFramesLen=iFrames.length;
   for (iNode=0; iNode < iFramesLen; iNode++) {
      try {
         var widgetName=iFrames(iNode).getAttribute("name");
         var src=iFrames(iNode).getAttribute("src");
         if (src.length>0) {
            var frameborder=iFrames(iNode).getAttribute("frameborder");
            var width=iFrames(iNode).getAttribute("width");
            var height=iFrames(iNode).getAttribute("height");
            var style=iFrames(iNode).getAttribute("style");
            var title = iFrames(iNode).getAttribute("title");
            var tabIndex = iFrames(iNode).getAttribute("widget_tabindex");
            var iFrameCode = "<IFRAME id=\"" + widgetName +
                             "\" name=\"" + widgetName +
                             "\" src=\"" + src +
                             "\" width=\"" + width +
                             "\" height=\"" + height +
                             "\" frameborder=\"" + frameborder +
                             "\" style=\"" + style +
                             "\" />";
            var divID="div_"+widgetName;
            document.all[divID].innerHTML=iFrameCode;
          }
      } catch(e) { }
   }
}

function reloadIframe(widgetName,urlStr) {
   document.all[widgetName].src=urlStr;
}


function getMDY(dateStr,MorDorY){
   var dateAR = dateStr.split("/");
   if (MorDorY=="M") {
      return dateAR[0];
   }
   if (MorDorY=="D") {
      return dateAR[1];
   }
   if (MorDorY=="Y") {
      return dateAR[2];
   }
}

function compareDates(controlDateStr,controlTimeStr,compareDateStr,compareTimeStr) {
   // getMDY splits date str in format mm/dd/yyyy to get num vals for MM, DD and YYYY
   var yr = getMDY(compareDateStr,"Y");
   var mo = getMDY(compareDateStr,"M");
   var da = getMDY(compareDateStr,"D");
   var hr = 0;
   var mi = 0;

   if (compareTimeStr.length>0) {
      var comTimeAR = compareTimeStr.split(":");
      hr = eval(comTimeAR[0]);
      mi = eval(comTimeAR[1]);
   }
   // objCompareDate is created with MM, DD, YYYY of compareDateStr
   objCompareDate = new Date(eval(yr),eval(mo)-1,eval(da),hr,mi,0);

   hr=23;
   mi=59;

   if (controlTimeStr.length>0) {
      var conTimeAR = controlTimeStr.split(":");
      hr = eval(conTimeAR[0]);
      mi = eval(conTimeAR[1]);
   }
   if (controlDateStr=="now" || controlDateStr==null || controlDateStr=="") {
      setDateTime();
      objNow=new Date(currentDate);
      yr = objNow.getYear();
      mo = objNow.getMonth();
      da = objNow.getDate();
      objControlDate = new Date(currentDate + " "+hr+":"+mi);
   } else {
      yr = getMDY(controlDateStr,"Y");
      mo = getMDY(controlDateStr,"M");
      da = getMDY(controlDateStr,"D");
      if (controlTimeStr.length>0) {
         var conTimeAR = controlTimeStr.split(":");
         hr = eval(conTimeAR[0]);
         mi = eval(conTimeAR[1]);
      }
      // objCompareDate is created with MM, DD, YYYY of compareDateStr
      objControlDate = new Date(eval(yr),eval(mo)-1,eval(da),hr,mi,0);
   }
   var dateDiff=0;
   dateDiff=objControlDate-objCompareDate;
   var days = dateDiff/(1000*60*60*24);
   return days;
}


//Use this function to check that widget date1 is before widget date2...
function check_dates(strDateWidget1, strDateWidget2) {

   if ( document.all[strDateWidget1].value == '')
     {
       alert('Please Enter A Valid Begin Date');
       document.all[strDateWidget1].focus();
       return false;
     }

   if ( document.all[strDateWidget2].value == '')
     {
       alert('Please Enter A Valid End Date');
       document.all[strDateWidget2].focus();
       return false;
     }

   var bgn_date = new Date(document.all[strDateWidget1].value);
   var end_date = new Date(document.all[strDateWidget2].value);

   if ( bgn_date > end_date )
     {
       alert('Begin Date Greater than End Date');
       document.all[strDateWidget1].focus();
       return false;
     }

   return true;
}

function getXMLresults(pQuery) {
   objResults = new ActiveXObject("Microsoft.XMLDOM");
   objResults.async=false;
   objResults.load(getServerPath()+pQuery);
   return objResults;
}


// taken from doWidgets.  A separate function used to be generated for each module_to_evoke.
function launchModalModule(pModule,mWidth,mHeight) {
   var keyxml=genKeyXML();

   var urlStr="!vm_main_pkg.main?pUser=" + currentUser + "&pSiteID=" + currentSite + "&pModule=" + pModule + "&pparamxml="+keyxml+"";
   var winProps="dialogWidth:" + mWidth + "px; dialogHeight:" + mHeight + "px; center:yes; status:no; resizeable:no; scroll:no;";

   if (pModule=="EM_REACT_EQ_FINDER") {
      prompt("url",urlStr);
   }

   var v_modalWin=showModalDialog(urlStr, pModule, winProps);
   try { doModalResult(v_modalWin); } catch(e) {}
}

// same as launchModalModule(), may want to delete
function launchModalFinder(pModule,mWidth,mHeight) {
   var keyxml=genKeyXML();
   var v_finderWin=showModalDialog("!vm_main_pkg.main?pUser=" + currentUser + "&pSiteID=" + currentSite + "&pModule=" + pModule + "&pparamxml="+keyxml+"", "", "dialogWidth:" + mWidth + "px; dialogHeight:" + mHeight + "px; center:yes; status:no; resizeable:no");
   try { doModalResult(v_finerWin) } catch(e) {}
}

function newConfirm(title,mess,icon,defbut,mods) {

   icon = (icon==0) ? 0 : 2;
   defbut = (defbut==0) ? 0 : 1;
   retVal = makeMsgBox(title,mess,icon,4,defbut,mods);
   retVal = (retVal==6);

   return retVal;
}

function cloneselect(sourceName,targetName) {
   source=document.all[sourceName];
   target=document.all[targetName];
   for(i=0; i<(source.options.length); i++) {
      currentElement = new Option(source.options[i].text, source.options[i].value)
      target.options[i] = currentElement;
   }
}

function noPerm() {
   alert("You do not have permission to access this feature.");
}

function resetSwilAreas() {
   for (aInt=0; aInt<swil_all_areas.length; aInt++) {
      document.all[swil_all_areas[aInt]].style.display="none";
   }
   for (dInt=0; dInt<swil_default_areas.length; dInt++) {
      document.all[swil_default_areas[dInt]].style.display="block";
   }
}

//cprice 09-28-2004 changed this function to operate as expected.
var tString = "";
function typeAhead() {
      // Global Variables
      var i = 0;
      var success = false;
      var elem = event.srcElement;
      var tLowElemText = "";
      var rExpr = "";

      // Get the unicode char of the keypress
      var eCode = event.keyCode;

      // Check if it's a vaid ASCII Character
      if (eCode == 27 || eCode == 63){
         tString = "";
         elem.selectedIndex = 0;
         try {
            elem.onchange();
         } catch(e) {}
      } else if ( (eCode >= 32) && (eCode <= 122)) {
        // Convert the Code to the corresponding character and add to searchstring
        tString += String.fromCharCode(eCode);
        // ... and perform the search starting from the top element in the listbox
          while (success == false) {
            i = 0;
            // Convert everything to lowercase; allows an easy comparison
            var tLowString = tString.toLowerCase();

            // Compose the regexp searchstring ...
            rExpr = eval("/^" +  tLowString + "/");

            while ((i < elem.length)&& (success == false) )  {
              tLowElemText = elem.options[i].text.toLowerCase();
              // success: Position the listbox on the (first) found element
              if (tLowElemText.search(rExpr) != -1) {
                  elem.selectedIndex = i;
                  try {
                    elem.onchange();
                  } catch(e) {}
                  success = true;
               } else {
                  i++;
               }
            }
            // if nothing is found in the entire list, the last character of the searstring is
            // removed to allow typing the correct "next" character
            if (success == false) {
               tString = tString.substr(0, tString.length-1);
            }
         }
      }

      window.event.returnValue = false;
      window.event.cancelBubble = true;
}

function getWidgetAttrib(widgetName,attr) {
   var objNodeList = objDoc.getElementsByTagName("widget");
   for (iNode=0; iNode < objNodeList.length; iNode++) {
      try {
         if (objNodeList(iNode).getAttribute("name")==widgetName) {
            return objNodeList(iNode).getAttribute(attr);
         }
      } catch(e) {}
   }
   return;
}

/* The purpose of this function is to keep tabs on all extra SWIL windows that
   are opened by adding the window object name to the childWindows array.   */
function openSwilWin(pModule,objName)
{
   if (objName==null || objName.length==0) {
      objName=pModule;
   }
   var winSize=getXMLresults("EM_XML_MODULE_SIZE_PRC?p_screen_id="+pModule);
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
   var urlStr="!vm_main_pkg.main?pUser=" + currentUser + "&pSiteID=" + currentSite + "&pModule="+pModule;
   var winleft = (screen.width / 2) - (w/2);
   var wintop = (screen.height / 2) - (h/2);
   var winProps="left="+winleft+",top="+wintop+",width="+w+",height="+h+",status=no,menubar=no,toolbar=no,resizable=no,scrollbars=no";
   var exeStr;
   try  {
	   //jesse 1/10/2011 
	   // For some versions of IE, window.close() does not really set the window object to be NULL. Even the window is already closed, 
	   // the object  is active. So, eval(theObj.focus()) is still executable without any problems. 
	   // The following codes will make sure the previous windiw (if appliable) is closed, and create a new one. 
	  exeStr=objName+".close()";
      eval(exeStr);
	  exeStr=objName+"=window.open(\""+urlStr+"\",\""+pModule+"\",\""+winProps+"\")";
	  eval(exeStr); 
	  //jesse 1/10/2011 
	  
      exeStr=objName+".focus()";
      eval(exeStr);
   } catch(e) {
      exeStr=objName+"=window.open(\""+urlStr+"\",\""+pModule+"\",\""+winProps+"\")";
      eval(exeStr);
      childWindows[childWindows.length]=objName;
   }
}

/* The purpose of this function is to keep tabs on all non-SWIL windows that
   are opened by adding the window object name to the childWindows array.   */
function openNonSwilWin(objName,urlStr,winProps)
{
   winProps=winProps.toLowerCase();
   winPropsAR=winProps.split(",");
   for (x=0;x<winPropsAR.length;x++) {
      exeStr=winPropsAR[x].replace("=","='")+"'";
      eval("var this"+exeStr);
   }
   var wintop = (screen.height / 2) - (thisheight/2);
   var winleft = (screen.width / 2) - (thiswidth/2);
   if (winProps.indexOf("left") == -1) {
      winProps="left="+winleft+","+winProps;
   }
   if (winProps.indexOf("top") == -1) {
      winProps="top="+wintop+","+winProps;
   }
   var exeStr;
   try  {
      exeStr=objName+".focus()";
      eval(exeStr);
   } catch(e) {
      exeStr=objName+"=window.open(\""+urlStr+"\",\""+objName+"Win\",\""+winProps+"\")";
      eval(exeStr);
      childWindows[childWindows.length]=objName;
   }
}

function imgBtnOnClick(widgetName) {
   var imgBtnTbl = widgetName + "_tbl";
   var imgBtn = document.all[imgBtnTbl];
   if(imgBtn.className == "imagebutton") {
      if (imgBtn.style.border.indexOf("outset") > -1) {
         imgBtn.style.border="2px inset white";
      } else {
         imgBtn.style.border="2px outset white";
      }
   }
}

function buttonOnFocus(pControl) {
   try {
      localonFocus(pControl);
   } catch(e) {}
}

function buttonOnBlur(pControl) {
   try {
      localonBlur(pControl);
   } catch(e) {}
}

/** setGridInnerHTML: fixes (hopefully) compatibility issues between the grid applet and
JVM versions by creating document elements.
*/
function refreshGridWidget(widget_name) {
  //alert('refreshGridWidget: ' + widget_name);
  refreshGridURL(widget_name,"em_do_grid_prc?p_widget_name="+widget_name+"&p_type=HTML");
}

function refreshGridURL(tagID,vURL) {
   //alert('refreshGridURL: ' + vURL);

   //get the results from the procedure (vURL)
   var objHTTP=new ActiveXObject("microsoft.XMLHTTP");
   objHTTP.open('POST', getServerPath()+vURL, false);
   objHTTP.send();

   //update the grid
   refreshGridHTML(tagID,objHTTP.responseText);
}


function refreshGridHTML(tagID,vHTML) {
   var split_params;
   var curNode;

   //alert('refreshGridHTML: ' + vHTML.length + ' bytes');
   //alert('refreshGridHTML: ' + vHTML);

   if (document.all[tagID].tagName == 'APPLET') {
     curNode = document.all[tagID].parentElement;
   }
   else {
     curNode = document.all[tagID];
   }

   mouseOn();

   //reset the innerHTML of the target div
   curNode.innerHTML="";

   //split up the resulting data into an accessible array.
   //(split on the line breaks)
   split_params = vHTML.split("\n");

   //check that the first line is <APPLET and the last line contains </APPLET>
   //var linefirst = split_params[0].toUpperCase();
   //if (linefirst.indexOf("<APPLET") == -1) {
   //  alert('ERROR: APPLET IS NOT PROPERLY FORMATTED!!');
   //}

   //create instance of doc element for the very first split out line
   //(this needs to be the start <applet...> tag in all cases).
   //var appElement = document.createElement(split_params[0]);

   //iterate through the rest of the responseText, creating/appending child elements
   //to the initial appElement.  Omits the end </applet> tag from this procedure
   //since it is not needed (already assumed with the above createElement).
   var appElement;
   for (i=0; i<(split_params.length); i++) {
       vparam = split_params[i];
       if (vparam.toUpperCase().indexOf("<APPLET") != -1) {
         appElement = document.createElement(vparam);
       }
       if (vparam.toUpperCase().indexOf("<PARAM") != -1) {
         paramElement = document.createElement(vparam);
         appElement.appendChild(paramElement);
       }
   }

   //insert our new element(s) into the document hierarchy.
   curNode.insertBefore(appElement);
   mouseOff();
}

/**********************************************************************/
/* this function works with the em_replace_chars function in eMARS db */
/**********************************************************************/
function renderHTML(valStr) {
   hCode=new Array(34,38,39,44,47,60,62,92,145,146,147,148);
   var hCodeStr="";

   for (x=0; x<hCode.length; x++) {
      hCodeStr=eval("/&#" + hCode[x] + ";/g");
      valStr=valStr.replace(hCodeStr,String.fromCharCode(hCode[x]));

      if (x<100) {
         hCodeStr=eval("/&#0" + hCode[x] + ";/g");
         valStr=valStr.replace(hCodeStr,String.fromCharCode(hCode[x]));
      }
   }
   return valStr;
}

//cprice added 12/20/2004 for stripping invalid characters from email address input.
function stripInvalid(value) {
    var result = "";
    var myregexp = new RegExp("\@.*");
    result = value;
    for(var x = 1; x < result.length; ++x) {
      //replace commas and spaces with nothing
      result = result.replace(",", "");
      result = result.replace(" ", "");
    }
    //strip @ and anything following it
    result = result.replace(myregexp, "");
    
    return result;
}

