// JavaScript Document
var currentWhepFacID;
var currentMachineType;
var currentMachineNo;
var currentMachinePtr;
var currentMachineClassCode;    
var siteName;
var currentECBMSubsite;

function bodyOnLoad() {

	currentWhepFacID=opener.currentWhepFacID;
	currentMachineType=opener.currentMachineType;
	currentMachineNo=opener.currentMachineNo;
	currentMachinePtr=opener.currentMachinePtr;
	currentMachineClassCode = opener.currentMachineClassCode;
	siteName = opener.siteName;
	currentECBMSubsite=opener.currentECBMSubsite;
	
	document.getElementById('btn_print2').style.display='none';

	document.getElementById('div_list').innerHTML = "Loading ...";
	document.getElementById('div_list').innerHTML = getChecklists();
}

function getChecklists()
{
	var objAcroList = null;
    var objAcros=new ActiveXObject('microsoft.XMLDOM');
    objAcros.async=false;
        
     /* get the existing APPS configuration */
   	objAcros.load("ecbm.ecbm_report_j.getAllChecklists?pSiteID="+currentSite+"&pSubsiteID="+currentECBMSubsite+"&pAcronym="+currentMachineType);
 
   if (objAcros.parsed) {
      	  objAcroList=objAcros.getElementsByTagName("ROW");
		  
		  var htmlStr =  '<table border="0" rules="rows" align="center" width="100%">';
          if(objAcroList.length ==0)
          {
                htmlStr += '<tr><th colspan="6">There is not any checklist available.</th></tr>';
          }else{
          
		  htmlStr += '<tr><th>Equip`t ID</tH><th>Equip`t No.</th><th>Equip`t Type</th><th>Class Code</th><th>Site Name</th></tr>';
                        
                for(var i=0;i<objAcroList.length;i++)
                {               
                        var bgco = '#C0C0C0';
                        if(i % 2 == 0) { bgco = '#D3D3D3'; }
                        htmlStr += '<tr bgcolor="' +bgco+'">';
                        var childList=objAcroList(i).childNodes;
                        for(var j=0;j<childList.length;j++)
                        {
                                if(j==0){
                                        htmlStr += '<td><input type="button" value="'+childList[j].text +'" name="btnMachineno"  id="btnMachineno" onClick="showChecklists('+childList[0].text +',\''+childList[1].text +'\',\''+childList[2].text +'\',\''+childList[3].text +'\',\''+childList[4].text +'\')"  /></td>'; 
                                        //Jesse 5/10/2006, some equipment number may contain white-space, e.g. APPS 001. it would cause trouble if it is not put inside ' ' (single quote)
                                }else{
                                                htmlStr += '<td>'+childList[j].text + '</td>'; 
                                }
                        }
                        htmlStr += '</tr>';
                }
                htmlStr+= '<tr bgcolor="lightgrey"><td colspan="6"><font color="#000000" size="-2">click on an equipment ID to see the checklists</font></td></tr>';
          } //if -else
          htmlStr += '</table>';
          
          return htmlStr;  
   }  
   return 'The list (XML) of checklists cannot be parsed.';
}


//display all checklist of a specific machine
var tempMachineNo=0;
var tempMachineType="";
var tempFacility="";
var tempClasscode="";
function showChecklists(machineid,machineno,machinetype,classcode,facility)
{
        tempMachineNo=machineno;
		currentMachineNo=machineno;
        tempMachineType=machinetype;
		tempClasscode=classcode;
        tempFacility=facility;
        
        document.getElementById('btn_back').style.display='block';
        document.getElementById('btn_back2').style.display='none';
        document.getElementById('div_list').innerHTML = "Loading ... checklists for machine " + machineid;
        
        var objAcroList = null;
        var objAcros1=new ActiveXObject('microsoft.XMLDOM');
        objAcros1.async=false;
        
        /* get the checklist*/
   		objAcros1.load("ecbm.ecbm_report_j.getChecklists2?p_machineid=" +machineid+"&pSiteID="+currentSite);
 
   if (objAcros1.parsed) {
          objAcroList=objAcros1.getElementsByTagName("ROW");
	  	  var htmlStr = '<span style="text-align:center;"><table width="100%"><tr>' + 
			'<td bgcolor="000099"><font color="cccccc">&nbsp;Facility:&nbsp;</font></td>' + 
			'<td bgcolor="cccccc"><font color="000099">&nbsp;&nbsp;' + tempFacility + '&nbsp;&nbsp;</font></td>' + 
			'<td bgcolor="000099"><font color="cccccc">&nbsp;Acronym:&nbsp;</font></td>' + 
			'<td bgcolor="cccccc"><font color="000099">&nbsp;&nbsp;' + tempMachineType + '&nbsp;&nbsp;</font></td>' + 
			'<td bgcolor="000099"><font color="cccccc">&nbsp;Equip`t No.:&nbsp;</font></td>' + 
			'<td bgcolor="cccccc"><font color="000099">&nbsp;&nbsp;' + tempMachineNo + '&nbsp;&nbsp;</font></td>' + 
			'<td bgcolor="000099"><font color="cccccc">&nbsp;Class Code:&nbsp;</font></td>' + 
			'<td bgcolor="cccccc"><font color="000099">&nbsp;&nbsp;' + tempClasscode + '&nbsp;&nbsp;</font></td>' + 
			'</tr></table></span>';
       
	      htmlStr +=  '<table border="0" rules="rows" align="center" width="100%">';

          if(objAcroList.length ==0)
          {
                htmlStr += '<tr><th colspan="3">There is not any checklist available.</th></tr>';
          }else{
          htmlStr += '<tr><th>Checklist No.</tH><th>Creating Date</th><th>Creator</th></tr>';
                        
                for(var i=0;i<objAcroList.length;i++)
                {               
                        var bgco = '#C0C0C0';
                        if(i % 2 == 0) { bgco = '#D3D3D3'; }
                        htmlStr += '<tr bgcolor="' +bgco+'">';
                        var childList=objAcroList(i).childNodes;
                        for(var j=0;j<3;j++)
                        {
                                if(j==0){
                                        htmlStr += '<td><input type="button" value="'+childList[j].text +'" name="btnTrackno"  id="btnTrackno" onClick="showAChecklist(' + childList[0].text + ',' + machineid +',' + childList[3].text + ',' + childList[4].text +',\'' +machinetype+'\',\''+machineno+'\',\''+classcode+'\',\''+facility+'\');"  /></td>'; 
                                }else{
                                        htmlStr += '<td>'+childList[j].text + '</td>'; 
                                }
                        }
                        htmlStr += '</tr>';
                }
                
                htmlStr+= '<tr bgcolor="lightgrey"><td colspan="3"><font color="#000000" size="-2">click on a checklist no. to see the details</font></td>';
                htmlStr+= '</tr>';
          } //if -else
          htmlStr += '</table>';
          
          document.getElementById('div_list').innerHTML = htmlStr; 
   }else{  
          document.getElementById('div_list').innerHTML = 'The list (XML) of checklists cannot be parsed';
   }
}



//display all steps of a specific checklist
var globe_machineid = 0;
var globe_trackno = 0;
var globe_siteid = 0;
var globe_fac = "";
var globe_mtype = "";
var globe_mno = 0;
var globe_classcode = "";
function showAChecklist(trackno,machineid,siteid,facid,mtype,mno,classcode,fac)
{	
    //only for the 'back' button, see btn_back2_onClick()
    globe_machineid = machineid;
	//for the further print purppose
	currentMachineNo=mno;
	currentMachinePtr=machineid;
    //for print iframe content
	globe_trackno = trackno;
	globe_siteid = siteid;
	globe_fac = fac;
	globe_mtype = mtype
	globe_mno = mno;
	globe_classcode = classcode;
	
   document.getElementById('btn_back2').style.display='block';
   document.getElementById('btn_back').style.display='none';
   //document.getElementById('btn_print2').style.display='block'; //this cannot be used to print the following iframe 
   document.getElementById('btn_print').style.display='none';
   document.getElementById('div_list').innerHTML = "Loading ... checklist: " + trackno;
   
	document.getElementById('div_list').innerHTML='<iframe width="930" height="600" name="iframeMain"></iframe>';
   	document.all.iframeMain.src = getServerPath()+"ecbm.ecbm_report_j.getChecklistDetailsForIframe?p_trackno=" +trackno+'&p_site='+siteid+'&p_fac='+fac+'&p_mtype='+mtype+'&p_mno='+mno+'&p_classcode='+classcode;
  
}


var currentTaskID = 0;
function showATask(taskID)
{
		//currentMachineType=opener.currentMachineType;
		
		var objHTTP=new ActiveXObject('microsoft.XMLHTTP');
        var urlStr=getServerPath()+'ecbm.ecbm_report_j.getMachineClassCode?p_site_id='+currentSite+'&p_equip_ptr='+currentMachinePtr;
   	    objHTTP.open('GET', urlStr, false);
        objHTTP.send();
		currentMachineClassCode = objHTTP.responseText;
		
        //open a new window to display the task details
        currentTaskID = taskID;
        openSwilWin('ECBM_TASK_DETAILS');
}


//back button for showing the checklists of ONE machine
function btn_back2_onClick()
{  
        // not work   window.history.back();
        showChecklists(globe_machineid,tempMachineNo,tempMachineType,tempClasscode,tempFacility); 
        document.getElementById('btn_back2').style.display='none';
        document.getElementById('btn_back').style.display='block';
		document.getElementById('btn_print2').style.display='none';
		document.getElementById('btn_print').style.display='block';
}

//back button for showing the all checklists
function btn_back_onClick()
{
        document.getElementById('div_list').innerHTML = getChecklists();
        document.getElementById('btn_back').style.display='none';
}

//print button for printing the checklist details dispalyed in an iFrame
function  btn_print2_onClick()
{		
		var newWin = null;
		newWin = window.open("","win2","alwaysRaised,status,menubar,height=600,width=800,resizable,scrollbars,top=0,left=0",true);
	    newWin.location.href = getServerPath()+"ecbm.ecbm_report_j.getChecklistDetailsForIframe?p_trackno=" +globe_trackno+'&p_site='+globe_siteid+'&p_fac='+globe_fac+'&p_mtype='+globe_mtype+'&p_mno='+globe_mno+'&p_classcode='+globe_classcode;;
		
		newWin.document.close();
	    newWin.focus();
		newWin.print();	
}

