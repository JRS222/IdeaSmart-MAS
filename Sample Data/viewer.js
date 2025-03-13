// JavaScript Document
// determine the viewer installed on end user's computer
// if return w - whip! 4 installed
// if return d - design reviewer installed
function getviewerflag(){
	var wflag = false;
	var vflag = false;
	var dflag = false;

		try {
		    vflag = new ActiveXObject("AdView.AdViewer.1");
		    vflag = true;
		} catch(ex) {
		    vflag = false;
		}
		try {
		    dflag = new ActiveXObject("DesignReview.ViewsCtrl.1");
		    dflag = true;
		} catch(ex) {
		    dflag = false;
		}
		try {
		    wflag = new ActiveXObject("WHIP.WhipCtrl.1");
		    wflag = true;
		} catch(ex) {
		    wflag = false;
		}
		if (dflag == true)
			return "d";
		else if (wflag == true)
			return "w";
		else if (vflag == true)
			return "v";
		else
			return "d";
					
}
