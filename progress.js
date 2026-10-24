var id;
var sctr = 5;
var visValue = "visible";

function noSplash() {
	var so=window.clearTimeout(id);
}

function hideSplash(){
	sctr--;
	if (sctr==0) {
		visValue = "hidden";
		noSplash();
	}		
	if(document.layers){
		 //thisbrowser="NN4";
		splashDiv = document.layers["splash"];
		splashDiv.document.open();
		splashDiv.document.style.visibility=visValue;
		splashDiv.document.close();
	}
	if(document.all){
		//thisbrowser="ie"
		splashDiv = document.all["splash"];
		splashDiv.style.visibility=visValue;
	}
	if(document.getElementById){
		//thisbrowser="NN6";
		splashDiv = document.getElementById("splash");
		splashDiv.style.visibility=visValue;
	}
	id = setTimeout("hideSplash();",1000);	
}

function splashDiv(divName){
	sctr--;
	if (sctr==0) {
		visValue = "hidden";
		noSplash();
	}		
	if(document.layers){
		 //thisbrowser="NN4";
		splashDiv = document.layers[divName];
		splashDiv.document.open();
		splashDiv.document.style.visibility=visValue;
		splashDiv.document.close();
	}
	if(document.all){
		//thisbrowser="ie"
		splashDiv = document.all[divName];
		splashDiv.style.visibility=visValue;
	}
	if(document.getElementById){
		//thisbrowser="NN6";
		splashDiv = document.getElementById(divName);
		splashDiv.style.visibility=visValue;
	}
	id = setTimeout("splashDiv();",1000);	
}

