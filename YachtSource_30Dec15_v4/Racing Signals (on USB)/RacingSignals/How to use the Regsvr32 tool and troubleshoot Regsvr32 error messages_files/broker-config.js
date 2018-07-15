/*
Copyright (c) 2014, comScore Inc. All rights reserved.
version: 5.0.3
*/
COMSCORE.SiteRecruit.Broker.config = {
	version: "5.0.3",
	//cddsDomains: 'store.microsoft.com|xbox.com|windowsphone.com|www.microsoft.com|go.microsoft.com|microsoftstore.com',
	cddsDomains: 'store.microsoft.com|xbox.com|www.microsoft.com|go.microsoft.com|microsoftstore.com',
	cddsInProgress: 'cddsinprogress',
	domainSwitch: 'tracking3p',
	domainMatch: '([\\da-z\.-]+\.com)',
	delay: 3000,
	
	//TODO:Karl extend cookie enhancements to ie userdata
		testMode: false,
	
	// cookie settings
	cookie:{
		name: 'msresearch',
		path: '/',
		domain:  '.microsoft.com' ,
		duration: 90,
		rapidDuration: 0,
		expireDate: ''
	},
	thirdPartyOptOutCookieEnabled : false,
	
	// optional prefix for pagemapping's pageconfig file
	prefixUrl: "",
	
	//events
	Events: {
		beforeRecruit: function() {
var _halt = true;			
var _days = false;
var _stC = readCookie("ST_GN_EN-US");
 if(_stC){
  var _s = _stC.split('.');
  _t = _s[0].split('_');
	_d = _t[1];
  var mult = 86400000;
  
  if(_d && _d > 0){
   var myDate = new Date().getTime();
   var b = myDate - (_d * mult);
   _days = (b/mult);
 }
}

if(_days && _days <= 90){
	_halt = true;
}

if(/((microsoftstore|xbox|windowsphone|microsoft|microsoftbusinesshub|virtualacademy|windows|office|skype|live|volumelicensing|onenote|visualstudio).com|windowsmediaplayer)/i.test(document.referrer) ) {
	_halt = true;
}
var screenWidth = self.screen.availWidth, screenHeight = self.screen.availHeight; //var screenWidth = self.screen.width, screenHeight = self.screen.Height;		
if(/iphone|ipad|ipod|android|opera mini|blackberry|bb10|mobile safari|windows (phone|ce)|iemobile|htc|nokia|mobile/i.test(navigator.userAgent) || screenWidth <= 450 ){
	_halt = true;
}

if(typeof ms != 'undefined' && typeof ms.ssversion != 'undefined'){
	if(ms.ssversion == '200'){
		//COMSCORE.SiteRecruit.Broker.config.mapping[0].halt = _halt;
		_halt = false;
	}
}
if(typeof document.getElementsByName('ms.ssversion')[0] != 'undefined' && typeof document.getElementsByName('ms.ssversion')[0].getAttribute('content') != 'undefined'){
	_ssvervion = document.getElementsByName('ms.ssversion')[0].getAttribute('content');	
	if(_ssvervion == '200'){
		//COMSCORE.SiteRecruit.Broker.config.mapping[0].halt = true;
		_halt = false;
	}
}
COMSCORE.SiteRecruit.Broker.config.mapping[0].halt = _halt;

					}
	},
	
	mapping:[
	// m=regex match, c=page config file (prefixed with configUrl), f=frequency
	{m: '', c: 'inv_c_p329970507.js', f: 0.10, p: 0 	}
	//,{m: 'test=comscore|forceorigin=osg', c: 'inv_c_p329970507.js', f: 0, p: 1 ,prereqs:{	content: [] ,cookie: [ {'name':'tstMode','value':'1'}] ,externalDomain: [] } }
	//,{m: 'sg1\.support\.services\.microsoft\.com\/en-us', c: 'inv_c_p329970507.js', f: 0, p: 0	,prereqs:{	content: [] ,cookie: [ {'name':'tstMode','value':'1'}] ,externalDomain: [] } }
]
};

function readCookie(name){var ca = document.cookie.split(';');  var nameEQ = name + "=";  for(var i=0; i < ca.length; i++) {    var c = ca[i];    while (c.charAt(0)==' ') c = c.substring(1, c.length); if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length, c.length);    }  return false;}
var gIdelay = 0;
if (readCookie("graceIncr") == 1) {
	gIdelay = 5000;
}
setTimeout(function(){_set_SessionCookie("graceIncr", 0)},gIdelay);

// START 5.1.3
function _set_SessionCookie(_name, _val) {	  
	if (_name == COMSCORE.SiteRecruit.Broker.config.domainSwitch) {
		var r = new RegExp(COMSCORE.SiteRecruit.Broker.config.domainMatch,'i');
		if (r.test(_val)) {
			_val = RegExp.$1 + RegExp.$2;
			var c = _name + '=' + _val + '; path=/' + '; domain=' + COMSCORE.SiteRecruit.Broker.config.cookie.domain;
			document.cookie = c; 
		}
	}else if(COMSCORE.isDDInProgress()){	
 		if(_name == "captlinks"){
 			if(/^http(s)?\:/i.test(_val)){
				var _reg = new RegExp("http(s)?://"+document.domain+"/", "i");
 				var _val = _val.replace(_reg, '');
 			}
 			if(_val && _val.length > 2){
				c_vals = readCookie("captlinks");
				if(c_vals){
   				if(c_vals.indexOf(_val) == -1){
   					var str = c_vals +"," + _val;
   					if(str.length <= 1240){
   						_val = str;
   					}else{ _val=false; }
   				}else{ _val = false; }
  			}
 			}
 		}
  	if(_val){
  		var c = _name+'=' + _val + '; path=/' + '; domain=' + COMSCORE.SiteRecruit.Broker.config.cookie.domain;
    	document.cookie = c;
    }
	}
}
// END 5.1.3
setTimeout('_set_SessionCookie("graceIncr","0")', 3000);
//START 5.1.3 CDDS-captLink-graceIncr handlers
function SRappendEventListener(srElement, _name, _val){
	if(srElement.addEventListener){
			srElement.addEventListener('click',function(event){	_set_SessionCookie(_name, _val); },false);
	}else{
			srElement.attachEvent('onclick',function(){	_set_SessionCookie(_name, _val); });
	}
}
var allLinks = document.getElementsByTagName("a");
for (var i = 0, n = allLinks.length; i < n; i++){
	var r = new RegExp(COMSCORE.SiteRecruit.Broker.config.cddsDomains,'i');
	var _clickURL = allLinks[i].href;
	
	if (r.test(_clickURL)) {
		if (/store.microsoft.com/i.test(_clickURL)) { _clickURL = ".microsoftstore.com" }
		//if (/www.microsoft.com\/windowsphone/i.test(_clickURL)) { _clickURL = ".windowsphone.com" }
		//if (/go.microsoft.com/i.test(_clickURL) && /279032/.test(_clickURL)) { _clickURL = ".windowsphone.com" }
		if (!(/microsoft.com/i.test(_clickURL))) {
			SRappendEventListener(allLinks[i], COMSCORE.SiteRecruit.Broker.config.domainSwitch, _clickURL);
		}
	}
	try{
	if(_clickURL && _clickURL != '' && !(/javascript\:void(0)/i.test(_clickURL)) ){
		//if(/login\.live|msacademicverify|(o15\.officeredir|office)\.microsoft\.com|login|LiveLogin|(office|office\.microsoft|live|skype|windowsphone|xbox|onedrive)\.com/i.test(_clickURL)){
		//REMOVED WINDOWSPHONE, XBOX
		if(/login\.live|msacademicverify|(o15\.officeredir|office)\.microsoft\.com|login|LiveLogin|(office|office\.microsoft|live|skype|onedrive)\.com/i.test(_clickURL)){
			//SRappendEventListener(allLinks[i], "graceIncr", _clickURL);
			SRappendEventListener(allLinks[i], "graceIncr", 1);
		}

		if( /(contactus\/(technicalsupport|setupandinstallation)) || (my\/(account|devicesoftware|supportrequests) ) /i.test(_clickURL)){
			if(/sign in/i.test(document.getElementById("idPPScarab").innerHTML)) {
				//SRappendEventListener(allLinks[i], "graceIncr", _clickURL);
				SRappendEventListener(allLinks[i], "graceIncr", 1);
			}
		} 
	}
 }catch(e){}
}

var cs_inputs = document.getElementsByTagName('input');
for(var c=0; c<cs_inputs.length; c++){
  if(cs_inputs[c].getAttribute('ms.cmpnm')=='signin'){
  	SRappendEventListener(cs_inputs[c], "graceIncr", 1);
	}
}

var cs_divs = document.getElementsByTagName('div');
for(var c=0; c<cs_divs.length; c++){
	if(/msame_Header_name msame_TxtTrunc|msame_Header msame_unauth/i.test(cs_divs[c].getAttribute('class'))){
    SRappendEventListener(cs_divs[c], "graceIncr", 1);
  }
}
//END 5.1.3 CDDS-captLink-graceIncr handlers


//CUSTOM - CHECK FOR THE CROSS-DOMAIN COOKIE. IF PRESENT, HALT RECRUITMENT AND SET DD TRACKING COOKIE
	function crossDomainCheck() {
		if (intervalMax > 0) {
			intervalMax --;
			
			var cookieName = COMSCORE.SiteRecruit.Broker.config.cddsInProgress;
			
			if (COMSCORE.SiteRecruit.Utils.UserPersistence.getCookieValue(cookieName) != false ) {
				COMSCORE.SiteRecruit.DDKeepAlive.setDDTrackerCookie();
				COMSCORE.SiteRecruit._halt = true;
				clearCrossDomainCheck();
			}
		}
		else {
			clearCrossDomainCheck();
		}
	}

	function clearCrossDomainCheck() {
		window.clearInterval(crossDomainInterval);
	}

	var intervalMax = 10;
	
	var crossDomainInterval = window.setInterval('crossDomainCheck()', '1000');
//END CROSS_DOMAIN DEPARTURE FUNCTIONALITY



//CUSTOM - ADD 5 SECOND DELAY ON CALLING BROKER.RUN()
if (COMSCORE.SiteRecruit.Broker.delayConfig == true)  {
	COMSCORE.SiteRecruit.Broker.config.delay = 1000;
}
/*
else{
	COMSCORE.SiteRecruit.Broker.config.delay = 3000;
}
*/

//CUSTOM - SUPPORT POC DELAY - 10 SECONDS ON RECRUITMENT
if ( ( !(/support.microsoft.com/i.test(document.referrer)) || document.referrer == '' ) && COMSCORE.SiteRecruit.Broker.isDDInProgress() == false) {
	COMSCORE.SiteRecruit.Broker.config.delay = 10000;
}

window.setTimeout('COMSCORE.SiteRecruit.Broker.run()', COMSCORE.SiteRecruit.Broker.config.delay);
//END CUSTOM