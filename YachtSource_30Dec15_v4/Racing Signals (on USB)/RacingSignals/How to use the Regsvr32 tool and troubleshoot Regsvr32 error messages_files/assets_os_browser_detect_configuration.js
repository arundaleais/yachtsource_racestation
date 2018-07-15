    
function BrowserVersion() {
    
    BrowserDetect.init();
    
    /* ie versions */
    
    if ((BrowserDetect.browser == "Internet Explorer") && (BrowserDetect.version == 10)) {
        return 'ie10';
    } else if ((BrowserDetect.browser == "Internet Explorer") && (BrowserDetect.version == 9)) {
        return 'ie9';
    } else if ((BrowserDetect.browser == "Internet Explorer") && (BrowserDetect.version <= 8)) {
        return 'ie8';
    } else if (BrowserDetect.browser != "Internet Explorer") {
        return 'non_ie';
    };  
    
}

function OSVersion() {

    BrowserDetect.init();
    
    /* win8 */
    
    if ((BrowserDetect.OS == "Windows") && (BrowserDetect.OSVersion == "Windows 8")) {
        return 'win8';
    };

    /* win7 */

    if ((BrowserDetect.OS == "Windows") && (BrowserDetect.OSVersion == "Windows 7")) {
        return 'win7';
    };
    
    /* vista */

    if ((BrowserDetect.OS == "Windows") && (BrowserDetect.OSVersion == "Windows Vista")) {
        return 'vista';
    };
    
    /* winXP */

    if ((BrowserDetect.OS == "Windows") && (BrowserDetect.OSVersion == "Windows XP Professional")) {
        return 'winxp';
    };

    /* other windows*/

    if ((BrowserDetect.OS == "Windows") && ((BrowserDetect.OSVersion != "Windows XP Professional") && (BrowserDetect.OSVersion != "Windows Vista") && (BrowserDetect.OSVersion != "Windows 7") && (BrowserDetect.OSVersion != "Windows 8"))) {
        return 'windows';
    };

    /* non windows - linux, etc.*/

    if ((BrowserDetect.OS != "Windows") && (BrowserDetect.OS != "Mac")) {
        return 'non_windows';
    };
    
    /* mac */
    
    if (BrowserDetect.OS == "Mac") {
        return 'mac_os';
    };
}


function current_text_init(currentId, currentHeaderId , additionalProductsId, locationTextId, Configuration) {
    
    var browser = BrowserVersion();
    var os = OSVersion();
    var infobanner = os + "_" + browser;
    
    var currentContent	= null;
    
    /*var currenttext = '';*/
    var currenttext = Configuration.currenttext_windows;
    var location_text_win_os = '';
    var product_text = '';
    var location_text = '';
    
    
    switch (infobanner) {
        case 'win8_ie10': currentContent = Configuration.win8_ie10; break;
        case 'win8_non_ie': currentContent = Configuration.win8_non_ie; break;
        case 'win7_ie10': currentContent = Configuration.win7_ie10; break;
        case 'win7_ie9': currentContent = Configuration.win7_ie9; break;   
        case 'win7_ie8': currentContent = Configuration.win7_ie8; break;
        case 'win7_non_ie': currentContent = Configuration.win7_non_ie; break;
        case 'vista_ie10': currentContent = Configuration.vista_ie10; break;
        case 'vista_ie9': currentContent = Configuration.vista_ie9; break;
        case 'vista_ie8': currentContent = Configuration.vista_ie8; break;
        case 'vista_non_ie': currentContent = Configuration.vista_non_ie; break;
        case 'winxp_ie8': currentContent = Configuration.winxp_ie8; break;
        case 'winxp_non_ie': currentContent = Configuration.winxp_non_ie; break;
        case 'windows_ie8': currentContent = Configuration.win_ie8; break;
        case 'windows': currentContent = Configuration.win; break;
        case 'non_windows': currentContent = Configuration.non_win; break;
        case 'mac_os': currentContent = Configuration.mac; break;
    }
	
    /*if (browser == 'non_ie') {
        currenttext = Configuration.currenttext_windows;
	location_text_win_os = Configuration.location_text_win_non_ie;
    } else {
	currenttext = Configuration.currenttext_windows_ie;
	location_text_win_os = Configuration.location_text_win_ie;
    }*/
    
    if ((os == 'non_windows') || (os == 'mac_os')){
        product_text = Configuration.non_windows_products;
	location_text = Configuration.location_text_non_win;
    } else {
	product_text = Configuration.windows_products;
	/*location_text = location_text_win_os;*/
        location_text = Configuration.location_text_win;
    }
    
    if ( currentContent != null )
    {
	var displaycurrent = $(Configuration.parentloc).clone().children(currentContent);
	
	$("#"+currentHeaderId).removeClass("current_disable").addClass("active");
	$('#'+currentHeaderId+ ' a').append(currenttext);
	$('#'+currentId).removeClass("current_disable").attr("style", "display:block;");
	$('#'+currentId).append(displaycurrent);
	$('#'+additionalProductsId+ ' a').append(product_text);
	$('#'+locationTextId+ ' p').append(location_text);
	$('.hide_products').addClass("current_disable");
	$(displaycurrent).attr("style", "display:block; width:100%;");
	$(displaycurrent).find("div").addClass("sec");
        $(displaycurrent).after('<div class="clear"></div>');
    }      
        
}

function smc_accordeon(div_id) {
    
        var hash = smc_getUrlVars();
    
    
        if ($.inArray("SegNo", hash) != -1) {
            $("#"+div_id).Accordion({ AllowAllClosed: true, width: '100%', ExpandedSegments: [hash.SegNo] });
        }
        else {
            $("#"+div_id).Accordion({ AllowAllClosed: true, width: '100%'});
        }
    

    function smc_getUrlVars() {
        var ParamArray = [], hash;

        var Parameters = window.location.search.substring(1).split('&');
        for (var iParam = 0; iParam < Parameters.length; iParam++) {
            hash = Parameters[iParam].split('=');
            ParamArray.push(hash[0]);
            ParamArray[hash[0]] = hash[1];
        }

        return ParamArray;
    }
}
