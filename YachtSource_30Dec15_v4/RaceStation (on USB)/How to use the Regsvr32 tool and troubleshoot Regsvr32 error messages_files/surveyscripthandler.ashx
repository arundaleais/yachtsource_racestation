
(function(){
var version = '81';
if (!window._ms_support_fms_loadedScriptLibraries)
{
	window._ms_support_fms_loadedScriptLibraries = [];
}
function addScriptLibrary(header, async)
{
	if (_ms_support_fms_loadedScriptLibraries[header])
	{
		return;
	}
	_ms_support_fms_loadedScriptLibraries[header] = true;
	var src = config.protocol + "//" + config.survey.host + "/common/script/" + header + "?" + version;
	if (async)
	{
		var script = document.createElement("script");
		script.type = "text/javascript";
		script.src = src;

		(document.body || document.documentElement).appendChild(script);
	}
	else
	{
		document.write("<script type=\"text/javascript\" src=\"" + src + "\"></scr" + "ipt>");
	}
}

if (!window._ms_support_fms_loadedStyleSheets)
{
	window._ms_support_fms_loadedStyleSheets = [];
}
function addStyleSheet(stylesheet, async)
{
	if (_ms_support_fms_loadedStyleSheets[stylesheet])
	{
		return;
	}
	_ms_support_fms_loadedStyleSheets[stylesheet] = true;
	var href = config.protocol + "//" + config.survey.host +"/common/css/" + stylesheet + "?" + version;
	if (document.createStyleSheet)
	{
		document.createStyleSheet(href);
	}
	else
	{
		if (async)
		{
			var link = document.createElement("link");
			link.rel = "stylesheet";
			link.type = "text/css";
			link.href = href;

			(document.getElementsByTagName("head")[0] || document.body || document.documentElement).appendChild(link);
		}
		else
		{
			document.write("<link rel=\"stylesheet\" type=\"text/css\" href=\"" + encodeURI(href) + "\" />");
		}
	}
}

function getSurveyContent(config)
{
	return '\x3cnoscript xmlns\x3asw\x3d\x22urn\x3aschemas-microsoft-com\x2fPSS\x2fPSS_Survey01\x22\x3e\x3cdiv class\x3d\x22SURVEYNOSCRIPT\x22\x3e\x3c\x2fdiv\x3e\x3c\x2fnoscript\x3e\x3cstyle type\x3d\x22text\x2fcss\x22 xmlns\x3asw\x3d\x22urn\x3aschemas-microsoft-com\x2fPSS\x2fPSS_Survey01\x22\x3e\x0a        body\x7b\x0a        display\x3a block\x3b\x0a        \x7d\x0a      \x3c\x2fstyle\x3e\x3cdiv id\x3d\x22SURVERCONTAINER_plugin0\x22 class\x3d\x22SURVEYCONTAINER\x22 xmlns\x3asw\x3d\x22urn\x3aschemas-microsoft-com\x2fPSS\x2fPSS_Survey01\x22\x3e\x3cinput type\x3d\x22hidden\x22 id\x3d\x22surveyname\x22 name\x3d\x22surveyname\x22 value\x3d\x22\x22\x3e\x3cinput type\x3d\x22hidden\x22 id\x3d\x22showpage\x22 name\x3d\x22showpage\x22 value\x3d\x220\x22\x3e\x3cdiv class\x3d\x22SURVEYTITLETEXT\x22\x3e\x3c\x2fdiv\x3e\x3cdiv id\x3d\x22SURVEYSECTIONCONTAINER\x22\x3e\x3cdiv id\x3d\x22SURVEYSECTION_132\x22 class\x3d\x22SURVEYSECTION\x22 style\x3d\x22display\x3anone\x3b\x22\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22id\x22 value\x3d\x22132\x22\x3e\x3cdiv name\x3d\x22CompoundingBranchRules\x22 style\x3d\x22display\x3anone\x3b\x22\x3e\x3c\x2fdiv\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22NextSectionByTopicRef\x22 value\x3d\x220\x22\x3e\x3cdiv id\x3d\x22SURVEYQUESTION_12616\x22 class\x3d\x22QUESTIONCONTAINER\x22\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22id\x22 value\x3d\x2212616\x22\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22type\x22 value\x3d\x22CHOICE\x22\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22required\x22 value\x3d\x220\x22\x3e\x3cdiv class\x3d\x22QUESTIONTEXT\x22 role\x3d\x22heading\x22\x3eWas this information helpful\x3f\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONREQUIREDPREFIXED\x22\x3e\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONREQUIREDCLEAR\x22\x3e\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONINSTRUCTION\x22\x3e\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONREQUIRED\x22\x3e\x3c\x2fdiv\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22randomization\x22 value\x3d\x22\x22\x3e\x3ctable cellspacing\x3d\x220\x22 cellpadding\x3d\x220\x22 border\x3d\x220\x22\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22\x3e\x3ctable cellspacing\x3d\x220\x22 cellpadding\x3d\x220\x22 border\x3d\x220\x22\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22 class\x3d\x22CHOICEROW\x22\x3e\x3cinput type\x3d\x22radio\x22 name\x3d\x22QUESTION_132_1261612616\x22 id\x3d\x22125\x22 value\x3d\x220\x22 title\x3d\x22Yes\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22ANSWERTEXT\x22\x3eYes\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22\x3e\x3ctable cellspacing\x3d\x220\x22 cellpadding\x3d\x220\x22 border\x3d\x220\x22\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22 class\x3d\x22CHOICEROW\x22\x3e\x3cinput type\x3d\x22radio\x22 name\x3d\x22QUESTION_132_1261612616\x22 id\x3d\x22126\x22 value\x3d\x221\x22 title\x3d\x22No\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22ANSWERTEXT\x22\x3eNo\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22\x3e\x3ctable cellspacing\x3d\x220\x22 cellpadding\x3d\x220\x22 border\x3d\x220\x22\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22 class\x3d\x22CHOICEROW\x22\x3e\x3cinput type\x3d\x22radio\x22 name\x3d\x22QUESTION_132_1261612616\x22 id\x3d\x2218874\x22 value\x3d\x222\x22 title\x3d\x22Somewhat\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22ANSWERTEXT\x22\x3eSomewhat\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2fdiv\x3e\x3cdiv id\x3d\x22SURVEYQUESTION_10576\x22 class\x3d\x22QUESTIONCONTAINER\x22\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22id\x22 value\x3d\x2210576\x22\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22type\x22 value\x3d\x22CHOICE\x22\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22required\x22 value\x3d\x220\x22\x3e\x3cdiv class\x3d\x22QUESTIONTEXT\x22 role\x3d\x22heading\x22\x3eHow much effort did you personally put forth to use this article\x3f\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONREQUIREDPREFIXED\x22\x3e\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONREQUIREDCLEAR\x22\x3e\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONINSTRUCTION\x22\x3e\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONREQUIRED\x22\x3e\x3c\x2fdiv\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22randomization\x22 value\x3d\x22\x22\x3e\x3ctable cellspacing\x3d\x220\x22 cellpadding\x3d\x220\x22 border\x3d\x220\x22\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22\x3e\x3ctable cellspacing\x3d\x220\x22 cellpadding\x3d\x220\x22 border\x3d\x220\x22\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22 class\x3d\x22CHOICEROW\x22\x3e\x3cinput type\x3d\x22radio\x22 name\x3d\x22QUESTION_132_1057610576\x22 id\x3d\x2216305\x22 value\x3d\x221\x22 title\x3d\x22Very low\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22ANSWERTEXT\x22\x3eVery low\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22\x3e\x3ctable cellspacing\x3d\x220\x22 cellpadding\x3d\x220\x22 border\x3d\x220\x22\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22 class\x3d\x22CHOICEROW\x22\x3e\x3cinput type\x3d\x22radio\x22 name\x3d\x22QUESTION_132_1057610576\x22 id\x3d\x2216306\x22 value\x3d\x222\x22 title\x3d\x22Low\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22ANSWERTEXT\x22\x3eLow\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22\x3e\x3ctable cellspacing\x3d\x220\x22 cellpadding\x3d\x220\x22 border\x3d\x220\x22\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22 class\x3d\x22CHOICEROW\x22\x3e\x3cinput type\x3d\x22radio\x22 name\x3d\x22QUESTION_132_1057610576\x22 id\x3d\x2263705\x22 value\x3d\x223\x22 title\x3d\x22Moderate\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22ANSWERTEXT\x22\x3eModerate\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22\x3e\x3ctable cellspacing\x3d\x220\x22 cellpadding\x3d\x220\x22 border\x3d\x220\x22\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22 class\x3d\x22CHOICEROW\x22\x3e\x3cinput type\x3d\x22radio\x22 name\x3d\x22QUESTION_132_1057610576\x22 id\x3d\x2216308\x22 value\x3d\x224\x22 title\x3d\x22High\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22ANSWERTEXT\x22\x3eHigh\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22\x3e\x3ctable cellspacing\x3d\x220\x22 cellpadding\x3d\x220\x22 border\x3d\x220\x22\x3e\x3ctr\x3e\x3ctd valign\x3d\x22top\x22 class\x3d\x22CHOICEROW\x22\x3e\x3cinput type\x3d\x22radio\x22 name\x3d\x22QUESTION_132_1057610576\x22 id\x3d\x2216309\x22 value\x3d\x225\x22 title\x3d\x22Very high\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22ANSWERTEXT\x22\x3eVery high\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2fdiv\x3e\x3cdiv id\x3d\x22SURVEYQUESTION_12617\x22 class\x3d\x22QUESTIONCONTAINER\x22\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22id\x22 value\x3d\x2212617\x22\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22type\x22 value\x3d\x22TEXT-BLOCK\x22\x3e\x3cinput type\x3d\x22hidden\x22 name\x3d\x22required\x22 value\x3d\x220\x22\x3e\x3cdiv class\x3d\x22QUESTIONTEXT\x22 role\x3d\x22heading\x22\x3eTell us what we can do to improve this article\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONREQUIREDPREFIXED\x22\x3e\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONREQUIREDCLEAR\x22\x3e\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONINSTRUCTION\x22\x3e\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22QUESTIONREQUIRED\x22\x3e\x3c\x2fdiv\x3e\x3ctextarea class\x3d\x22ANSWERBOX\x22 name\x3d\x22TEXTBLOCK_12617\x22 rows\x3d\x2212\x22 cols\x3d\x2280\x22 onpaste\x3d\x22enforceMaxLength\x28this,1024,event\x29\x22 onkeyup\x3d\x22enforceMaxLength\x28this,1024,event\x29\x22\x3e\x3c\x2ftextarea\x3e\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22NAVIGATION\x22\x3e\x3ctable class\x3d\x22NAVBUTTONCONTAINER\x22\x3e\x3ctr\x3e\x3ctd class\x3d\x22PADDINGCELL\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22NAVBUTTONCELL\x22 id\x3d\x22PreviousButton\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22NAVBUTTONCELL\x22\x3e\x3cinput class\x3d\x22NAVBUTTON\x22 type\x3d\x22button\x22 name\x3d\x22next\x22 data-nav\x3d\x22true\x22 data-name\x3d\x22btnNext\x22 value\x3d\x22Submit\x22 id\x3d\x22SURVEYSECTION_132_next\x22 onclick\x3d\x22plugin0.next\x28\x29\x3bthis.blur\x28\x29\x3b\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22ENDPADDINGCELL\x22\x3e\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2fdiv\x3e\x3c\x2fdiv\x3e\x3cdiv ID\x3d\x22DIV_CLOSE\x22 style\x3d\x22display\x3anone\x22\x3e\x3cdiv class\x3d\x22SURVEYTHANKYOUCONTAINER\x22\x3e\x3cdiv class\x3d\x22SURVEYINTROTEXT\x22\x3eThank you\x21 Your feedback is used to help us improve our support content. For more assistance options, please visit the \x3ca href\x3d\x22http\x3a\x2f\x2fsupport.microsoft.com\x22\x3eHelp and Support Home Page\x3c\x2fa\x3e.\x0d\x0a\x3cscript type\x3d\x22text\x2fjavascript\x22\x3e\x0d\x0a\x24\x28function\x28\x29\x7b\x0d\x0a setTimeout\x28function\x28\x29\x7b\x24\x28\x22div.kbSurvey div.fms \x23SURVEYSECTIONCONTAINER .SURVEYSECTION .NAVIGATION table.NAVBUTTONCONTAINER\x22\x29.show\x28\x29\x3b\x7d, 2000\x29\x3b\x0d\x0a\x7d\x29\x3b\x0d\x0a\x3c\x2fscript\x3e\x3c\x2fdiv\x3e\x3c\x2fdiv\x3e\x3cdiv class\x3d\x22NAVIGATION\x22\x3e\x3ctable class\x3d\x22NAVBUTTONCONTAINER\x22 nowrap\x3d\x22yes\x22\x3e\x3ctr\x3e\x3ctd class\x3d\x22PADDINGCELL\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22NAVBUTTONCELL\x22\x3e\x3c\x2ftd\x3e\x3ctd class\x3d\x22ENDPADDINGCELL\x22\x3e\x3c\x2ftd\x3e\x3c\x2ftr\x3e\x3c\x2ftable\x3e\x3c\x2fdiv\x3e\x3c\x2fdiv\x3e\x3c\x2fdiv\x3e\x3c\x2fdiv\x3e';
}
var callbackid = parseInt('0', 10) || 0;
var callback = _ms_support_fms_surveyScriptHandlerCallback[callbackid];
var config = callback.config;
callback.callback(getSurveyContent);

function startSurvey ()
{
	var survey = new MS.Support.Fms.Survey("SURVERCONTAINER_plugin" + callbackid, config);
	window["plugin" + callbackid] = survey;
	if (config.survey.isInvitation)
	{
		survey.isInvitation = true;
	}
	new MS.Support.Fms.Plugin.Survey(survey, config);
	if (config.survey.isInvitation)
	{
		new MS.Support.Fms.SurveyInvite(survey, config);
	}

	survey.start();
}

function isFmsPluginRuntimeLoaded()
{
	return (window.MS && MS.Support && MS.Support.Fms && MS.Support.Fms.Survey && MS.Support.Fms.Plugin && MS.Support.Fms.Plugin.Survey && ((!config.survey.isInvitation) || MS.Support.Fms.SurveyInvite));
}

var RenderOptionEnum = {
	"default": 0,
	"nodefault": 1,
	"overridedefault": 2,
	"overrideltr": 4,
	"overridertl": 8,
	"overridedirections": 4 | 8,
	"overrideall": 2 | 4 | 8
};

function getRenderOption()
{
	var renderOption = RenderOptionEnum["default"];
	if (config.survey.renderOption)
	{
		var options = config.survey.renderOption.split(",");
		for (var i = 0; i < options.length; ++i)
		{
			var option = options[i];
			var value = parseInt(option);
			if (isNaN(value))
			{
				option = option.toLowerCase();
				if (option in RenderOptionEnum)
				{
					value = RenderOptionEnum[option];
				}
				else
				{
					return RenderOptionEnum["default"];
				}
			}
			renderOption |= value;
		}
	}
	return renderOption;
}

var async = callback.target ? true : false;
var renderOption = getRenderOption();
if (!(renderOption & RenderOptionEnum["nodefault"]))
{
	addStyleSheet("fx/surveypage.css", async);
	addStyleSheet("fx/survey_" + (config.survey.isRTL? "rtl" : "ltr") + ".css", async);
}

if (renderOption & RenderOptionEnum["overrideall"])
{
	var altStyle = config.survey.altStyle;
	if (altStyle != null && /^\w+$/i.test(altStyle))
	{
		if (renderOption & RenderOptionEnum["overridedefault"])
		{
			addStyleSheet("fx/" + altStyle + "/default.css", async);
		}

		if ((renderOption & RenderOptionEnum["overrideltr"]) && (!config.survey.isRTL))
		{
			addStyleSheet("fx/" + altStyle + "/ltr.css", async);
		}

		if ((renderOption & RenderOptionEnum["overridertl"]) && (config.survey.isRTL))
		{
			addStyleSheet("fx/" + altStyle + "/rtl.css", async);
		}
	}
}

addScriptLibrary("fx/survey.js", async);
addScriptLibrary("fx/pluginsurvey.js", async);

if (config.survey.isInvitation)
{
	addStyleSheet("fx/surveyinvitation.css", async);
	addStyleSheet("fx/surveyinvitation_" + (config.survey.isRTL? "rtl" : "ltr") + ".css", async);
	addScriptLibrary("fx/surveyinvite.js", async);
}

var intervalHandler = 0;
intervalHandler = window.setInterval(function (){if (isFmsPluginRuntimeLoaded()) {startSurvey();window.clearInterval(intervalHandler);}}, 20);

if (config.enableLTS && !window._ms_support_fms_ltsHandlerAttached)
{
	function _ms_support_fms_LtsUnloadHandler()
	{
		var dwellTime = (new Date()).getTime() - window._ms_support_fms_ltsStartTime;
		var ltsImg = document.getElementById("StatsDotNetImg");
		ltsImg.src = window._ms_support_fms_ltsURL + "&unload=true&dwelltime=" + encodeURIComponent(dwellTime);

		//delay a while to ensure LTS get logged
		var timeout = (parseInt('300') || 125) + (new Date()).getTime();
		while(timeout >= (new Date()).getTime())
		{
		}
	}

	function getCookieIncrement(key)
	{
		var value = parseInt(window._ms_support_fms_utility_getCookie(key)) || 0;
		++value;
		window._ms_support_fms_utility_setCookie(key, value, config.site.cookieDomain);
		return value;
	}

	function _ms_support_fms_LtsLoadHandler()
	{
		var config = window._ms_support_fms_ltsConfig;

		var params = {
			SsId: config.site.id,
			SiteLcid: config.site.lcid,
			SiteCulture: config.site.culture,
			EventCollectionID: 1,
			SsVersion: config.site.version,
			ContentType: config.content.type,
			ContentLcid: config.content.lcid,
			ContentCulture: config.content.culture,
			ContentId: window._ms_support_fms_ltsContentId,
			ContentProperties: config.content.properties,
			BrandId: config.site.brand,
			EventSeqNo: getCookieIncrement("sdninc"),
			Platform: 'windows nt 5.1' || 'winxp',
			URL: window.location.href,
			RefURL: document.referrer || ""
		};

		window._ms_support_fms_ltsURL = ltsHostUrl + "?" + window._ms_support_fms_utility_packageQueryString(params);
		var ltsImg = document.createElement("img");

		ltsImg.id = "StatsDotNetImg";
		ltsImg.name = "StatsDotNetImg";
		ltsImg.alt = "";
		ltsImg.height = 0;
		ltsImg.width = 0;

		(document.body || document.documentElement).appendChild(ltsImg);

		ltsImg.src = window._ms_support_fms_ltsURL;
	}

	var ltsHostUrl = config.protocol + "//" + ('support.microsoft.com\x2fLTS\x2fdefault.aspx' || (config.survey.host + "/LTS/"));

	window._ms_support_fms_ltsHandlerAttached = true; // if multiple surveys are hosted on this page, only the first lts handler will be attached.
	window._ms_support_fms_ltsStartTime = (new Date()).getTime();
	window._ms_support_fms_ltsConfig = config;

	var addEventHandler = window.attachEvent ? function (el, ev, fp){el.attachEvent("on" + ev, fp);} : function(el, ev, fp){el.addEventListener(ev, fp, false);};

	if (callback.target)
	{
		_ms_support_fms_LtsLoadHandler();
	}
	else
	{
		addEventHandler(window, "load", _ms_support_fms_LtsLoadHandler);
	}

	addEventHandler(window, "unload", _ms_support_fms_LtsUnloadHandler);
}
})();

