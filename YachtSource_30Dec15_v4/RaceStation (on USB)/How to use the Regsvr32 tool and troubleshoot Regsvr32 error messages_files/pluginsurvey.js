if (!window.MS)
{
    window.MS = {};
}

if (!MS.Support)
{
    MS.Support = {};
}

if (!MS.Support.Fms)
{
    MS.Support.Fms = {};
}

if (!MS.Support.Fms.Plugin)
{
    MS.Support.Fms.Plugin = {};
}

if (!MS.Support.Fms.Plugin.Survey)
{
    MS.Support.Fms.Plugin.Survey = function(survey, config)
    {
        //using
        var Fms = MS.Support.Fms;
        var Survey = Fms.Survey;
        var Utils = Fms.Utils;
        var addEventHandler = Utils.addEventHandler;
        var removeEventHandler = Utils.removeHandler;

        function delayHalfSecond(delay)
        {
            try
            {
                if (!delay) delay = 500;
                var today = new Date();
                var now = today.getTime();
                while (1)
                {
                    var today2 = new Date();
                    var now2 = today2.getTime();
                    if ((now2 - now) > delay) { break; };
                }
            }
            catch (e) { }
        }

        function createHiddenFrame(name)
        {
            var container = document.createElement("DIV");
            container.innerHTML = "<IFRAME id=\"" + name + "\" name=\"" + name + "\" src=\"about:blank\" border=0 frameBorder=0 style=\"width:0px;height:0px;border:none;padding:0px;margin:0px;\"></IFRAME>";
            (document.body || document.documentElement).appendChild(container);

            return document.getElementById(name);
        }

        function createSubmitForm(target)
        {
            var submitForm = document.createElement("form");
            submitForm.style.height = 0;
            submitForm.style.width = 0;
            submitForm.style.padding = 0;
            submitForm.style.margin = 0;
            submitForm.style.border = "none";
            submitForm.method = "POST";
            submitForm.encoding = "application/x-www-form-urlencoded";
            submitForm.enctype = "application/x-www-form-urlencoded";
            submitForm.action = "https://" + config.survey.host + "/common/survey.aspx?FMSAction=FinishPlugin";
            submitForm.target = target;

            (document.body || document.documentElement).appendChild(submitForm);

            return submitForm;
        }

        var QuitMode =
		{
		    "giveup": 1,
		    "cancel": 2,
		    "persist": 0,
		    "suspend": 3
		};

        function getQuitModeByAction(action)
        {
            var quitMode = QuitMode[action.toLowerCase()];

            if (quitMode == null)
            {
                throw "Unknown action: " + action;
            }

            return quitMode;
        }

        function postback(survey, action)
        {
            var target = "submitframe_" + survey.id;
            var submitFrame = createHiddenFrame(target);
            var submitForm = createSubmitForm(target);

            var quitMode = getQuitModeByAction(action);
            var surveyAnswers = survey.encodeAnswers(function(input) { return unicodeFixup(escape(input)); });
            survey.addSubmitField("DATALENGTH", surveyAnswers.split("|").length);
            survey.addSubmitField("SURVEYANSWERS", surveyAnswers + "|" + quitMode);

            for (var field in survey.submitFields)
            {
                var fieldElement = document.createElement("input");
                fieldElement.type = "hidden";
                fieldElement.name = field;
                fieldElement.value = survey.submitFields[field];

                submitForm.appendChild(fieldElement);
            }

            if (survey.isInvitation)
            {
                submitForm.target = "_self";
                try
                {
                    submitForm.submit();
                    delayHalfSecond();
                }
                catch (e)
                {
                }
                return;
            }
            else
            {
                var waitCallback = true;
                if (quitMode == QuitMode.persist)
                {
                    try
                    {
                        var frameWindow = window.frames[target];
                        frameWindow.document.open();
                        frameWindow.document.close();
                        addEventHandler(frameWindow,
							"unload",
							function()
							{
							    if (survey.thankyou != null)
							    {
							        survey.thankyou.show();
							    }

							    try
							    {
							        if (window.navigator.userAgent.indexOf("MSIE") > -1)
							        {
							            window.setTimeout(function() { submitFrame.parentNode.removeChild(submitFrame); }, 1000);
							        }
							        else if (window.navigator.userAgent.indexOf("Firefox") > -1)
							        {
							            submitFrame.parentNode.removeChild(submitFrame);
							        }
							        frameWindow.document.open();
							        frameWindow.document.close();
							    }
							    catch (e)
							    {
							    }
							}
						);
                    }
                    catch (e)
                    {
                        waitCallback = false;
                    }
                    window._ms_support_fms_utility_setSubmitted(config.hash, config.site.cookieDomain);
                }

                try
                {
                    submitForm.submit();

                    if (quitMode == QuitMode.persist && (!waitCallback) && survey.thankyou != null)
                    {
                        survey.thankyou.show();
                    }

                    if (quitMode == QuitMode.giveup)
                    {
                        delayHalfSecond(1500);
                    }
                }
                catch (e)
                {
                }
            }
        }

        function handleInvitationSubmit(survey, action)
        {
            var surveyDiv;

            survey.addSubmitField("TRIGGERID", config.triggerConfig.entity.TriggerId);

            var isPersist = ((action == "persist") ? true : false);

            if (!isPersist)
            {
                postback(survey, "cancel");
            }

            try
            {
                surveyDiv = window.parent.document.getElementById('surveyDivBlock');
                surveyDiv.style.display = 'none';
                surveyDiv.style.height = 0;
                surveyDiv.firstChild.style.height = 0;
            }
            catch (e) { }

            if (isPersist)
            {
                if (config.triggerConfig.entity.Event.toLowerCase() == "onunload")
                {
                    var header = "<div id=\"header\">" + config.triggerConfig.tracking.header + "</div>";
                    var content = "<div id=\"content\">" + survey.getTrackingText() + "</div>";
                    var footer = "<div id=\"footer\">" + config.triggerConfig.tracking.footer + "</div>";
                    var pages = [];
                    for (var i = 0; i < config.triggerConfig.entity.Pages.length; ++i)
                    {
                        pages[i] = escape(config.triggerConfig.entity.Pages[i]);
                    }
                    var submitFields = "";
                    for (var field in survey.submitFields)
                    {
                        submitFields += "<input type=\"hidden\" name=\"" + escape(field) + "\"" + " value=\"" + escape(survey.submitFields[field]) + "\" />"
                    }

                    // verify explicit domain again, because document.domain may has been changed after the trigger snippet
                    var explicitDomain = (window.location.hostname != document.domain) ? true : config.site.explicitDomain;

                    var script = explicitDomain ? ("<script type=\"text/javascript\">document.domain=\"" + document.domain + "\";</scri" + "pt>") : "";

                    var blessedDomainsScript = "";
                    if (window.blessedDomains && typeof (JSON) != "undefined" && JSON.stringify) {
                        blessedDomainsScript = "<script type=\"text/javascript\">" + "var blessedDomains =" + JSON.stringify(window.blessedDomains) + "</script>";
                    }

                    var version = "0";
                    try {
                        var temp = $('script[src*="/pluginsurvey"]').attr('src');
                        if (temp.indexOf("?") > -1)
                            version = temp.substring(temp.indexOf("?") + 1);
                    }
                    catch (e) { }

                    var trackerHtml = "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">" +
						"<html>" +
							"<head>" +                                
								"<style type=\"text/css\">" +
									"*{margin:0;padding:0;}" +
									"html,body,#wrap{height:100%}" +
									"body{direction:" + (config.triggerConfig.entity.IsRTL ? "rtl" : "ltr") + "}" +
									"#header,#content,#footer {width:100%;}" +
									"#footer{position:absolute;bottom:0;}" +
								"</style>" +
								script +
                                blessedDomainsScript +
							"</head>" +
							"<body style=\"margin:0\">" +
									"<input id=\"triggerContains\" type=\"hidden\" value=\"" + escape(config.triggerConfig.entity.Contains) + "\" />" +
									"<input id=\"triggerPages\" type=\"hidden\" value=\"" + escape(pages.join(",")) + "\" />" +
									"<input id=\"triggerSurveyUrl\" type=\"hidden\" value=\"" + escape(config.triggerConfig.entity.Redirect || config.triggerConfig.fullSurveyUrl.replace(/surveyinvite\.aspx/ig, "survey.aspx")) + "\" />" +
									"<input id=\"openerUrl\" type=\"hidden\" value=\"" + escape(window.top.location.href) + "\" />" +
									"<input id=\"surveyHost\" type=\"hidden\" value=\"" + escape(config.survey.host) + "\" />" +
									"<input id=\"cookieDomain\" type=\"hidden\" value=\"" + escape(config.site.cookieDomain) + "\" />" +
									"<input id=\"timeout\" type=\"hidden\" value=\"" + config.triggerConfig.tracking.timeout + "\" />" +									
                                    "<input id=\"hostedDomains\" type=\"hidden\" value=\"" + config.triggerConfig.tracking.hostedDomains + "\" />" +
                                    "<input id=\"parentPollingTimeout\" type=\"hidden\" value=\"" + config.triggerConfig.tracking.parentPollingTimeout + "\" />" + 
                                    "<input id=\"navigationTimeout\" type=\"hidden\" value=\"" + config.triggerConfig.tracking.navigationTimeout + "\" />" +
									"<div id=\"submitFields\" style=\"display:none\">" +
										submitFields +
									"</div>" +
									"<div id=\"wrap\">" +
										header +
										content +
										footer +
									"</div>" +
								"</body>" +
                                "<script type=\"text/javascript\" src =\"" + config.protocol + "//ajax.aspnetcdn.com/ajax/jQuery/jquery-1.7.2.min.js\"></scr" + "ipt>" +
                                "<script type=\"text/javascript\" src=\"" + config.protocol + "//" + config.survey.host + "/common/script/fx/plugintriggertracking.js?" + version + "\"></scri" + "pt>" +
							"</html>";

                    var url = "about:blank";

                    if (explicitDomain && window.navigator.userAgent.indexOf("MSIE") > -1 || window.opera)
                    {
                        url = "javascript:(function(){document.open();document.domain=\"" + document.domain + "\";document.close();})();";
                    }

                    if (window.blessedDomains) {
                        var protocol = window.location.protocol;
                        if (protocol != "http:" && protocol != "https:") {
                            protocol = "http:";
                        }
                        url = protocol + "//" + document.domain + "/survey.html";
                        var surveywinwidth = config.triggerConfig.tracking.width;
                        var surveywinheight = config.triggerConfig.tracking.height;
                        var surveywintop = screen.height - surveywinheight - (screen.height*(config.triggerConfig.tracking.bottomMargin/100));
                        var surveywinleft = screen.width - surveywinwidth - (screen.width*(config.triggerConfig.tracking.rightMargin/100));
                        var surveywin = window.top.open(url, 'trackingWindow', "resizable=1,left=" + surveywinleft + ",top=" + surveywintop + ",width=" + surveywinwidth + ",height=" + surveywinheight + ",scrollbars=1,status=1");
                    }
                    else
                    {
                        var surveywin = window.top.open(url, "_blank", "resizable=1,left=200,top=200,width=1024,height=700,scrollbars=1,status=1");
                    }

                    try {
                        if (window.sessionStorage) {
                            var newDate = new Date();
                            window.sessionStorage.tabSessionID = newDate.getTime();
                        }
                        document.cookie = "fmshd=" + escape(window.location.protocol + "//" + window.location.hostname + ";")  + "; domain=" + config.site.cookieDomain + "; path=/";
                    } catch (e) { }

                    if (explicitDomain && (window.navigator.userAgent.indexOf("MSIE") > -1 || window.opera))
                    {
                        surveywin.execScript("document.open();document.write(unescape(\"" + escape(trackerHtml) + "\"));document.close();");
                    }
                    else
                    {
                        surveywin.document.open();
                        surveywin.document.write(trackerHtml);
                        surveywin.document.close();
                    }

                    // for Netscape display like IE, surveywin will always open in new tab and surveywin.blur() will hide the whole window
                    if (!(window.navigator.userAgent.indexOf("Netscape") > 0 && window.navigator.userAgent.indexOf("MSIE") > 0))
                    {
                        surveywin.blur();
                    }
                }
                else
                {
                    window.open(config.triggerConfig.entity.Redirect || config.triggerConfig.fullSurveyUrl.replace(/surveyinvite\.aspx/ig, "survey.aspx"), "_blank");
                }
            }
            else if (surveyDiv != null && action == "cancel")
            {
                addEventHandler(
					window,
					"unload",
					function()
					{
					    try
					    {
					        if (window.navigator.userAgent.indexOf("MSIE") > -1)
					        {
					            window.parent.setTimeout("var surveyDiv = document.getElementById(\"surveyDivBlock\");surveyDiv.removeChild(surveyDiv.firstChild);", 1000);
					        }
					        else if (window.navigator.userAgent.indexOf("Firefox") > -1)
					        {
					            surveyDiv.removeChild(surveyDiv.firstChild);
					        }
					        document.open();
					        document.close();
					    }
					    catch (e) { }
					}
				);
            }
            else
            {
                delayHalfSecond(1500);
            }
        }

        function submit(survey, action)
        {
            if (survey.submitted)
            {
                return;
            }
            survey.submitted = true;

            if (survey.isInvitation)
            {
                handleInvitationSubmit(survey, action);
            }
            else
            {
                postback(survey, action);
            }
        }

        function validateErrorHandler(survey, validateResult)
        {
            alert(validateResult.errorMessage);
        }

        survey.submitHandler = submit;
        survey.onValidateError.add(new Fms.SurveyEventDelegate(null, validateErrorHandler));

        function unicodeFixup(s)
        {
            var result = new String();
            var c = '';
            var i = -1;
            var l = s.length;
            result = '';
            for (i = 0; i < l; i++)
            {
                c = s.substring(i, i + 1);
                if (c == '%')
                {
                    result += c; i++;
                    c = s.substring(i, i + 1);
                    if (c != 'u')
                    {
                        if (parseInt('0x' + s.substring(i, i + 2)) > 128) { result += 'u00'; }
                    }
                }
                    /* Product Studio Bug 37129
                    This fix is needed to preserve '+' in the input when client-side escaped strings are decoded in server-side code.
                    Jscript escape() does not escape a '+' to '%2B'.
                    System.Web.HttpUtility.UrlDecode() replaces '+' with a space, but decodes '%2B' just fine.
                    Jscript unescape() also decodes '%2B' just fine. */
                else if (c == '+')
                {
                    c = '%2B';
                }
                result += c;
            }
            return result;
        }

        if (survey.isInvitation)
        {
            addEventHandler(window, "beforeunload", function() { survey.giveup(); });

            function prepareHttpsTunnel()
            {
                var img = document.createElement("img");
                img.src = "https://" + config.survey.host + "/library/images/support/cn/onepix.gif";
                img.width = 0;
                img.height = 0;
                (document.body || document.documentElement).appendChild(img);
            }

            prepareHttpsTunnel();
        }

        var fmsurl = window.location.toString();
        survey.addSubmitField("FMSURL", (survey.isInvitation ? window.top : window).location.toString());


        survey.addSubmitField("SURVEYSCID", config.survey.scid);
        survey.addSubmitField("SURVEYLANGCODE", config.survey.language);
        survey.addSubmitField("SURVEYID", config.survey.id);
        survey.addSubmitField("SITE", config.site.name);
        survey.addSubmitField("REGIONID", config.site.culture);
        survey.addSubmitField("SITECULTURE", config.site.culture);
        survey.addSubmitField("SSID", config.site.id);
        survey.addSubmitField("SITEBRANDID", config.site.brand);
        survey.addSubmitField("SSVERSION", config.site.version);
        survey.addSubmitField("PARAMS", config.parameters);
        survey.addSubmitField("PARAMLENGTH", config.parameters.length);
        survey.addSubmitField("SUBMISSIONGUID", config.submissionGuid);
        survey.addSubmitField("CONTENTTYPE", config.content.type);
        survey.addSubmitField("CONTENTCULTURE", config.content.culture);
        survey.addSubmitField("CONTENTID", config.content.id);
        survey.addSubmitField("CONTENTLCID", config.content.lcid);
        survey.addSubmitField("CONTENTPROPERTIES", config.content.properties);
        survey.addSubmitField("HASH", config.hash + "_" + config.callbackId);
    }
}
