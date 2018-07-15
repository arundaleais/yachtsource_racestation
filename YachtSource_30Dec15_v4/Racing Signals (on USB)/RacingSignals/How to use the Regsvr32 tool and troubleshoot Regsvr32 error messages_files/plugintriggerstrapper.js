if (!window._ms_support_fms_surveyScriptHandlerCallback)
{
    window._ms_support_fms_surveyScriptHandlerCallback = [];
}

if (!window._ms_support_fms_utility_getCookie)
{
    window._ms_support_fms_utility_getCookie = function(key)
    {
        var entities = document.cookie.split(";");
        for (var i = 0; i < entities.length; ++i)
        {
            var j = entities[i].indexOf("=");
            var s = entities[i].substring(0, j);
            if (s != "" && (s == key || s == " " + key))
            {
                return entities[i].substring(j + 1);
            }
        }
        return null;
    }
}

if (!window._ms_support_fms_utility_setCookie)
{
    window._ms_support_fms_utility_setCookie = function(key, value, domain, expires, path)
    {
        domain = domain || document.domain;
        path = path || "/";
        document.cookie = key + "=" + value + "; domain=" + domain + "; path=" + path + "; " + (expires ? ("expires=" + expires.toGMTString() + ";") : "");
    }
}

if (!window._ms_support_fms_utility_removeCookie)
{
    window._ms_support_fms_utility_removeCookie = function(key, domain, path)
    {
        window._ms_support_fms_utility_setCookie(key, "", domain, (new Date(0, 0, 0)),  path);
    }
}

if (!window._ms_support_fms_utility_pauseTracking)
{
    window._ms_support_fms_utility_pauseTracking = function(domain, timeout)
    {
        var expires = null;
        if (timeout)
        {
            expires = new Date();
            expires.setMilliseconds(expires.getMilliseconds() + timeout);
        }
        window._ms_support_fms_utility_setCookie("fmsphb", "1", domain, expires);
    }
}

if (!window._ms_support_fms_utility_resumeTracking)
{
    window._ms_support_fms_utility_resumeTracking = function(domain)
    {
        window._ms_support_fms_utility_removeCookie("fmsphb", domain);
    }
}

if (!window._ms_support_fms_utility_packageQueryString)
{
    window._ms_support_fms_utility_packageQueryString = function(obj)
    {
        var queryString = "";
        for (var key in obj)
        {
            if (obj.hasOwnProperty(key))
            {
                queryString += "&" + key + "=" + encodeURIComponent(obj[key]);
            }
        }
        return queryString;
    }
}

(function()
{
    var triggerConfig = typeof (_ms_support_fms_surveyTriggerConfig) != "undefined" ? _ms_support_fms_surveyTriggerConfig : false;
    if (triggerConfig)
    {
        var protocol = window.location.protocol;
        if (protocol != "http:" && protocol != "https:")
        {
            protocol = "http:";
        }

        var config =
		{
		    "protocol": protocol,
		    "host": triggerConfig.host || "support.microsoft.com",
		    "noHassle": isNaN(triggerConfig.noHassle) ? 7776000 : triggerConfig.noHassle,
		    "site":
			{
			    "id": parseInt(triggerConfig.site.id),
			    "name": triggerConfig.site.name,
			    "culture": triggerConfig.site.culture || "en-us",
			    "lcid": parseInt(triggerConfig.site.lcid) || 1033,
			    "brand": parseInt(triggerConfig.site.brand) || 0,
			    "version": unescape(triggerConfig.site.version),
			    "explicitDomain": (window.location.hostname != document.domain) || (triggerConfig.site.explicitDomain ? true : false),
			    "cookieDomain": triggerConfig.site.cookieDomain || document.domain,
			    "previousCookieDomain": triggerConfig.site.previousCookieDomain || null,
			    "loginUrlPattern": unescape(triggerConfig.site.loginUrlPattern) || null
			},
		    "content":
			{
			    "id": unescape(triggerConfig.content.id),
			    "type": triggerConfig.content.type,
			    "culture": triggerConfig.content.culture || "en-us",
			    "lcid": parseInt(triggerConfig.content.lcid) || 1033,
			    "properties": unescape(triggerConfig.content.properties),
			    "group": unescape(triggerConfig.content.group),
			    "keywords": unescape(triggerConfig.content.keywords),
			    "lastreviewed": unescape(triggerConfig.content.lastreviewed),
			    "technologies": unescape(triggerConfig.content.technologies)
			},
		    "invitation":
			{
			    "width": unescape(triggerConfig.invitation.width) || 0,
			    "header": unescape(triggerConfig.invitation.header),
			    "footer": unescape(triggerConfig.invitation.footer)
			},
		    "tracking":
			{
			    "header": unescape(triggerConfig.tracking.header),
			    "footer": unescape(triggerConfig.tracking.footer),
			    "timeout": parseInt(triggerConfig.tracking.timeout) || 3000,
			    "loginTimeout": parseInt(triggerConfig.tracking.loginTimeout) || 60000,
			    "width": parseInt(triggerConfig.tracking.width) || 460,
			    "height": parseInt(triggerConfig.tracking.height) || 320,
			    "bottomMargin": parseInt(triggerConfig.tracking.bottomMargin) || 15,
			    "rightMargin": parseInt(triggerConfig.tracking.rightMargin) || 10,
			    "hostedDomains": unescape(triggerConfig.tracking.hostedDomains),
			    "parentPollingTimeout": parseInt(triggerConfig.tracking.parentPollingTimeout) || 60000,
			    "parentPollingDelay": parseInt(triggerConfig.tracking.parentPollingDelay) || 5000,
			    "navigationTimeout": parseInt(triggerConfig.tracking.navigationTimeout) || 5000
			},
		    "parameters": triggerConfig.parameters ? triggerConfig.parameters : []
		};

        if (triggerConfig.blessedDomains)
        {
            window.blessedDomains = triggerConfig.blessedDomains;
            window.hostedDomains = config.tracking.hostedDomains;
            window.cookieDomain = config.site.cookieDomain;
            window.parentPollingTimeout = config.tracking.parentPollingTimeout;
            window.parentPollingDelay = config.tracking.parentPollingDelay;
        }

        if (window.blessedDomains && !window.crossDomainInitialized) {
            var version = "0";            
            try {                
                var temp = ($('script').filter(function () { if ($(this).attr('src')) return $(this).attr('src').toLowerCase().indexOf('plugintriggerstrapper') > -1; })).attr('src');                
                if (temp.indexOf("?") > -1)
                    version = temp.substring(temp.indexOf("?") + 1);
            }
            catch (e) { }
            var protocol = window.location.protocol;
            if (protocol != "http:" && protocol != "https:") {
                protocol = "http:";
            }          

            var header = "crossdomain.js";
            var src = protocol + "//" + config.host + "/common/script/fx/" + header + "?" + version;
            var script = document.createElement("script");
            script.type = "text/javascript";
            script.src = src;

            (document.body || document.documentElement).appendChild(script);
        }

        window.isDomainTracking = function()
        {
            var entry = _ms_support_fms_utility_getCookie('fmshb');
            if (entry)
            {
                try
                {
                    return entry.split(',')[0] == '1' ? true : false;
                }
                catch (e)
                {
                    return false;
                }
            }
            return false;
        };

        var addEventHandler = window.attachEvent ? function(el, ev, fp) { el.attachEvent("on" + ev, fp); } : function(el, ev, fp) { el.addEventListener(ev, fp, false); };
        var domLoadEvent = typeof (document.onreadystatechange) != "undefined" ? [document, "readystatechange"] : (typeof (document.onDOMContentLoaded) != "undefined" ? [document, "DOMContentLoaded"] : [window, "load"]);

        function handleMixedAuthenticationLogin()
        {
            if (config.site.loginUrlPattern != null)
            {
                var loginUrlTester = new RegExp(config.site.loginUrlPattern, "ig");
                var timeout = config.tracking.loginTimeout;
                addEventHandler(domLoadEvent[0], domLoadEvent[1],
					function(e)
					{
					    addEventHandler(document, "click",
							function(e)
							{
							    var a = e.srcElement ? e.srcElement : e.target;
							    while (a && a.tagName && a.tagName.toLowerCase() != "a")
							    {
							        a = a.parentNode;
							    }

							    if (a && a.href && loginUrlTester.test(a.href.toString()))
							    {
							        window._ms_support_fms_utility_pauseTracking(config.site.cookieDomain, timeout);
							    }
							}
						);
					}
				);
            }
        }

        var optOut = _ms_support_fms_utility_getCookie("fmsOptOut" + config.site.name.toUpperCase());
        if (optOut && optOut == "1")
        {
            // user may have already accept a survey with opt-out checked, continue send heart beat
            if (isDomainTracking())
            {
                window.setInterval(function() { _ms_support_fms_utility_setCookie("fmshb", (isDomainTracking() ? '1' : '0') + ',' + new Date().getTime(), config.site.cookieDomain); }, 1000);
                handleMixedAuthenticationLogin();
            }

            return;
        }
        else
        {
            window.setInterval(function() { _ms_support_fms_utility_setCookie("fmshb", (isDomainTracking() ? '1' : '0') + ',' + new Date().getTime(), config.site.cookieDomain); }, 1000);
            handleMixedAuthenticationLogin();
        }

        if (isDomainTracking())
        {
            return;
        }

        window.fmsSurveyExpired = function(days)
        {
            var MiliDay = 86400000;
            var visits = _ms_support_fms_utility_getCookie(("ST_" + config.site.name + "_" + config.site.culture).toUpperCase());
            
            if (null == visits)
            {
                return true;
            }

            if (triggerConfig.blessedDomains && config.site.previousCookieDomain != null)
            {
                var fmsvs = _ms_support_fms_utility_getCookie("fmsvs");
                if (document.domain == config.site.previousCookieDomain && fmsvs == null)
                {
                    var expiresDate = new Date();
                    expiresDate.setFullYear(expiresDate.getFullYear() + 10);
                    _ms_support_fms_utility_removeCookie(("ST_" + config.site.name + "_" + config.site.culture).toUpperCase(), config.site.previousCookieDomain, "/");
                    _ms_support_fms_utility_setCookie(("ST_" + config.site.name + "_" + config.site.culture).toUpperCase(), visits, config.site.cookieDomain, expiresDate, "/");
                    _ms_support_fms_utility_setCookie("fmsvs", "15S10", config.site.cookieDomain, expiresDate, "/");
                }
            }

            var parts = visits.split('_');
            if (parts.length != 3 || isNaN(parts[0]))
            {
                return true;
            }

            var origDate = parseInt(parts[1]);
            var curDate = new Date();
            return ((curDate.getTime() / MiliDay - days) >= origDate);
        }

        if (!fmsSurveyExpired(config.noHassle / (24 * 60 * 60)))
        {
            return;
        }

        var callback = {
            config: config
        };

        config.callbackId = window._ms_support_fms_surveyScriptHandlerCallback.push(callback) - 1;
        var triggerHost = encodeURI(protocol + "//" + config.host);

        var params = {
            "CallbackId": config.callbackId,
            "Site": config.site.name,
            "Region": config.site.culture,
            "ContentId": config.content.id,
            "ContentType": config.content.type,
            "Group": config.content.group,
            "Keywords": config.content.keywords,
            "LastReviewed": config.content.lastreviewed,
            "Technologies": config.content.technologies,
            "Title": document.title,
            "Url": window.location.href,
            "ReferringUrl": document.referrer,
            "Followup": _ms_support_fms_utility_getCookie("fmsfollowups" + ("ST_" + config.site.name + "_" + config.site.culture).toUpperCase()) || "",
            "History": _ms_support_fms_utility_getCookie("fmsmemo") || ""
        };

        function packageQueryStringWithOverflow(obj, baseLength, maxLength)
        {
            var overflow = "";
            var queryString = "";

            for (var key in obj)
            {
                if (obj.hasOwnProperty(key))
                {
                    var entity = "&" + key + "=" + encodeURIComponent(obj[key]);

                    if (baseLength + queryString.length + entity.length > maxLength)
                    {
                        overflow += entity;
                    }
                    else
                    {
                        queryString += entity;
                    }
                }
            }

            return { query: queryString, overflow: overflow };
        }

        var MAX_URL_LENGTH = 2048;
        var MAX_INT_LENGTH = 23;

        var timestamp = (new Date()).getTime().toString(32);
        var randomNumber = Math.random().toString(32).substring(2);
        var chunkHeadLength = ("&ts=" + timestamp + "&rnd=" + randomNumber + "&chkc=").length + MAX_INT_LENGTH;

        var handler_src_base = triggerHost + "/common/surveyscripthandler.ashx?template=trigger";
        var queryString = packageQueryStringWithOverflow(params, handler_src_base.length, MAX_URL_LENGTH - chunkHeadLength);
        if (handler_src_base.length + queryString.query.length + queryString.overflow.length < MAX_URL_LENGTH)
        {
            queryString.query += queryString.overflow;
            queryString.overflow = "";
        }
        var handler_src = handler_src_base + queryString.query;

        function addScriptHandler(extension)
        {
            var handler = document.createElement("script");
            handler.src = handler_src + (extension || "");
            handler.type = "text/javascript";
            (document.body || document.documentElement).appendChild(handler);
        }

        function packageChunks(value, triggerHost)
        {
            var chunks = [];
            var dsBaseUrl = triggerHost + "/common/SurveyDataStorageHandler.ashx?ts=" + timestamp + "&rnd=" + randomNumber + "&chkid=";
            var maxChunkIdLength = 23;
            var encodingMargin = 3;
            var valuePrefix = "&v=";
            var dsBaseLength = dsBaseUrl.length + maxChunkIdLength + valuePrefix.length;
            var maxChunkLength = MAX_URL_LENGTH - dsBaseLength - encodingMargin;

            var encodedValue = encodeURIComponent(value);
            var length = encodedValue.length;
            var latest = length - 1;

            for (var i = 0; i < length; i += maxChunkLength)
            {
                var end = Math.min(i + maxChunkLength, latest);
                chunks.push(encodedValue.substring(i, end));
            }

            for (var i = 0; i < chunks.length; ++i)
            {
                if (chunks[i][chunks[i].length - 1] == "%")
                {
                    chunks[i + 1] = "%" + (chunks[i + 1] || "");
                    chunks[i] = chunks[i].substring(0, chunks[i].length - 2);
                }
                else if (chunks[i][chunks[i].length - 2] == "%")
                {
                    chunks[i + 1] = "%" + chunks[i][chunks[i].length - 1] + (chunks[i + 1] || "");
                    chunks[i] = chunks[i].substring(0, chunks[i].length - 3);
                }
            }

            return { baseUrl: dsBaseUrl, timestamp: timestamp, randomNumer: randomNumber, valuePrefix: valuePrefix, chunks: chunks };
        }

        var hasChunks = queryString.overflow ? true : false;

        if (hasChunks)
        {
            var chunks = packageChunks(queryString.overflow, triggerHost);

            function sendChunks(onComplete)
            {
                var count = 0;
                if (!window._ms_support_fms_dataStorage_response)
                {
                    window._ms_support_fms_dataStorage_response = [];
                }
                window._ms_support_fms_dataStorage_response[timestamp + "" + randomNumber] =
				{
				    callback: function(chunkId)
				    {
				        ++count;
				        if (count == 1)
				        {
				            // once there is a response for the first chunk, send other chunks in bulk.
				            for (var i = 1; i < chunks.chunks.length; ++i)
				            {
				                var handler = document.createElement("script");
				                handler.src = chunks.baseUrl + i + chunks.valuePrefix + chunks.chunks[i];
				                handler.type = "text/javascript";
				                (document.body || document.documentElement).appendChild(handler);
				            }
				        }
				        if (count == chunks.chunks.length)
				        {
				            // all chunks have been sent to server, send the major request
				            onComplete("&ts=" + timestamp + "&rnd=" + randomNumber + "&chkc=" + chunks.chunks.length);
				        }
				    }
				};

                // send the first chunk
                var handler = document.createElement("script");
                handler.src = chunks.baseUrl + 0 + chunks.valuePrefix + chunks.chunks[0];
                handler.type = "text/javascript";
                (document.body || document.documentElement).appendChild(handler);
            }
        }

        addEventHandler(domLoadEvent[0], domLoadEvent[1], function(e)
        {
            if (typeof (document.readyState) == "undefined" || document.readyState == "complete")
            {
                if (hasChunks)
                {
                    sendChunks(addScriptHandler);
                }
                else
                {
                    addScriptHandler();
                }
            }
        });
    }
})();