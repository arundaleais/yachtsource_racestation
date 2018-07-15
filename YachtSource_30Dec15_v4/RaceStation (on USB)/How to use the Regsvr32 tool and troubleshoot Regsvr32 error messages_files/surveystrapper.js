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
	};
}

if (!window._ms_support_fms_utility_setCookie)
{
	window._ms_support_fms_utility_setCookie = function(key, value, domain, expires, path)
	{
		domain = domain || document.domain;
		path = path || "/";
		document.cookie = key + "=" + value + "; domain=" + domain + "; path=" + path + "; " + (expires ? ("expires=" + expires.toGMTString() + ";") : "");
	};
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
	};
}

if (!window._ms_support_fms_utility_hasSubmitted)
{
	window._ms_support_fms_utility_hasSubmitted = function(hash)
	{
		var submitCookie = window._ms_support_fms_utility_getCookie("fmsSubmitted");
		var submitHistory = submitCookie ? submitCookie.split(",") : [];
		for (var i = 0; i < submitHistory.length; ++i)
		{
			if (submitHistory[i] == hash)
			{
				return true;
			}
		}
		return false;
	};
}

if (!window._ms_support_fms_utility_setSubmitted)
{
	window._ms_support_fms_utility_setSubmitted = function(hash, domain)
	{
		var submitCookie = window._ms_support_fms_utility_getCookie("fmsSubmitted");
		var submitHistory = submitCookie ? submitCookie.split(",") : [];
		for (var i = 0; i < submitHistory.length; ++i)
		{
			if (submitHistory[i] == hash)
			{
				return;
			}
		}
		submitHistory.push(hash);
		window._ms_support_fms_utility_setCookie("fmsSubmitted", submitHistory.join(","), domain);
	};
}

(function()
{
	var surveyConfig = typeof (_ms_support_fms_surveyConfig) != "undefined" ? _ms_support_fms_surveyConfig : false;
	if (surveyConfig)
	{
		var protocol = window.location.protocol;
		if (protocol != "http:" && protocol != "https:")
		{
			protocol = "https:";
		}
		if (surveyConfig.parameters)
		{
			for (var i = 0; i < surveyConfig.parameters.length; ++i)
			{
				if (typeof (surveyConfig.parameters[i]) != "undefined")
				{
					surveyConfig.parameters[i] = encodeURIComponent(unescape(surveyConfig.parameters[i]));
				}
			}
		}
		var config = {
			"target": unescape(surveyConfig.target),
			"template": surveyConfig.template || "default",
			"enableLTS": surveyConfig.enableLTS ? true : false,
			"survey": {
				"host": surveyConfig.survey.host || "support.microsoft.com",
				"language": surveyConfig.survey.language || "en",
				"id": parseInt(surveyConfig.survey.id) || 0,
				"isRTL": surveyConfig.survey.isRTL ? true : false,
				"features": surveyConfig.survey.features ? surveyConfig.survey.features : ["Title,Introduction,AllNormalPages,AllNavigationButtons,ProgressBar", "Thankyou,CloseButton"],
				"isInvitation": (surveyConfig.survey.isInvitation ? 1 : 0),
				"altStyle": surveyConfig.survey.altStyle || "",
				"renderOption": surveyConfig.survey.renderOption || "Default"
			},
			"parameters": surveyConfig.parameters ? surveyConfig.parameters : [],
			"site": {
				"name": surveyConfig.site.name,
				"culture": surveyConfig.site.culture || "en-us",
				"lcid": parseInt(surveyConfig.site.lcid) || 1033,
				"id": parseInt(surveyConfig.site.id) || 0,
				"brand": parseInt(surveyConfig.site.brand) || 0,
				"version": surveyConfig.site.version || "",
				"explicitDomain": (window.location.hostname != document.domain) || (surveyConfig.site.explicitDomain ? true : false),
				"cookieDomain": surveyConfig.site.cookieDomain || document.domain
			},
			"content": {
				"type": surveyConfig.content.type,
				"id": unescape(surveyConfig.content.id),
				"culture": surveyConfig.content.culture || "en-us",
				"lcid": parseInt(surveyConfig.content.lcid) || 1033,
				"properties": surveyConfig.content.properties || "",
				"aggregateId": unescape(surveyConfig.content.aggregateId)
			},
			"context": surveyConfig.context || {},
			"protocol": protocol,
			"triggerConfig": surveyConfig.triggerConfig
		};
		if (config.enableLTS)
		{
			if (typeof (window._ms_support_fms_ltsContentId) == "undefined")
			{
				window._ms_support_fms_ltsContentId = config.content.id;
			}
			else if (window._ms_support_fms_ltsContentId != config.content.id)
			{
				window._ms_support_fms_ltsContentId = config.content.aggregateId;
			}
		}
		function getPreventMultipleResponsesCookieKey(surveyId, language)
		{
			return ("fmspmr_" + surveyId + "_" + language).toUpperCase();
		}
		var key = getPreventMultipleResponsesCookieKey(config.survey.id, config.survey.language);
		var surveySuppressed = window._ms_support_fms_utility_getCookie(key) == "1";
		if (surveySuppressed)
		{
			if (config.enableLTS)
			{
				config.template = "lts";
			}
			else
			{
				return;
			}
		}
		function safeAdd(a, b)
		{
			var l = (a & 0xFFFF) + (b & 0xFFFF);
			var h = (a >> 16) + (b >> 16) + (l >> 16);
			return (h << 16) | (l & 0xFFFF);
		}
		function safeMultiply(a, b)
		{
			var la = a & 0xFFFF;
			var ha = (a >> 16);
			var lb = b & 0xFFFF;
			var hb = (b >> 16);
			return safeAdd(safeAdd(lb * la, (lb * ha) << 16), safeAdd(hb * la, (hb * ha) << 16) << 16);
		}
		function calculateStringHash(s)
		{
			var m1 = 0x15051505;
			var m2 = m1;
			var p = 0;
			var v;

			for (var i = s.length; i > 0; i -= 4)
			{
				v = s.charCodeAt(p);
				if (p < s.length - 1)
				{
					v |= s.charCodeAt(p + 1) << 16;
				}
				m1 = safeAdd(safeAdd((m1 << 5), m1), (m1 >>> 0x1b)) ^ v;
				if (i <= 2)
				{
					break;
				}
				p += 2;
				v = s.charCodeAt(p);
				if (p < s.length - 1)
				{
					v |= s.charCodeAt(p + 1) << 16;
				}
				m2 = safeAdd(safeAdd((m2 << 5), m2), (m2 >>> 0x1b)) ^ v;
				p += 2;
			}
			return (safeAdd(m1, safeMultiply(m2, 0x5d588b65))).toString(10);
		}
		var pack = window._ms_support_fms_utility_packageQueryString;
		config.hash = calculateStringHash([config.template, pack(config.parameters), pack(config.site), pack(config.content), pack(config.survey)].join(""));
		var hasSubmitted = false;
		var hash = window.location.hash.substring(1);
		if (hash && hash.split("_")[0] == config.hash)
		{
			window._ms_support_fms_utility_setSubmitted(config.hash, config.site.cookieDomain);
			hasSubmitted = true;
		}
		else
		{
			hasSubmitted = window._ms_support_fms_utility_hasSubmitted(config.hash);
		}
		config.feature = config.survey.features[(hasSubmitted ? 1 : 0)];
		var surveyHost = encodeURI(protocol + "//" + config.survey.host);
		var scid = "sw;" + config.survey.language + ";" + config.survey.id;
		config.survey.scid = scid;
		config.submissionGuid = ("s" + (new Date()).getTime());
		var target = document.getElementById(config.target);
		var callback = {
			"target": target,
			"config": config,
			"callback": function(getSurveyContent)
			{
				var content = getSurveyContent(this.config);
				var bookmarkName = this.config.hash + "_" + this.config.callbackId;
				var bookmark = "<a name=\"" + bookmarkName + "\"></a>";
				if (this.target)
				{
					this.target.innerHTML = bookmark + content;
					if (window.location.hash == "#" + bookmarkName)
					{
						if (this.target.scrollIntoView)
						{
							this.target.scrollIntoView();
						}
						else
						{
							window.location.hash = bookmarkName;
						}
					}
				}
				else
				{
					document.write(bookmark + content);
				}
			}
		};
		config.callbackId = window._ms_support_fms_surveyScriptHandlerCallback.push(callback) - 1;
		var params = {
			Host: surveyHost,
			CallbackId: config.callbackId,
			EnableLTS: config.enableLTS,
			IsInvitation: config.survey.isInvitation,
			Feature: config.feature,
			Scid: scid,
			SurveyLangCode: config.survey.language,
			SurveyId: config.survey.id,
			Site: config.site.name,
			Region: config.site.culture,
			SiteCulture: config.site.culture,
			SsId: config.site.id,
			SiteBrandId: config.site.brand,
			SsVersion: config.site.version,
			Params: config.parameters,
			ParamLength: config.parameters.length,
			SubmissionGuid: config.submissionGuid,
			ContentType: config.content.type,
			ContentCulture: config.content.culture,
			ContentId: config.content.id,
			ContentLcid: config.content.lcid,
			ContentProperties: config.content.properties,
			Hash: config.hash + "_" + config.callbackId
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
			return {
				query: queryString,
				overflow: overflow
			};
		}
		var MAX_URL_LENGTH = 2048;
		var MAX_INT_LENGTH = 23;
		var timestamp = (new Date()).getTime().toString(32);
		var randomNumber = Math.random().toString(32).substring(2);
		var chunkHeadLength = ("&ts=" + timestamp + "&rnd=" + randomNumber + "&chkc=").length + MAX_INT_LENGTH;
		var handler_src_base = surveyHost + "/common/surveyscripthandler.ashx?template=" + encodeURIComponent(config.template);
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
			return {
				baseUrl: dsBaseUrl,
				timestamp: timestamp,
				randomNumer: randomNumber,
				valuePrefix: valuePrefix,
				chunks: chunks
			};
		}
		var hasChunks = queryString.overflow ? true : false;
		if (hasChunks)
		{
			var chunks = packageChunks(queryString.overflow, surveyHost);

			function sendChunksAsync(onComplete)
			{
				var count = 0;
				if (!window._ms_support_fms_dataStorage_response)
				{
					window._ms_support_fms_dataStorage_response = [];
				}
				window._ms_support_fms_dataStorage_response[timestamp + "" + randomNumber] = {
					callback: function(chunkId)
					{
						++count;
						if (count == 1)
						{
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
							onComplete("&ts=" + timestamp + "&rnd=" + randomNumber + "&chkc=" + chunks.chunks.length);
						}
					}
				};
				var handler = document.createElement("script");
				handler.src = chunks.baseUrl + 0 + chunks.valuePrefix + chunks.chunks[0];
				handler.type = "text/javascript";
				(document.body || document.documentElement).appendChild(handler);
			}
			function sendChunksSync()
			{
				for (var i = 0; i < chunks.chunks.length; ++i)
				{
					document.write("<script type=\"text/javascript\" src=\"" + chunks.baseUrl + i + chunks.valuePrefix + chunks.chunks[i] + "\"></s" + "cript>");
				}
				return "&ts=" + timestamp + "&rnd=" + randomNumber + "&chkc=" + chunks.chunks.length;
			}
		}
		if (target)
		{
			function addScriptHandler()
			{
				var handler = document.createElement("script");
				handler.src = handler_src;
				handler.type = "text/javascript";
				(document.body || document.documentElement).appendChild(handler);
			}

			var addEventHandler;

			if (window.addEventListener)
			{
				addEventHandler = function(el, ev, fp)
				{
					el.addEventListener(ev, fp, false);
				};
			}
			else if (window.attachEvent)
			{
				addEventHandler = function(el, ev, fp)
				{
					el.attachEvent("on" + ev, fp);
				};
			}
			else
			{
				addEventHandler = function(el, ev, fp)
				{
				};
			}

			var e = typeof (document.onreadystatechange) != "undefined" ? [document, "readystatechange"] : (typeof (document.onDOMContentLoaded) != "undefined" ? [document, "DOMContentLoaded"] : [window, "load"]);
			addEventHandler(e[0], e[1], function(e)
			{
				if (typeof (document.readyState) == "undefined" || document.readyState == "complete")
				{
					if (hasChunks)
					{
						sendChunksAsync(addScriptHandler);
					}
					else
					{
						addScriptHandler();
					}
				}
			});
		}
		else
		{
			var extension = hasChunks ? sendChunksSync() : "";
			document.write("<script type=\"text/javascript\" src=\"" + handler_src + extension + "\"></s" + "cript>");
		}
	}
})();