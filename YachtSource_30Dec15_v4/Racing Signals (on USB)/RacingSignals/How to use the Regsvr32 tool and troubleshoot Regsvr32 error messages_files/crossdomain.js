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

if (!MS.Support.Fms.CrossDomain)
{
    MS.Support.Fms.CrossDomain = function()
    {
        var anchorClicked = false;
        var blessedDomains = [];
        var parentPollingTimer = null;

        window.crossDomainInitialized = 1;

        function getCookie(key)
        {
            var value = document.cookie;
            var start = value.indexOf(" " + key + "=");
            if (start == -1)
            {
                start = value.indexOf(key + "=");
            }
            if (start == -1)
            {
                value = null;
            }
            else
            {
                start = value.indexOf("=", start) + 1;
                var end = value.indexOf(";", start);
                if (end == -1)
                {
                    end = value.length;
                }
                value = unescape(value.substring(start, end));
            }
            return value;
        }

        function setCookie(key, value, expiryDays, domain)
        {
            var expiryDate = new Date();
            expiryDate.setDate(expiryDate.getDate() + expiryDays);
            document.cookie = key + "=" + escape(value) + "; domain=" + domain + ((expiryDays == null) ? "" : "; expires=" + expiryDate.toUTCString()) + "; path=/";
        }

        function isCookieEnabled()
        {
            setCookie("testcookie", "testvalue", null, window.cookieDomain);

            if (getCookie("testcookie") == "testvalue")
                return true;

            return false;
        }

        function isIE8()
        {
            try
            {
                if (navigator.appName == 'Microsoft Internet Explorer')
                {
                    if (navigator.appVersion.indexOf("MSIE 8") != -1)
                    {
                        return true;
                    }
                }
            } catch (e) { }
            return false;
        }

        function isIE9()
        {
            try
            {
                if (navigator.appName == 'Microsoft Internet Explorer')
                {
                    if (navigator.appVersion.indexOf("MSIE 9") != -1)
                    {
                        return true;
                    }
                }
            } catch (e) { }
            return false;
        }

        function isIE10()
        {
            try
            {
                if (navigator.appName == 'Microsoft Internet Explorer')
                {
                    if (navigator.appVersion.indexOf("MSIE 10") != -1)
                    {
                        return true;
                    }
                }
            } catch (e) { }
            return false;
        }

        function isChrome()
        {
            try
            {
                if (window.chrome)
                {
                    return true;
                }
            } catch (e) { }
            return false;
        }

        function getChromeVersion()
        {
            var ver = 0;
            try
            {
                ver = parseInt(window.navigator.appVersion.match(/Chrome\/(\d+)\./)[1], 10);
            } catch (e) { }
            return ver;
        }

        function isIE11()
        {
            try
            {
                if (navigator.userAgent.match(/Trident.*rv.*11\./))
                {
                    return true;
                }
            } catch (e) { }
            return false;
        }

        function isMozillaFirefox()
        {
            try
            {
                if (navigator.userAgent.toLowerCase().indexOf('firefox') > -1)
                    return true;
            }
            catch (e) { }
            return false;
        }

        function isSafari()
        {
            try
            {
                if (navigator.userAgent.search("Safari") > -1 && navigator.userAgent.search("Chrome") < 0)
                    return true;
            }
            catch (e) { }
            return false;
        }

        function getDomainFromLocation(location)
        {
            var domain = $('<a>').prop('href', location).prop('hostname').toLowerCase();
            if (domain.match(/^www\./))
            {
                domain = domain.substring(4);
            }
            return domain;
        }

        function isBlessedDomain(location)
        {
            try
            {
                var domain = getDomainFromLocation(location);
                if (blessedDomains.indexOf(domain) != -1)
                {
                    return true;
                }
            } catch (e) { }
            return false;
        }

        function isFMSDomain(location)
        {
            try
            {
                var domain = getDomainFromLocation(location);
                if ($.inArray(domain, window.hostedDomains.split(';')) != -1)
                {
                    return true;
                }
            } catch (e) { }
            return false;
        }

        function ispostMessageSupportedBrowser()
        {
            if (isIE8() || isIE9() || isIE10() || isChrome() || isIE11() || isMozillaFirefox() || isSafari())
            {
                return true;
            }
            return false;
        }

        function updateSessionStorage(protocol, hostname)
        {
            try
            {
                var sta = getCookie("P_STA");
                if (!sta)
                {
                    return;
                }

                var hd = getCookie('fmshd');
                if (hd && $.inArray(protocol + "//" + hostname, hd.split(';')) != -1)
                {
                    return;
                }
                else
                {
                    if ($.inArray(hostname, window.hostedDomains.split(';')) != -1)
                    {
                        if (window.sessionStorage)
                        {
                            var newDate = new Date();
                            window.sessionStorage.tabSessionID = newDate.getTime();
                        }
                        hd = ((hd == null) ? "" : hd) + protocol + "//" + hostname + ";";
                        setCookie('fmshd', hd, null, window.cookieDomain);
                    }
                }
            } catch (e) { }
        }

        function removeSessionStorage(protocol, hostname)
        {
            var hd = getCookie('fmshd');
            if (!hd)
            {
                return;
            }
            var hds = hd.split(';');
            var index = hds.indexOf(protocol + "//" + hostname);
            if (index > -1)
            {
                if (window.sessionStorage)
                {
                    window.sessionStorage.removeItem('tabSessionID');
                }
                hds.splice(index, 1);
                setCookie('fmshd', hds.join(';'), null, window.cookieDomain);
            }
        }

        function sendLocationtoChild(location)
        {
            if (ispostMessageSupportedBrowser())
            {
                var sta = getCookie("P_STA");
                if (sta == null)
                {
                    sta = 0;
                }
                if (sta != 0)
                {
                    try
                    {
                        if (sessionStorage.tabSessionID)
                        {
                            try
                            {
                                popup = window.open("", "trackingWindow");
                            }
                            catch (e)
                            {
                                removeSessionStorage(window.location.protocol, window.location.hostname);
                                return;
                            }
                            try
                            {
                                if (typeof (popup) == "undefined" || popup == null)
                                {
                                    removeSessionStorage(window.location.protocol, window.location.hostname);
                                    return;
                                }
                                else if (popup.location.href == "about:blank")
                                {
                                    popup.close();
                                    removeSessionStorage(window.location.protocol, window.location.hostname);
                                    return;
                                }
                            } catch (e) { }
                            if (isIE8() || isIE9())
                            {
                                popup.postMessagePassthrough(location);
                            }
                            else
                            {
                                popup.postMessage(location, "*");
                            }

                            if (isChrome() && getChromeVersion() < 33)
                            {
                                if (isBlessedDomain(location) || isFMSDomain(location))
                                {
                                    popup.window.open("about:blank").close();
                                }
                            }

                            if (parentPollingTimer != null)
                            {
                                clearTimeout(parentPollingTimer);
                                parentPollingTimer = null;
                            }                            
                            parentPollingTimer = window.setTimeout(function() { sendLocationtoChild(window.location.href); }, (window.parentPollingTimeout - window.parentPollingDelay));
                        }
                    }
                    catch (e) { }
                }
            }
        }

        if (ispostMessageSupportedBrowser() && isCookieEnabled() && window.blessedDomains)
        {

            if (window.blessedDomains && typeof (Object) != "undefined" && Object.keys)
            {
                blessedDomains = Object.keys(window.blessedDomains);
            }

            if (!sessionStorage.tabSessionID)
            {
                updateSessionStorage(window.location.protocol, window.location.hostname);
            }

            sendLocationtoChild(window.location.href);

            $(window.document).on("click", "a", function()
            {
                try
                {
                    if (event.ctrlKey || event.shiftKey)
                    {
                        return;
                    }
                } catch (e) { }
                var href = this.href;
                if (!href.match(/^javascript:/))
                {
                    anchorClicked = true;
                    sendLocationtoChild(href);
                }
            })

            $(window).on('beforeunload', function()
            {
                if (!anchorClicked)
                {
                    try
                    {
                        if ($('#oasp_clientcomponent').length > 0)
                        {
                            var signinRedirect = $('#oasp_clientcomponent').contents().find('.signinRedirect');
                            if (signinRedirect.length > 0 && signinRedirect.attr('href'))
                            {
                                sendLocationtoChild(signinRedirect.attr('href'));
                                return;
                            }
                        }
                    }
                    catch (e)
                    { }
                    sendLocationtoChild("CheckShowSurvey");
                }
            });
        }
    }
}

if (!window.crossDomainInitialized)
{
    new MS.Support.Fms.CrossDomain();
}
