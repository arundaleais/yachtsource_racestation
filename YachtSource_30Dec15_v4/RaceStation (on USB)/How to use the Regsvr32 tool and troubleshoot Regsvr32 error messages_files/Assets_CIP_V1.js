(function () {
    var downloadJS = function (jsSrc) {
        var script = document.createElement("script");
        script.src = jsSrc;
        script.type = 'text/javascript';
        document.getElementsByTagName("head")[0].appendChild(script);
    }

    var downloadCSS = function ($, cssSrc) {
        $('<link/>', {
            rel: 'stylesheet',
            type: 'text/css',
            href: cssSrc
        }).appendTo("head");
    }
    // $.fn.jquery == '1.3.2' or "1.5.1" is to be compatible with Bemis and VKB sites.
    var bemisjQueryVersion = '1.3.2';
    var vkbjQueryVersion = '1.5.1';
    var prefix = window.location.protocol == "http:" ? "http:" : "https:";

    if ((typeof ($) == 'undefined' && typeof (jQuery) == 'undefined') || $.fn.jquery == bemisjQueryVersion || $.fn.jquery == vkbjQueryVersion) {
        downloadJS(prefix + '//ajax.aspnetcdn.com/ajax/jQuery/jquery-1.7.2.min.js');
    }

    var AssetCIPExecute = function (callback) {
        if (window.jQuery && $.fn.jquery != bemisjQueryVersion && $.fn.jquery != vkbjQueryVersion) {
            callback(jQuery);
        }
        else {
            window.setTimeout(function () { AssetCIPExecute(callback); }, 100);
        }
    };

    AssetCIPExecute(function ($) {

        (function ($) {
            if ($('link[href*="http://support.microsoft.com/library/StyleSheet/support/en-US/Assets_CIP_Common.css"]').length == 0) {
                downloadCSS($, prefix + "//support.microsoft.com/library/StyleSheet/support/en-US/assets_DynamicContentBar.css");
                downloadCSS($, prefix + "//support.microsoft.com/library/StyleSheet/support/en-US/Assets_CIP_Common.css");
                downloadJS('//support.microsoft.com/library/JavaScript/support/en-US/assets_os_browser_detect_script.js');
                downloadJS('//support.microsoft.com/library/JavaScript/support/en-US/assets_os_browser_detect_configuration.js');
                downloadJS('//support.microsoft.com/library/JavaScript/support/en-US/assets_DynamicContentBar.js');
            }
        })($);

        //************ 1.Assets_CIP Start ********************
        (function ($) {

            $.Assets_CIP = function () {
                $.Assets_CIP.CIP_Patterns = {
                    "Mats": "//support.microsoft.com/library/images/support/en-US/assets_mats1.png",
                    "Fixit": "//support.microsoft.com/library/images/support/en-US/assets_fixit1.png",
                    "ThirdParty": "//support.microsoft.com/library/images/support/en-US/assets_thirdParty1.png",
                    "SolutionWizard": "//support.microsoft.com/library/images/support/en-US/assets_solutionWizard1.png",
                    "VideoPlayer": "//support.microsoft.com/library/images/support/en-US/assets_video1.png"
                };
                $.Assets_CIP.CurrentSiteTypeFlag = $.Assets_CIP.GetSiteType();
                $.Assets_CIP.CurrentSiteLanguage = $.Assets_CIP.GetSiteLanguage();
                $.Assets_CIP.CurrentSiteDir = $.MSCOSCodeFree.GetDir();
                $.Assets_CIP.CIPImgArray = new Array();

                $('img[src*="//support.microsoft.com/library/images/support/en-US/assets_"]').each(function () {
                    var currentImg = this;
                    var imgSrc = $(currentImg).attr("src");
                    var pattern;
                    $.each($.Assets_CIP.CIP_Patterns, function (key, value) {
                        pattern = new RegExp(value, "gi");
                        if (pattern.test(imgSrc)) {
                            ProcessImg(currentImg, key, $.Assets_CIP.CIPImgArray);
                        }
                    });
                });
                DeleteCIPImgArray($.Assets_CIP.CIPImgArray);
                $.Assets_CIP.Folding();
            };

            var ProcessImg = function (ImgObject, TypeofCIP, CIPImgArray) {
                $.Assets_CIP.CIPImgArray.push(ImgObject);
                switch (TypeofCIP) {
                    case "Mats": $.Assets_CIP.Create_One_Mats(ImgObject);
                        break;
                    case "Fixit": $.Assets_CIP.Create_One_Fixit(ImgObject);
                        break;
                    case "ThirdParty": $.Assets_CIP.Create_One_ThirdParty(ImgObject);
                        break;
                    case "SolutionWizard": $.Assets_CIP.Create_One_SolutionWizard(ImgObject);
                        break;
                    case "VideoPlayer": $.Assets_CIP.Create_One_VideoPlayer(ImgObject);
                        break;
                    default:
                }
            };
            var DeleteCIPImgArray = function (CIPImgArray) {
                while (CIPImgArray.length) {
                    $.Assets_CIP.DeleteHTMLRangeofCIP(CIPImgArray.pop());
                }
            };
        })($);
        //************ 1.Assets_CIP End ********************

        //************ 2.Common Start ********************
        (function ($) {
            $.Assets_CIP.GetSiteType = function () {
                var siteTypeFlag;
                var locationURL = window.location.href;
                var reg = new RegExp("http://bemis", "gi");
                if (reg.test(locationURL)) {
                    siteTypeFlag = "bemis";
                }
                else {
                    reg = new RegExp("(http://vkb)|(https://vkb)|(http://visualkb)|(https://visualkb)", "gi");
                    if (reg.test(locationURL)) {
                        siteTypeFlag = "vkb"
                    }
                    else {
                        reg = new RegExp("(http://smallbusiness.support.microsoft.com)|(https://smallbusiness.support.microsoft.com)", "gi");
                        if (reg.test(locationURL)) {
                            siteTypeFlag = "ssb";
                        }
                        else {
                            reg = new RegExp("microsoft.com|(157.56.56.32)", "gi");
                            if (reg.test(locationURL)) {
                                siteTypeFlag = "smc";
                            }
                        }
                    }
                }
                return siteTypeFlag;
            };
            $.Assets_CIP.GetSiteLanguage = function () {
                var language = navigator.browserLanguage || navigator.userLanguage;
                if ($.Assets_CIP.GetSiteType() == "bemis") {
                    var $nobr = $('nobr:contains("Language:")').first();
                    if ($nobr.length > 0) {
                        if ($nobr.parent().parent()[0].outerText != undefined) {
                            language = $nobr.parent().parent()[0].outerText.split(":")[1].toLowerCase();
                        }
                        else {
                            language = "en-us";
                        }
                    }
                    return $.trim(language.toLowerCase());
                }

                $metaLanguage = $('meta[http-equiv*="content-language"]').first();
                if ($metaLanguage.length > 0) {
                    language = $metaLanguage.attr("content");
                    return $.trim(language.toLowerCase());
                }

                var pattern = new RegExp("(?:kb|gw)/[^/]+/(.+)", "gi");
                var matches = pattern.exec(window.location.href);
                if (matches != null) {
                    language = matches[1].toLocaleLowerCase();
                }
                else {
                    pattern = new RegExp("(?:.microsoft.com|157.56.56.32)/(.+)/(?:kb|gw)/[^/]+", "gi");
                    matches = pattern.exec(window.location.href);
                    if (matches != null) {
                        language = matches[1].toLocaleLowerCase();
                    }
                    else {
                        var languageCookie = "gssLANG";
                        if (document.cookie && document.cookie != '') {
                            var cookies = document.cookie.split(';');
                            for (var i = 0; i < cookies.length; i++) {
                                var cookie = $.trim(cookies[i]);
                                if (cookie.substring(0, languageCookie.length + 1) == (languageCookie + '=')) {
                                    if (cookie.length > languageCookie.length + 1) {
                                        language = decodeURIComponent(cookie.substring(languageCookie.length + 1));
                                    }
                                }
                            }
                        }
                    }
                }
                if (language == null && BrowserDetect.browser != 'Firefox') {
                    language = 'en-us';
                }
                return $.trim(language.toLowerCase());
            };
            $.Assets_CIP.DeleteAllAssetsImages = function () {
                $('img[src*="//support.microsoft.com/library/images/support/en-US/assets_"]').each(function () {
                    switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                        case "ssb":
                        case "smc":
                            $(this).parent().parent().remove();
                            break;
                        case "bemis":
                            $(this).parent().remove();
                            break;
                        case "vkb":
                            $(this).remove();
                            break;
                    }
                    $(this).parent().parent().remove();
                });
            };

            //Get the HTML Code Segment of CIP Element
            $.Assets_CIP.GetHTMLRangeofCIP = function (CurrentImage) {
                var HTMLStr;
                var startTag;
                switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                    case "ssb":
                    case "smc":
                        startTag = CurrentImage.parentNode.parentNode;
                        HTMLStr = GetHTMLRangeofCIP_smc(startTag);
                        break;
                    case "bemis": startTag = CurrentImage.parentNode;
                        HTMLStr = GetHTMLRangeofCIP_bemis(startTag);
                        break;
                    case "vkb": startTag = CurrentImage;
                        HTMLStr = GetHTMLRangeofCIP_vkb(startTag);
                        break;
                    default:
                }
                return HTMLStr;
            };
            var GetHTMLRangeofCIP_smc = function (startTag) {
                var HTMLStr = "</div>";
                var count = 20;
                while (count--) {
                    if (startTag.nextSibling.nodeName == "#text") {
                        HTMLStr += startTag.nextSibling.nodeValue;
                        startTag = startTag.nextSibling;
                    }
                    else if (startTag.nextSibling.nodeName != "DIV") {
                        HTMLStr += startTag.nextSibling.outerHTML;
                        startTag = startTag.nextSibling;
                    }
                    else if (startTag.nextSibling.nodeName == "DIV" && startTag.nextSibling.getAttribute("class") == "pLink") {
                        startTag = startTag.nextSibling;
                    }
                    else {
                        break;
                    }
                }
                return HTMLStr;
            };
            var GetHTMLRangeofCIP_bemis = function (startTag) {
                var HTMLStr = "</div>";
                var count = 20;
                while (count--) {
                    if (startTag.nextSibling.nodeName == "#text") {
                        HTMLStr += startTag.nextSibling.nodeValue;
                        startTag = startTag.nextSibling;
                    }
                    else if (startTag.nextSibling.nodeName != "DIV") {
                        HTMLStr += startTag.nextSibling.outerHTML;
                        startTag = startTag.nextSibling;
                    }
                    else {
                        break;
                    }
                }
                return HTMLStr;
            };
            var GetHTMLRangeofCIP_vkb = function (startTag) {
                var HTMLStr = "</div>";
                var count = 20;
                while (count--) {
                    if (startTag.nextSibling.nodeName == "#text") {
                        HTMLStr += startTag.nextSibling.nodeValue;
                        startTag = startTag.nextSibling;
                    }
                    else if (startTag.nextSibling.nodeName != "IMG") {
                        HTMLStr += startTag.nextSibling.outerHTML;
                        startTag = startTag.nextSibling;
                    }
                    else {
                        break;
                    }
                }
                return HTMLStr;
            };

            //Delete the HTML Code Segment of CIP_Element
            $.Assets_CIP.DeleteHTMLRangeofCIP = function (CurrentImage) {
                var startTag;
                var count = 20;
                switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                    case "ssb":
                    case "smc":
                        startTag = CurrentImage.parentNode.parentNode;
                        DeleteHTMLRangeofCIP_smc(startTag);
                        break;
                    case "bemis": startTag = CurrentImage.parentNode;
                        DeleteHTMLRangeofCIP_bemis(startTag);
                        break;
                    case "vkb": startTag = CurrentImage;
                        DeleteHTMLRangeofCIP_vkb(startTag);
                        break;
                    default:
                }
            };
            var DeleteHTMLRangeofCIP_smc = function (startTag) {
                var count = 20;
                while (count--) {
                    if (startTag.nextSibling.nodeName != "DIV") {
                        startTag.parentNode.removeChild(startTag.nextSibling);
                    }
                    else if (startTag.nextSibling.nodeName == "DIV" && startTag.nextSibling.getAttribute("class") == "pLink") {
                        startTag.parentNode.removeChild(startTag.nextSibling);
                    }
                    else if (startTag.nextSibling.nodeName == "DIV" && startTag.nextSibling.getAttribute("class") == "kb_outergraphicwrapper kb_outergraphicwrapper_closed") {
                        startTag.parentNode.removeChild(startTag.nextSibling);
                        startTag.parentNode.removeChild(startTag);
                        break;
                    }
                }
            };
            var DeleteHTMLRangeofCIP_bemis = function (startTag) {
                var count = 20;
                while (count--) {
                    if (startTag.nextSibling.nodeName != "DIV") {
                        startTag.parentNode.removeChild(startTag.nextSibling);
                    }
                    else {
                        startTag.parentNode.removeChild(startTag.nextSibling);
                        startTag.parentNode.removeChild(startTag);
                        break;
                    }
                }
            };
            var DeleteHTMLRangeofCIP_vkb = function (startTag) {
                var count = 20;
                while (count--) {
                    if (startTag.nextSibling.nodeName != "IMG") {
                        startTag.parentNode.removeChild(startTag.nextSibling);
                    }
                    else {
                        startTag.parentNode.removeChild(startTag.nextSibling);
                        startTag.parentNode.removeChild(startTag);
                        break;
                    }
                }
            };

            //inert an asset
            $.Assets_CIP.InsertObject = function (currentImg, objectElement) {
                switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                    case "ssb":
                    case "smc":
                        currentImg.parentNode.parentNode.parentNode.insertBefore(objectElement, currentImg.parentNode.parentNode);
                        break;
                    case "bemis": currentImg.parentNode.parentNode.insertBefore(objectElement, currentImg.parentNode);
                        break;
                    case "vkb": currentImg.parentNode.insertBefore(objectElement, currentImg);
                        break;
                    default:
                }
            };
        })($);
        //************ 2.Common End ********************

        //************ 3.Fixit Start ********************
        (function ($) {
            $.Assets_CIP.Create_One_Fixit = function (currentImg) {
                var info_up = "";
                var info_down = "";
                var link_url = "";

                var HTMLRangeofFixit = $.Assets_CIP.GetHTMLRangeofCIP(currentImg);
                pattern = new RegExp("<a[^>]* href=\"([^\"]*)[\\w\\W]*\"?>([\\w\\W]*?)</a>", "gi");
                var matches = pattern.exec(HTMLRangeofFixit);
                if (matches != null) {
                    link_url = matches[1].replace(new RegExp("<br>", "gi"), "");
                    info_up = matches[2].replace(new RegExp("<br>", "gi"), "");
                }

                pattern = new RegExp("<i>([\\w\\W]*?)</i>", "gi");
                matches = pattern.exec(HTMLRangeofFixit);
                if (matches != null) {
                    info_down = matches[1].replace(new RegExp("<br>", "gi"), "");
                }
                if ($.Assets_CIP.CurrentSiteTypeFlag == "vkb") {
                    matches = pattern.exec(HTMLRangeofFixit);
                    if (matches != null) {
                        info_down = matches[1].replace(new RegExp("<br>", "gi"), "");
                    }
                }

                var div = document.createElement("div");
                div.className = "Fixit_Div";

                //Create <a class="Fixit_Button" href="http://go.microsoft.com/?linkid=9774319" target="_blank" title="Microsoft Fix it" />
                var a_picture = document.createElement("a");
                a_picture.className = "Fixit_Button";
                a_picture.setAttribute("href", link_url);
                a_picture.setAttribute("target", "_blank");
                a_picture.setAttribute("title", "Microsoft Fix it");

                //Create <img alt="Microsoft Fix it" src="http://support.microsoft.com/library/images/support/en-US/FixItButton1.jpg" />, and append it to the link.
                var a_img = document.createElement("img");
                a_img.src = prefix + "//support.microsoft.com/library/images/support/en-US/FixItButton1.jpg";
                a_img.width = "139";
                a_img.height = "56";
                a_img.setAttribute("alt", "Fix this problem");
                a_picture.appendChild(a_img);

                div.appendChild(a_picture);

                //Create <a href="http://go.microsoft.com/?linkid=9774319" target="_blank" title="Microsoft Fix it">Fix error 0x80004002</a>
                var a_text = document.createElement("a");
                a_text.setAttribute("href", link_url);
                a_text.setAttribute("target", "_blank");
                a_text.setAttribute("title", "Microsoft Fix it");
                var a_content = document.createTextNode(info_up);
                a_text.appendChild(a_content);
                div.appendChild(a_text);

                //Create <br />
                var br = document.createElement("br");
                div.appendChild(br);

                //Create "Microsoft Fix it 50687"
                //var text = document.createTextNode("Microsoft Fix it 50687");
                var text = document.createTextNode(info_down);
                div.appendChild(text);

                //Fixit.appendChild(div);
                //Compatibile with IE7 and IE8
                var outDiv = document.createElement("div");
                outDiv.appendChild(div);
                $.Assets_CIP.InsertObject(currentImg, outDiv);
            };
        })($);
        //************ 3.Fixit End ********************

        //************ 4.Mats Start ********************
        (function ($) {
            $.Assets_CIP.Create_One_Mats = function (currentImg) {

                var run_now_url = "";
                var learn_more_url = "";
                var p1 = "";
                var p2 = "";
                var learnMore = "";
                var runNow = "";
                var HTMLRangeofMats = "";
                var HTMLRangeofMats = $.Assets_CIP.GetHTMLRangeofCIP(currentImg);

                //Get learn_more_url
                pattern = new RegExp("<u[^>]*><a[^>]* href=\"([^\"]*)[\\w\\W]*\"?>", "gi");
                if ($.Assets_CIP.CurrentSiteTypeFlag == "vkb") {
                    pattern = new RegExp("<a[^>]* href=\"([^\"]*)[\\w\\W]*\"?>", "gi");
                }
                var matches = pattern.exec(HTMLRangeofMats);
                if (matches != null) {
                    learn_more_url = matches[1];
                }

                //Get run_now_url
                pattern = new RegExp("<i[^>]*><a[^>]* href=\"([^\"]*)[\\w\\W]*\"?>", "gi");
                matches = pattern.exec(HTMLRangeofMats);
                if (matches != null) {
                    run_now_url = matches[1];
                }

                //Get p1
                pattern = new RegExp("</div>([\\w\\W]*?)<i>", "gi");
                matches = pattern.exec(HTMLRangeofMats);
                if (matches != null) {
                    p1 = matches[1];
                }

                //Get p2
                pattern = new RegExp("<i>([\\w\\W]*?)</i>", "gi");
                matches = pattern.exec(HTMLRangeofMats);
                if (matches != null) {
                    p2 = matches[1];
                }

                //Get learnMore
                pattern = new RegExp("<a([\\w\\W]*?)>([\\w\\W]*?)</a>", "gi");
                matches = pattern.exec(HTMLRangeofMats);
                if (matches != null) {
                    learnMore = matches[2];
                }

                //Get runNow
                matches = pattern.exec(HTMLRangeofMats);
                if (matches != null) {
                    runNow = matches[2];
                }

                var MatsDiv = document.createElement("div");
                MatsDiv.innerHTML = GetMatsString(run_now_url, learn_more_url, p1, p2, learnMore, runNow);
                $.Assets_CIP.InsertObject(currentImg, MatsDiv);
            };
            var GetMatsString = function (run_now_url, learn_more_url, p1, p2, learnMore, runNow) {
                var MatsString = "";
                MatsString = "    <div>";
                MatsString += "        <div id=\"rnp_launch_container\" style=\"width: 554px; display: block;\">";
                MatsString += "            <img src=\"" + prefix + "//support.microsoft.com/library/images/support/en-US/fixit_launchbanner.png\" alt=\"Microsoft fix it banner\" width=\"554\" height=\"54\" />";
                MatsString += "            <div style=\"background-color: #eff7fb; border-color: #2c4b78; border-width: 1px;border-style: solid; padding-top: 10px; padding-bottom: 15px; text-align: left;\">";
                MatsString += "                <table>";
                MatsString += "                    <tr>";
                MatsString += "                        <td style=\"border-right-style: solid; border-right-width: 1px; border-right-color: #C0C0C0; width:90%\">";
                MatsString += "                            <p style=\"margin: 0px 10px 0px 10px; font: normal 11px/14px;\">" + p1 + "</p><br \/>";
                MatsString += "                            <p style=\"margin: 0px 10px 0px 10px; font: normal 11px/14px;\">";
                MatsString += "                                " + p2 + " <a href=\"" + learn_more_url + "\">" + learnMore + "</a><\/p>";
                MatsString += "                        <\/td>";
                MatsString += "                        <td style=\"padding: 0 8px\" >";
                MatsString += "                                                <a id=\"L_2321134\" style=\"cursor: pointer; text-decoration: none; color: white; font-size: 12pt; font-weight: bold; white-space: nowrap;\" onclick=\"javascript:{MS_HandleClick(this, \'L_2321134\', true);saveMatsSessionToCookie();}\" href=\"" + run_now_url + "\" rel=\"nofollow\">" + runNow + " <\/a>";
                MatsString += "                        <\/td>";
                MatsString += "                    <\/tr>";
                MatsString += "                <\/table>";
                MatsString += "            <\/div>";
                MatsString += "        <\/div>";
                MatsString += "    <\/div>";
                MatsString += "    <div>";
                MatsString += "        <div id=\"rnp_launch_container_disabled\" style=\"width: 554px; display: none;\">";
                MatsString += "            <img src=\"" + prefix + "//support.microsoft.com/library/images/support/en-US/fixit_launchbanner.png\" alt=\"Microsoft fix it banner\" width=\"554\" height=\"54\" />";
                MatsString += "            <div style=\"background-color: #eff7fb; border-color: #2c4b78; border-width: 1px; border-style: solid; padding-top: 10px; padding-bottom: 15px; text-align: left;width: width: 554px;\">";
                MatsString += "                <table>";
                MatsString += "                    <tr>";
                MatsString += "                        <td>";
                MatsString += "                            <p style=\"margin: 0px 10px 0px 10px; font: normal 11px\/14px;\">";
                MatsString += "                                The Fix Internet Explorer add-on problems when IE hangs or freezes troubleshooter";
                MatsString += "                                may automatically fix the problem described in this article.";
                MatsString += "                            </p>";
                MatsString += "                            <br \/>";
                MatsString += "                            <p style=\"margin: 0px 10px 0px 10px; font: normal 11px\/14px;\">";
                MatsString += "                                This troubleshooter fixes many problems. Learn more";
                MatsString += "                            <\/p>";
                MatsString += "                        <\/td>";
                MatsString += "                    <\/tr>";
                MatsString += "                    <tr>";
                MatsString += "                        <td style=\"padding: 0 233px\" >";
                MatsString += "                            <a id=\"L_2321134\" style=\"cursor: nw-resize; text-decoration: none; color: white; font-size: 12pt; font-weight: bold; white-space: nowrap;\" onclick=\"javascript;\" rel=\"nofollow\">Run now<\/a>";
                MatsString += "                        <\/td>";
                MatsString += "                    <\/tr>";
                MatsString += "                <\/table>";
                MatsString += "                <br \/>";
                MatsString += "                <div class=\"rnp_pna_error\">";
                MatsString += "                    <div class=\"rnp_pna_error_icon_message\">";
                MatsString += "                        <img alt=\"error icon\" src=\"" + prefix + "//support.microsoft.com/library/images/support/cn/icon_error_sml.png\" \/>";
                MatsString += "                        <strong>We\'re sorry, but your operating system is not supported by Microsoft Automated Troubleshooting Services at this time. <\/strong>";
                MatsString += "                    <\/div>";
                MatsString += "                <\/div>";
                MatsString += "            <\/div>";
                MatsString += "        <\/div>";
                MatsString += "    <\/div>";
                return MatsString;
            };
            saveMatsSessionToCookie = function () {
                var gsfxcookie = fetchcookieval("GsfxSessionCookie");
                var matsrun_sessionid = gsfxcookie + "_" + StatsDotNet.eventSeqNo + "_KB";
                setcookieval("matsrun_sessionid", matsrun_sessionid, null, false);
            };
            var Mats_Choice = function () {
                if (document.location.hash.substring(1) == "noJavaScript") {
                }
                else {
                    if (navigator.userAgent.toLowerCase().indexOf("windows nt") <= 1) {
                        document.getElementById("rnp_launch_container_disabled").style.display = 'block';
                        document.getElementById("rnp_launch_container").style.display = 'none';
                    }
                }
            };
        })($);
        //************ 4.Mats End ********************

        //************ 5.SolutionWizard Start ********************
        (function ($) {
            $.Assets_CIP.Create_One_SolutionWizard = function (currentImg) {
                var HTMLRangeofSolutionWizard = $.Assets_CIP.GetHTMLRangeofCIP(currentImg);
                var linkUrl = "";
                var linkText = "";

                var pattern = new RegExp("<a[^>]* href=\"([^\"]*)[\\w\\W]*\"?>([\\w\\W]*?)</a>", "gi");
                var matches = pattern.exec(HTMLRangeofSolutionWizard);
                if (matches != null) {
                    linkUrl = matches[1];
                    linkText = matches[2];
                }
                var newLink = document.createElement("a");
                newLink.className = "CIP_SolutionWizard";
                newLink.href = linkUrl;
                newLink.appendChild(document.createTextNode(linkText));
                $.Assets_CIP.InsertObject(currentImg, newLink);
            };

            $.Assets_CIP.Create_One_SolutionWizard_Old = function (currentImg) {
                var HTMLRangeofSolutionWizard = $.Assets_CIP.GetHTMLRangeofCIP(currentImg);
                var head_Bold = "";
                var text_A = "";
                var link_url = "";
                var text_B = "";
                var text_C = "";

                var pattern = new RegExp("<b>([\\w\\W]*?)</b>", "gi");
                var matches = pattern.exec(HTMLRangeofSolutionWizard);
                if (matches != null) {
                    head_Bold = matches[1];
                }

                pattern = new RegExp("<i>([\\w\\W]*?)<u>", "gi");
                matches = pattern.exec(HTMLRangeofSolutionWizard);
                if (matches != null) {
                    text_A = matches[1];
                }

                pattern = new RegExp("<a[^>]* href=\"([^\"]*)[\\w\\W]*\"?>([\\w\\W]*?)</a>", "gi");
                matches = pattern.exec(HTMLRangeofSolutionWizard);
                if (matches != null) {
                    link_url = matches[1];
                    text_B = matches[2];
                }

                pattern = new RegExp("</u>([\\w\\W]*?)</i>", "gi");
                matches = pattern.exec(HTMLRangeofSolutionWizard);
                if (matches != null) {
                    text_C = matches[1];
                }

                //Create <b>Online Solution Wizard</b>
                var b_tag = document.createElement("b");
                var b_text = document.createTextNode(head_Bold);
                b_tag.appendChild(b_text);

                //Create <br/>
                var br_tag = document.createElement("br");

                //Create <a href="http://support.microsoft.com/common/survey.aspx?scid=sw;en;1844&showpage=1"> <img src="http://bemis/Library/Images/2550691.png" border="0"/></a>
                var a_tag = document.createElement("a");
                a_tag.setAttribute("href", link_url);
                var img_tag = document.createElement("img");
                img_tag.setAttribute("src", prefix + "//bemis/Library/Images/2550691.png");
                img_tag.setAttribute("border", "0");
                a_tag.appendChild(img_tag);

                //Create <br/>
                var br_tag2 = document.createElement("br");

                //Create "A"
                var text1 = document.createTextNode(text_A);

                //<a href="http://support.microsoft.com/common/survey.aspx?scid=sw;en;1844&showpage=1">step-by-step solution</a>
                var a_tag2 = document.createElement("a");
                a_tag2.setAttribute("href", link_url);
                var a_text2 = document.createTextNode(text_B);
                a_tag2.appendChild(a_text2);

                //Create "is available to resolve this problem."
                var text2 = document.createTextNode(text_C);

                //Create table
                var table = document.createElement("table");
                table.setAttribute("border", "2");
                table.setAttribute("border", "2");

                table.setAttribute("cellpadding", "3");

                //Createtbody
                var tbody = document.createElement("tbody");
                table.appendChild(tbody);

                //Create first Line
                tbody.insertRow(0);
                tbody.rows[0].insertCell(0);
                tbody.rows[0].cells[0].setAttribute("align", "center");
                tbody.rows[0].cells[0].appendChild(b_tag);
                tbody.rows[0].cells[0].appendChild(br_tag);
                tbody.rows[0].cells[0].appendChild(a_tag);
                tbody.rows[0].cells[0].appendChild(br_tag2);
                tbody.rows[0].cells[0].appendChild(text1);
                tbody.rows[0].cells[0].appendChild(a_tag2);
                tbody.rows[0].cells[0].appendChild(text2);

                //Create <div>
                var div_tag = document.createElement("div");
                div_tag.setAttribute("align", "center");
                div_tag.appendChild(table);

                //Solution_Wizard.appendChild(div_tag);
                //Compatibile with IE7 and IE8
                var outDiv = document.createElement("div");
                outDiv.appendChild(div_tag);
                $.Assets_CIP.InsertObject(currentImg, outDiv);
            };
        })($);
        //************ 5.SolutionWizard End ********************

        //************ 6.ThirdParty Start ********************
        (function ($) {
            $.Assets_CIP.Create_One_ThirdParty = function (currentImg) {
                var link_url = "";
                var link_text = "";
                var link_title = "";
                var HTMLRangeofThirdParty = $.Assets_CIP.GetHTMLRangeofCIP(currentImg);
                var pattern = new RegExp("<a[^>]* href=\"([^\"]*)[\\w\\W]*\"?>([\\w\\W]*?)</a>", "gi");
                var matches = pattern.exec(HTMLRangeofThirdParty);
                if (matches != null) {
                    link_url = matches[1];
                    link_text = matches[2];
                }

                pattern = new RegExp("<i[^>]*>([\\w\\W]*?)</i>", "gi");
                matches = pattern.exec(HTMLRangeofThirdParty);
                if (matches != null) {
                    link_title = matches[1];
                }

                //Create <a class="CIP_Third_Party" href="http://www.google.com/toolbar/ie/install.html" target="_blank" title="Google Toolbar Latest Update"></a>
                var a_tag = document.createElement("a");
                a_tag.className = "CIP_Third_Party";
                a_tag.setAttribute("href", link_url);
                a_tag.setAttribute("target", "_blank");
                a_tag.setAttribute("title", link_title);

                //Create <img alt="Find it on a third-party site" src="http://support.microsoft.com/library/images/support/en-US/thirdparty.png" />, and append it to the link.
                var a_img = document.createElement("img");
                a_img.src = prefix + "//support.microsoft.com/library/images/support/en-US/thirdparty.png";
                a_img.width = "134";
                a_img.height = "45";
                a_img.setAttribute("alt", "Find it on a third-party site");
                a_tag.appendChild(a_img);

                //Create <a href="http://www.google.com/toolbar/ie/install.html" target="_blank" title="Google Toolbar Latest Update">Find it on a third-party site</a>
                var a_tag2 = document.createElement("a");
                a_tag2.setAttribute("href", link_url);
                a_tag2.setAttribute("target", "_blank");
                a_tag2.setAttribute("title", link_title);
                var text = document.createTextNode(link_text);
                a_tag2.appendChild(text);

                //Create <div align="center">
                var div_tag = document.createElement("div");
                div_tag.setAttribute("align", "center");
                div_tag.appendChild(a_tag);
                div_tag.appendChild(a_tag2);

                //Third_Party.appendChild(div_tag);
                //Compatibile with IE7 and IE8
                var outDiv = document.createElement("div");
                outDiv.appendChild(div_tag);
                $.Assets_CIP.InsertObject(currentImg, outDiv);
            };
        })($);
        //************ 6.ThirdParty End ********************

        //************ 7.Video Start ********************
        (function ($) {
            //Generate Video_Player Code
            $.Assets_CIP.Create_One_VideoPlayer = function (currentImg) {
                var culture = GetPlayerLanguage(window.location.href);
                var uuid = "";
                var HTMLRangeofVideoPlayer = $.Assets_CIP.GetHTMLRangeofCIP(currentImg);
                var pattern = new RegExp("<i>([\\w\\W]*?)</i>", "gi");
                var matches = pattern.exec(HTMLRangeofVideoPlayer);
                if (matches != null) {
                    uuid = matches[1].replace(new RegExp(" ", "gi"), "");
                }
                //Create <div>
                var div = document.createElement("div");
                var objectCode = GetObjectCode(culture, uuid);
                div.innerHTML = ReplaceHttp(objectCode);
                $.Assets_CIP.InsertObject(currentImg, div);
            };

            //Get Object（player）Code
            var GetObjectCode = function (culture, uuid) {
                objectStr = "<iframe height=\"360\" width=\"640\" allowfullscreen=\"true\" frameborder=\"0\" marginwidth=\"0\" marginheight=\"0\" scrolling=\"no\" "
                objectStr += "src=\"http://hubs-video.ssl.catalog.video.msn.com/hub/ShowcaseMSN2?csid=ux-cms-en-us-msoffice&iframe=true&uuid=" + uuid + "&PlaybackMode=inline&Quality=HQ&AutoPlayVideo=false&width=640&height=360\"";
                objectStr += "></iframe>";
                return objectStr;

            };
            //Get Player's Language
            var GetPlayerLanguage = function (URL) {
                //Follow SMC URL strucure rule
                var pattern = new RegExp("kb/[0-9]+/(.+)", "gi");
                var matches = pattern.exec(URL);
                if (matches == null) {
                    return "en-US"
                }
                switch (matches[1].toLocaleLowerCase()) {
                    case "ja":
                        return "ja-JP";
                    case "fr":
                        return "fr-FR";
                    case "es":
                        return "es-ES";
                    case "de":
                        return "de-DE";
                    case "zh-cn":
                        return "zh-CN";
                    case "ko":
                        return "ko-KR";
                    case "ru":
                        return "ru-RU";
                    case "it":
                        return "it-IT";
                    case "zh-tw":
                        return "zh-TW";
                    case "pt-br":
                        return "pt-BR";
                    case "nl":
                        return "nl-NL";
                    case "pl":
                        return "pl-PL";
                    case "sv":
                        return "sv-SE";
                    case "fi":
                        return "fi-FI";
                    case "da":
                        return "da-DK";
                    case "no":
                        return "nb-NO";
                    case "cs":
                        return "cs-CZ";
                    case "tr":
                        return "tr-TR";
                    default:
                        return "en-US";
                }
            };
            //Show the matched URL according to the protocle used by the site
            var ReplaceHttp = function (objectCode) {
                var URL = window.location.href;
                var pattern = new RegExp("https", "gi");
                var matches = pattern.exec(URL);
                if (matches != null) {
                    pattern = new RegExp("http", "gi");
                    objectCode = objectCode.replace(pattern, "https")
                }
                return objectCode;
            };
        })($);
        //************ 7.Video End ********************

        //************ 8.Folding Start ********************
        (function ($) {
            $.Assets_CIP.Folding = function () {
                $('img[src*="//support.microsoft.com/library/images/support/en-US/assets_folding_start_"]').each(function () {
                    var startTag;
                    var endTag;
                    switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                        case "ssb":
                        case "smc":
                            startTag = $(this).parent().parent()[0];
                            break;
                        case "bemis": startTag = $(this).parent()[0];
                            break;
                        case "vkb": startTag = $(this)[0];
                            break;
                        default:
                    }

                    $(startTag).parent().contents().filter(function () {
                        return this.nodeType == 3;
                    }).wrap("<span></span>");
                    endTag = GetEndTag(this);
                    if ($(this).attr("src") == "http://support.microsoft.com/library/images/support/en-US/assets_folding_start_expanded.png" || $(this).attr("src") == "https://support.microsoft.com/library/images/support/en-US/assets_folding_start_expanded.png") {
                        $(startTag).nextUntil($(endTag)).wrapAll('<div class="CIP_Section_Expand"></div>');
                    }
                    else if ($(this).attr("src") == "http://support.microsoft.com/library/images/support/en-US/assets_folding_start_collapsed.png" || $(this).attr("src") == "https://support.microsoft.com/library/images/support/en-US/assets_folding_start_collapsed.png") {
                        $(startTag).nextUntil($(endTag)).wrapAll('<div class="CIP_Section_Collapse"></div>');
                    }

                    $(startTag).remove();
                    $(endTag).remove();
                });
                if (!IsSpecialCase()) {
                    SpecialSectionLoad();
                    SectionLoad();
                    LeftAlign();
                }
                else {
                    SpecialSectionLoad();
                    SectionLoad_SpecialCase();
                    RightAlign();
                }

                ExpandAll_CollapseAll();
            };

            var ExpandAll_CollapseAll = function () {
                $('#expandAll').click(function () {
                    $(".CIP_SectionHeadCollapse").click();
                    for (var i = 0; i < SpecialSectionDivArray.length; i += 2) {
                        if ($(SpecialSectionDivArray[i + 1]).css("display") == "none") {
                            $(SpecialSectionDivArray[i]).click();
                        }
                    }
                });
                $('#collapseAll').click(function () {
                    $(".CIP_SectionHeadExpand").click();
                    for (var i = 0; i < SpecialSectionDivArray.length; i += 2) {
                        if ($(SpecialSectionDivArray[i + 1]).css("display") == "block") {
                            $(SpecialSectionDivArray[i]).click();
                        }
                    }
                });
            };
            var IsSpecialCase = function () {
                return $.Assets_CIP.CurrentSiteDir == 'rtl';
            };
            var SpecialSectionLoad = function () {
                ProcessSpace();
                SpecialSectionDivArray = new Array();
                $('img[src*="//support.microsoft.com/library/images/support/en-US/assets_head_folding_start.png"]').each(function () {
                    var startTag;
                    switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                        case "ssb":
                        case "smc":
                            startTag = $(this).parent().parent();
                            InitSpecialSectionDiv(startTag);
                            break;
                        case "bemis": startTag = $(this).parent();
                            InitSpecialSectionDiv(startTag);
                            break;
                        case "vkb": startTag = $(this)[0];
                            InitSpecialSectionDiv(startTag);
                            break;
                        default:
                    }
                });
            };
            var ProcessSpace = function () {
                $('img[src*="//support.microsoft.com/library/images/support/en-US/assets_head_folding_"]').each(function () {
                    var startTag;
                    switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                        case "ssb":
                        case "smc":
                            startTag = $(this).parent().parent();
                            break;
                        case "bemis": startTag = $(this).parent();
                            break;
                        case "vkb": startTag = $(this)[0];
                            break;
                        default:
                    }
                    var reg = new RegExp(" $", "gi");
                    if ($(startTag).prevAll(':first')[0] && $(startTag).prevAll(':first')[0].innerHTML != "") {
                        var tempStr = $(startTag).prevAll(':first')[0].innerHTML;
                        var matches = reg.exec(tempStr);
                        while (matches != null) {
                            tempStr = tempStr.replace(reg, "");
                            matches = reg.exec(tempStr);
                        }
                        $(startTag).prevAll(':first')[0].innerHTML = tempStr;
                    }
                    if ($(startTag).next()[0] && $(startTag).next()[0].innerHTML != "") {
                        tempStr = $(startTag).next()[0].innerHTML;
                        reg = new RegExp("^ ", "gi");
                        matches = reg.exec(tempStr);
                        if (matches != null) {
                            tempStr = tempStr.replace(reg, "");
                            matches = reg.exec(tempStr);
                        }
                        $(startTag).next()[0].innerHTML = tempStr;
                    }

                    if (this.src == "http://support.microsoft.com/library/images/support/en-US/assets_head_folding_start.png" || this.src == "https://support.microsoft.com/library/images/support/en-US/assets_head_folding_start.png") {
                        if ($(startTag).prevAll(':first')[0] && $(startTag).prevAll(':first')[0].innerHTML != "") {
                            var tempStr = $(startTag).prevAll(':first')[0].innerHTML;
                            reg = new RegExp("$", "gi");
                            $(startTag).prevAll(':first')[0].innerHTML = tempStr.replace(reg, "&nbsp;");
                        }
                    }
                    else if (this.src == "http://support.microsoft.com/library/images/support/en-US/assets_head_folding_end.png" || this.src == "https://support.microsoft.com/library/images/support/en-US/assets_head_folding_end.png") {
                        if ($(startTag).next()[0] && $(startTag).next()[0].innerHTML != "") {
                            var tempStr = $(startTag).next()[0].innerHTML;
                            reg = new RegExp("^", "gi");
                            $(startTag).next()[0].innerHTML = tempStr.replace(reg, "&nbsp;");
                        }
                    }
                });
            };
            var InitSpecialSectionDiv = function (ImgDiv_Start) {
                if ($.Assets_CIP.CurrentSiteTypeFlag == "smc" || $.Assets_CIP.CurrentSiteTypeFlag == "bemis") {
                    $(ImgDiv_Start).nextUntil($('div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_head_folding_end.png"])')).wrapAll('<span></span>');
                }
                else if ($.Assets_CIP.CurrentSiteTypeFlag == "vkb") {
                    $(ImgDiv_Start).nextUntil($('img[src*="//support.microsoft.com/library/images/support/en-US/assets_head_folding_end.png"]')).wrapAll('<span></span>');
                }

                var specialSectionDiv = $(ImgDiv_Start).nextAll(".CIP_Section_Collapse:first")[0];
                while (IsSpecialDiv(specialSectionDiv)) {
                    if ($(specialSectionDiv).nextAll(".CIP_Section_Collapse:first")[0]) {
                        specialSectionDiv = $(specialSectionDiv).nextAll(".CIP_Section_Collapse:first")[0];
                    }
                    else {
                        break;
                    }
                }
                SpecialSectionDivArray.push($(ImgDiv_Start).next()[0]);
                SpecialSectionDivArray.push(specialSectionDiv);
                $(ImgDiv_Start).next().css("cursor", "pointer");
                $(ImgDiv_Start).next().toggle(function () {
                    $(specialSectionDiv).css("display", "block");
                }, function () {
                    $(specialSectionDiv).css("display", "none");
                });

                $(ImgDiv_Start).next().next().remove();
                $(ImgDiv_Start).remove();
            };
            var IsSpecialDiv = function (currentDiv) {
                for (var i = 1; i < SpecialSectionDivArray.length; i += 2) {
                    if (SpecialSectionDivArray[i] && SpecialSectionDivArray[i].outerHTML == currentDiv.outerHTML) {
                        return true;
                    }
                }
                return false;
            };
            var RightAlign = function () {
                $('div[class*="CIP_Section_Expand"]').each(function () {
                    $(this).css("padding-right", "25px");
                });
                $('div[class*="CIP_Section_Collapse"]').each(function () {
                    $(this).css("padding-right", "25px");
                });
                $.each(SpecialSectionDivArray, function () {
                    $(this).css("padding-right", "0px");
                });
            };
            var LeftAlign = function () {
                $('div[class*="CIP_Section_Expand"]').each(function () {
                    $(this).css("padding-left", "25px");
                });
                $('div[class*="CIP_Section_Collapse"]').each(function () {
                    $(this).css("padding-left", "25px");
                });
                $.each(SpecialSectionDivArray, function () {
                    $(this).css("padding-left", "0px");
                });
            };
            var SectionLoad = function () {
                var expandSectionList = $(".CIP_Section_Expand");
                var collapseSectionList = $(".CIP_Section_Collapse");
                collapseSectionList.addClass("CIP_Hide");
                var header;
                expandSectionList.each(function () {
                    if (IsSpecialDiv(this)) {
                        return;
                    }
                    header = $(this).prevAll("h3,h4,h5,b").first();
                    if (header.length) {
                        var headerText = $(header).text();
                        $(header).text("");

                        $('<a>', {
                            text: headerText,
                            title: "Click to collapse",
                            href: 'javascript:;'
                        }).appendTo(header);

                        $(header).children().first().prepend(
                            $('<img>', {
                                src: prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_minus.png",
                                width: "20",
                                height: "20",
                                alt: ""
                            })
                        );

                        $(header).addClass("CIP_SectionHead");
                        $(header).addClass("CIP_SectionHeadExpand");
                        $(header).toggle(function () {
                            $(this).removeClass("CIP_SectionHeadExpand");
                            $(this).addClass("CIP_SectionHeadCollapse");
                            $(this).find('a').attr("title", "Click to expand");
                            $(this).find('img').attr("src", prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_plus.png");
                            $(this).nextAll(".CIP_Section_Expand:first").css("display", "none");
                        }, function () {
                            $(this).removeClass("CIP_SectionHeadCollapse");
                            $(this).addClass("CIP_SectionHeadExpand");
                            $(this).find('a').attr("title", "Click to collapse");
                            $(this).find('img').attr("src", prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_minus.png");
                            $(this).nextAll(".CIP_Section_Expand:first").css("display", "block");
                        });
                    }
                });
                collapseSectionList.each(function () {
                    if (IsSpecialDiv(this)) {
                        return;
                    }
                    header = $(this).prevAll("h3,h4,h5,b").first();
                    if (header.length) {
                        var headerText = $(header).text();
                        $(header).text("");

                        $('<a>', {
                            text: headerText,
                            title: "Click to expand",
                            href: 'javascript:;'
                        }).appendTo(header);

                        $(header).children().first().prepend(
                            $('<img>', {
                                src: prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_plus.png",
                                width: "20",
                                height: "20",
                                alt: ""
                            })
                        );

                        $(header).addClass("CIP_SectionHead");
                        $(header).addClass("CIP_SectionHeadCollapse");
                        $(header).toggle(function () {
                            $(this).removeClass("CIP_SectionHeadCollapse");
                            $(this).addClass("CIP_SectionHeadExpand");
                            $(this).find('a').attr("title", "Click to collapse");
                            $(this).find('img').attr("src", prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_minus.png");
                            $(this).nextAll(".CIP_Section_Collapse:first").css("display", "block");
                        }, function () {
                            $(this).removeClass("CIP_SectionHeadExpand");
                            $(this).addClass("CIP_SectionHeadCollapse");
                            $(this).find('a').attr("title", "Click to expand");
                            $(this).find('img').attr("src", prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_plus.png");
                            $(this).nextAll(".CIP_Section_Collapse:first").css("display", "none");
                        });
                    }
                });
            };
            var SectionLoad_SpecialCase = function () {
                var expandSectionList = $(".CIP_Section_Expand");
                var collapseSectionList = $(".CIP_Section_Collapse");
                collapseSectionList.addClass("CIP_Hide");
                var header;
                expandSectionList.each(function () {
                    header = $(this).prevAll("h3,h4,h5,b").first();
                    if (header.length) {
                        var headerText = $(header).text();
                        $(header).text("");

                        $('<a>', {
                            text: headerText,
                            title: "Click to collapse",
                            href: 'javascript:;'
                        }).appendTo(header);

                        $(header).children().first().prepend(
                            $('<img>', {
                                src: prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_minus.png",
                                width: "20",
                                height: "20",
                                alt: ""
                            })
                        );

                        $(header).addClass("CIP_SectionHead");
                        $(header).addClass("CIP_SectionHeadExpand_SpecialCase");
                        $(header).toggle(function () {
                            $(this).removeClass("CIP_SectionHeadExpand_SpecialCase");
                            $(this).addClass("CIP_SectionHeadCollapse_SpecialCase");
                            $(this).find('a').attr("title", "Click to expand");
                            $(this).find('img').attr("src", prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_plus.png");
                            $(this).nextAll(".CIP_Section_Expand:first").css("display", "none");
                        }, function () {
                            $(this).removeClass("CIP_SectionHeadCollapse_SpecialCase");
                            $(this).addClass("CIP_SectionHeadExpand_SpecialCase");
                            $(this).find('a').attr("title", "Click to collapse");
                            $(this).find('img').attr("src", prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_minus.png");
                            $(this).nextAll(".CIP_Section_Expand:first").css("display", "block");
                        });
                    }
                });
                collapseSectionList.each(function () {
                    header = $(this).prevAll("h3,h4,h5,b").first();
                    if (header.length) {
                        var headerText = $(header).text();
                        $(header).text("");

                        $('<a>', {
                            text: headerText,
                            title: "Click to expand",
                            href: 'javascript:;'
                        }).appendTo(header);

                        $(header).children().first().prepend(
                            $('<img>', {
                                src: prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_plus.png",
                                width: "20",
                                height: "20",
                                alt: ""
                            })
                        );

                        $(header).addClass("CIP_SectionHead");
                        $(header).addClass("CIP_SectionHeadCollapse_SpecialCase");
                        $(header).toggle(function () {
                            $(this).removeClass("CIP_SectionHeadCollapse_SpecialCase");
                            $(this).addClass("CIP_SectionHeadExpand_SpecialCase");
                            $(this).find('a').attr("title", "Click to collapse");
                            $(this).find('img').attr("src", prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_minus.png");
                            $(this).nextAll(".CIP_Section_Collapse:first").css("display", "block");
                        }, function () {
                            $(this).removeClass("CIP_SectionHeadExpand_SpecialCase");
                            $(this).addClass("CIP_SectionHeadCollapse_SpecialCase");
                            $(this).find('a').attr("title", "Click to expand");
                            $(this).find('img').attr("src", prefix + "//support.microsoft.com/library/images/support/en-us/20x20_grey_plus.png");
                            $(this).nextAll(".CIP_Section_Collapse:first").css("display", "none");
                        });
                    }
                });
            };
            var GetEndTag = function (StartImg) {
                var startTag;
                var endTag;
                switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                    case "ssb":
                    case "smc":
                        startTag = $(StartImg).parent().parent();
                        endTag = GetEndTag_smc_bemis(startTag);
                        break;
                    case "bemis": startTag = $(StartImg).parent();
                        endTag = GetEndTag_smc_bemis(startTag);
                        break;
                    case "vkb": startTag = $(StartImg)[0];
                        endTag = GetEndTag_vkb(startTag);
                        break;
                    default:
                }
                return endTag;
            };
            var GetEndTag_smc_bemis = function (startTag) {
                var endTag;
                var count = 20;
                var validNext;
                var isMatch = 1;
                while (--count) {
                    validNext = $(startTag).nextUntil($('div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_folding_"])'));
                    if ($(validNext).length != 0) {
                        endTag = $(validNext).filter(':last').next();
                    }
                    else {
                        endTag = $(startTag).next();
                    }

                    if ($(endTag).has('img[src*="//support.microsoft.com/library/images/support/en-US/assets_folding_end"]').length != 0) {
                        if (isMatch > 0) {
                            return endTag;
                        }
                        isMatch += 2;
                        startTag = endTag;
                    }
                    else {
                        isMatch -= 2;
                        startTag = endTag;
                    }
                };
            };
            var GetEndTag_vkb = function (startTag) {
                var endTag;
                var count = 20;
                var validNext;
                var isMatch = 1;
                while (--count) {
                    validNext = $(startTag).nextUntil($('img[src*="//support.microsoft.com/library/images/support/en-US/assets_folding_"]'));
                    if ($(validNext).length != 0) {
                        endTag = $(validNext).filter(':last').next();
                    }
                    else {
                        endTag = $(startTag).next();
                    }

                    var reg = new RegExp("//support.microsoft.com/library/images/support/en-US/assets_folding_end", "gi");
                    if (reg.test($(endTag).attr("src"))) {
                        if (isMatch > 0) {
                            return endTag;
                        }
                        isMatch += 2;
                        startTag = endTag;
                    }
                    else {
                        isMatch -= 2;
                        startTag = endTag;
                    }
                };
            };
        })($);
        //************ 8.Folding End ********************

        //*******************9.GWT Start**************************
        window.getGWT = function (id, config) {
            var language = $.Assets_CIP.GetSiteLanguage();
            switch (language) {
                case "zh-cn":
                case "zh-tw":
                case "pt-pt":
                case "pt-br":
                    break;
                default:
                    language = language.split('-')[0];
                    break;
            }
            _ms_support_fms_surveyConfig = {
                "template": "default",
                "enableLTS": 0,
                "survey": {
                    "language": language,   //survey language such as: en-us, pt-br etc.
                    "id": id, 	//survey id such as: 1959, 2008 etc.
                    "host": 'support.microsoft.com',
                    "features": ["Title,AllNormalPages,Thankyou,NextButton", "Title,Thankyou"]
                },
                "site": {
                    "name": "SMB",	//site name
                    "culture": language,
                    "lcid": '1033',
                    "id": "0",
                    "brand": ""
                },
                "content": {
                    "id": id,
                    "type": "gw",
                    "culture": language,
                    "lcid": '1033',
                    "aggregateId": ""
                },
                "parameters": ['gw', language, id, '', '', '', '', '', '', '']
            };

            if (language.toLowerCase() === "ar" || language.toLowerCase() === "he") {// layout in ar or he should be from right to left
                _ms_support_fms_surveyConfig.survey.isRTL = true;
            }

            window._gwt_config_settings = window._gwt_config_settings || { disabledFeatures: ["TITLE"] };//remove GWT title from KB by default

            if (Object.prototype.toString.call(config) === "[object Boolean]") {//overload, {boolean: enable title}, {array: disabled features}, {object: config}
                if (config == true) {
                    window._gwt_config_settings.enabledFeatures = ["TITLE"];
                }
            } else if (Object.prototype.toString.call(config) === "[object Array]") {
                window._gwt_config_settings.disabledFeatures = config;
            } else if (Object.prototype.toString.call(config) === "[object Object]") {
                window._gwt_config_settings = config;
            }

            var surveystrapperURL = prefix + '//support.microsoft.com/common/script/fx/surveystrapper.js';
            document.write("<script type=\"text/javascript\" src=\"" + surveystrapperURL + "\"></script>");
        };
        //*******************9.GWT End**************************


        //*******************10.Tab Start**************************
        (function ($) {
            $.Assets_CIP.Tab = function () {
                var tabTag = {
                    StartImgOfTab: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_tab_start"]',
                    EndImgOfTab: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_tab_end"]',
                    StartTagOfTab: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_tab_start"])',
                    EndTagOfTab: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_tab_end"])',

                    StartImgOfTabBody: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_tabBody_start"]',
                    EndImgOfTabBody: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_tabBody_end"]',
                    StartTagOfTabBody: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_tabBody_start"])',
                    EndTagOfTabBody: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_tabBody_end"])'
                };
                var tab_Class = {
                    tabClass: "assets_tab",
                    tabBodyClass: "assets_tabBody",
                    currentTextClass: "current_text",
                    contentCollectionClass: "contentCollection"
                };
                var getTag = function (imageObj) {
                    switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                        case "ssb":
                        case "smc":
                            return $(imageObj).parent().parent();
                        case "bemis": return $(imageObj).parent();
                        case "vkb": return $(imageObj);
                    }
                };
                $(tabTag.StartImgOfTab).each(function () {
                    var $startTagOfTab = getTag(this);
                    $startTagOfTab.parent().contents().filter(function () {
                        return this.nodeType == 3;
                    }).wrap("<span></span>");

                    var $assets_tab = $("<div>").addClass(tab_Class.tabClass);
                    var reg = new RegExp("\\?([\\d]+)", "gi");
                    var matches = reg.exec($(this)[0].src);
                    if (matches != null) {
                        $assets_tab.css("width", matches[1] + "px").css("margin", "0 auto");
                    }

                    var $current_text = $("<div>").addClass(tab_Class.currentTextClass).css("padding", "5px");
                    var $contentCollection = $("<div>").addClass(tab_Class.contentCollectionClass).css("display", "none");
                    $assets_tab.append($current_text).append($contentCollection);

                    var isVKBSite = $.Assets_CIP.CurrentSiteTypeFlag == "vkb" ? true : false;
                    var $allTabBody;

                    if (isVKBSite) {
                        $allTabBody = $startTagOfTab.nextUntil($(tabTag.EndImgOfTab));
                    }
                    else {
                        $allTabBody = $startTagOfTab.nextUntil($(tabTag.EndTagOfTab));
                    }

                    $allTabBody.filter(isVKBSite ? $(tabTag.StartImgOfTabBody) : $(tabTag.StartTagOfTabBody)).each(function (key, value) {
                        var $assets_tabBody = $("<div>").addClass(tab_Class.tabBodyClass + key);
                        $(value).nextUntil(isVKBSite ? tabTag.EndImgOfTabBody : tabTag.EndTagOfTabBody).appendTo($assets_tabBody);
                        $contentCollection.append($assets_tabBody);
                    });
                    $assets_tab.insertBefore($startTagOfTab);
                    GenerateNavigateBar($assets_tab);
                });

                function GenerateNavigateBar($assets_tab) {
                    var $this = $assets_tab;
                    var $kb_dynamic_bar = $("<div>").addClass("kb_dynamic_bar");
                    var $current_text = $this.children(".current_text:eq(0)");
                    var $contents = $this.children('.contentCollection');
                    var $dynamicTabs = $("<div>").addClass("dynamicTabs");

                    $contents.children("div").each(function (i, o) {
                        var $o = $(o);
                        var contentClass = $o.attr("class") || "";
                        var contentTitle = $o.children("h3,h4,h5,b").first().text() || $o.attr("class") || "";
                        $o.children("h3,h4,h5,b").first().remove();
                        var tabClass = "dynamicTab";

                        if (i == 0) {
                            tabClass = "dynamicTabActive";
                        }
                        var $tab = $("<div>").addClass(tabClass)
                            .attr("data-target", "." + contentClass);
                        var $tabLink = $("<a>").text(contentTitle)
                            .attr({ "href": "javascript:void(0)" });

                        $tab.append($tabLink);
                        $dynamicTabs.append($tab);
                    });

                    var $tabs = $dynamicTabs.children("div")
                    $tabs.click(function () {
                        var $source = $(this);

                        $tabs.removeClass("dynamicTabActive").addClass("dynamicTab");
                        $source.removeClass("dynamicTab").addClass("dynamicTabActive");

                        var target = $source.attr("data-target");
                        var $currentContent = $contents.find(target).clone(false);
                        $current_text.empty().append($currentContent);
                        if (BrowserDetect.version == 9)//fix IE ul ol bug
                            setTimeout(function () {
                                $current_text.find("ul,ol").hide().show();
                            }, 100);
                    });

                    $kb_dynamic_bar.append($dynamicTabs);
                    $current_text.before($kb_dynamic_bar);

                    $dynamicTabs.children(".dynamicTabActive:eq(0)").click();
                }
            };
        })($);
        //*******************10.Tab End**************************


        //*******************11.FixOther Start**************************
        (function ($) {
            $.Assets_CIP.FixOther = function () {
                var td = $('.Fixit_Div').parent().parent().filter('td');
                $(td).css("border", "0px");
                $(td).parent().parent().parent().filter('table').css("border", "0px").css("border-collapse", "collapse");
            };
        })($);
        //*******************11.FixOther End**************************


        //*******************12.Versions of Windows_IE Start**************************
        (function ($) {
            $.Assets_CIP.VersionsOfWindows_IE = function () {
                var windowsTag = {
                    StartImgOfWindows: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows_start"]',
                    EndImgOfWindows: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows_end"]',
                    StartTagOfWindows: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows_start"])',
                    EndTagOfWindows: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows_end"])',

                    StartImgOfWindows8: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows8_start"]',
                    EndImgOfWindows8: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows8_end"]',
                    StartTagOfWindows8: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows8_start"])',
                    EndTagOfWindows8: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows8_end"])',

                    StartImgOfWindows7: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows7_start"]',
                    EndImgOfWindows7: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows7_end"]',
                    StartTagOfWindows7: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows7_start"])',
                    EndTagOfWindows7: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_windows7_end"])',

                    StartImgOfWindowsVista: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_windowsVista_start"]',
                    EndImgOfWindowsVista: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_windowsVista_end"]',
                    StartTagOfWindowsVista: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_windowsVista_start"])',
                    EndTagOfWindowsVista: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_windowsVista_end"])',

                    StartImgOfWindowsXP: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_windowsXP_start"]',
                    EndImgOfWindowsXP: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_windowsXP_end"]',
                    StartTagOfWindowsXP: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_windowsXP_start"])',
                    EndTagOfWindowsXP: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_windowsXP_end"])'
                };
                var IETag = {
                    StartImgOfIE: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE_start"]',
                    EndImgOfIE: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE_end"]',
                    StartTagOfIE: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE_start"])',
                    EndTagOfIE: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE_end"])',

                    StartImgOfIE10: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE10_start"]',
                    EndImgOfIE10: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE10_end"]',
                    StartTagOfIE10: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE10_start"])',
                    EndTagOfIE10: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE10_end"])',

                    StartImgOfIE9: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE9_start"]',
                    EndImgOfIE9: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE9_end"]',
                    StartTagOfIE9: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE9_start"])',
                    EndTagOfIE9: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE9_end"])',

                    StartImgOfIE8: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE8_start"]',
                    EndImgOfIE8: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE8_end"]',
                    StartTagOfIE8: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE8_start"])',
                    EndTagOfIE8: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE8_end"])',

                    StartImgOfIE7: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE7_start"]',
                    EndImgOfIE7: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE7_end"]',
                    StartTagOfIE7: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE7_start"])',
                    EndTagOfIE7: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE7_end"])',

                    StartImgOfIE6: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE6_start"]',
                    EndImgOfIE6: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE6_end"]',
                    StartTagOfIE6: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE6_start"])',
                    EndTagOfIE6: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_IE6_end"])'
                };
                var bitTag = {
                    ImgOf32Bit: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_32bit"]',
                    TagOf32Bit: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_32bit"])',

                    ImgOf64Bit: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_64bit"]',
                    TagOf64Bit: 'div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_64bit"])'
                };
                var windows_class = {
                    windows8: "windows8",
                    windows7: "windows7",
                    windowsVista: "vista",
                    windowsXP: "winxp"
                };
                var IE_class = {
                    IE10: "IE10",
                    IE9: "IE9",
                    IE8: "IE8",
                    IE7: "IE7",
                    IE6: "IE6"
                };
                var Bit_class = {
                    bit32: "32",
                    bit64: "64"
                };
                var getTag = function (imageObj) {
                    switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                        case "ssb":
                        case "smc":
                            return $(imageObj).parent().parent();
                        case "bemis": return $(imageObj).parent();
                        case "vkb": return $(imageObj);
                    }
                };

                var Win_IE_Config = {
                    non_ie: "non_ie",
                    ie6: "Windows Internet Explorer 6",
                    ie7: "Windows Internet Explorer 7",
                    ie8: "Windows Internet Explorer 8",
                    ie9: "Windows Internet Explorer 9",
                    ie10: "Windows Internet Explorer 10",
                    ieElse: "ieElse",
                    non_win: "non_win",
                    win8: "Windows 8 or Windows Server 2012",
                    win7: "Windows 7 or Windows Server 2008 R2",
                    vista: "Windows Vista or Windows Server 2008",
                    winxp: "Windows XP",
                    winElse: "Windows Else",
                    Mac: "mac"
                }
                var GetCurrentWindows = function () {
                    if ((BrowserDetect.OS == "Windows") && (BrowserDetect.OSVersion == "Windows 8")) {
                        return Win_IE_Config.win8;
                    };
                    if ((BrowserDetect.OS == "Windows") && (BrowserDetect.OSVersion == "Windows 7")) {
                        return Win_IE_Config.win7;
                    };
                    if ((BrowserDetect.OS == "Windows") && (BrowserDetect.OSVersion == "Windows Vista")) {
                        return Win_IE_Config.vista;
                    };
                    if ((BrowserDetect.OS == "Windows") && (BrowserDetect.OSVersion == "Windows XP Professional")) {
                        return Win_IE_Config.winxp;
                    };
                    if ((BrowserDetect.OS == "Windows") && ((BrowserDetect.OSVersion != "Windows XP Professional") && (BrowserDetect.OSVersion != "Windows Vista") && (BrowserDetect.OSVersion != "Windows 7") && (BrowserDetect.OSVersion != "Windows 8"))) {
                        return Win_IE_Config.winElse;
                    };
                    if ((BrowserDetect.OS != "Windows") && (BrowserDetect.OS != "Mac")) {
                        return Win_IE_Config.non_win;
                    };
                    if (BrowserDetect.OS == "Mac") {
                        return Win_IE_Config.Mac;
                    };
                };
                var GetCurrentIE = function () {
                    if (BrowserDetect.browser != "Internet Explorer") {
                        return Win_IE_Config.non_ie;
                    } else if (BrowserDetect.version == 10) {
                        return Win_IE_Config.ie10;
                    } else if (BrowserDetect.version == 9) {
                        return Win_IE_Config.ie9;
                    } else if (BrowserDetect.version == 8) {
                        return Win_IE_Config.ie8;
                    } else if (BrowserDetect.version == 7) {
                        return Win_IE_Config.ie7;
                    } else if (BrowserDetect.version == 6) {
                        return Win_IE_Config.ie6;
                    } else {
                        return Win_IE_Config.ieElse;
                    }
                };
                var GetCurrentOSBit = function () {
                    var isWindowsRT = navigator.userAgent.indexOf("ARM");
                    var OSwowNavigator = navigator.userAgent.indexOf("WOW64");
                    var OSwinNavigator = navigator.userAgent.indexOf("Win64");
                    if (isWindowsRT != -1) {
                        return 'winRT';
                    }
                    else if (OSwowNavigator != -1 || OSwinNavigator != -1) {
                        return 'x64-based';
                    }
                    else {
                        return 'x86-based';
                    }
                };
                var tipInfo = {
                    mac_os: "You are currently using Apple Mac operating system instead of Windows operating system！",
                    non_win: "You are not using Windows operating system!",
                    non_ie: "You are not using Windodws Internet Explorer!"
                };

                var GenerateCollapseBox = function ($startTagOfWindowsOrIE, isWindows) {
                    $startTagOfWindowsOrIE.parent().contents().filter(function () {
                        return this.nodeType == 3;
                    }).wrap("<span></span>");

                    var $collapseBox = $("<div>").addClass("CollapsedBox SupportOptions");
                    var $current_header_text = $("<span>");

                    var $tip = $("<div>").css("color", "red").css("line-height", "30px");
                    var currentWindows = GetCurrentWindows();
                    var currentOSBit = GetCurrentOSBit();
                    var currentIE = GetCurrentIE();
                    var tipText;

                    //Get MT's Content
                    var $startImgOfMT = $('img[src*="//support.microsoft.com/library/images/support/en-US/assets_MT_start"]').first();
                    var $startTagOfMT;
                    var $endTagOfMT;
                    switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                        case "ssb":
                        case "smc":
                            $startTagOfMT = $startImgOfMT.parent().parent();
                            $endTagOfMT = $startTagOfMT.nextAll('div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_MT_end"])').first();
                            break;
                        case "bemis": $startTagOfMT = $startImgOfMT.parent();
                            $endTagOfMT = $startTagOfMT.nextAll('div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_MT_end"])').first();
                            break;
                        case "vkb": $startTagOfMT = $startImgOfMT;
                            $endTagOfMT = $startTagOfMT.nextAll('img[src*="//support.microsoft.com/library/images/support/en-US/assets_MT_end"]').first();
                            break;
                    }
                    var $MTContent = $startTagOfMT.nextUntil($endTagOfMT);
                    $MTContent.css("display", "none");

                    if (isWindows == true) {
                        switch (currentWindows) {
                            case Win_IE_Config.non_win:
                                tipText = tipInfo.non_win;
                                break;
                            case Win_IE_Config.Mac:
                                tipText = tipInfo.mac_os;
                                break;
                            default:
                                if (currentOSBit == "winRT") {
                                    //tipText = "You are currently using the Windows RT!";
                                    if ($MTContent.length > 0) {
                                        tipText = $MTContent.text() + " " + "Windows RT!";
                                    }
                                    else {
                                        tipText = "You are currently using the Windows RT!";
                                    }
                                }
                                else {
                                    //tipText = "You are currently using the " + currentOSBit + " version of " + currentWindows + "!";
                                    if ($MTContent.length > 0) {
                                        tipText = $MTContent.text() + " " + currentOSBit + " version of " + currentWindows + "!"
                                    }
                                    else {
                                        tipText = "You are currently using the " + currentOSBit + " version of " + currentWindows + "!";
                                    }
                                }
                                break;
                        }
                    }
                    else {
                        switch (currentIE) {
                            case Win_IE_Config.non_ie:
                                tipText = tipInfo.non_ie;
                                break;
                            default:
                                //tipText = "You are currently using the "+currentIE+"!";
                                if ($MTContent.length > 0) {
                                    tipText = $MTContent.text() + " " + currentIE + "!";
                                }
                                else {
                                    tipText = "You are currently using the " + currentIE + "!";
                                }
                                break;
                        }
                    }
                    $tip.text(tipText);

                    var $current_text = $("<div>").addClass("current_text CollapsedContent");
                    var isVKBSite = $.Assets_CIP.CurrentSiteTypeFlag == "vkb" ? true : false;
                    if (isVKBSite) {
                        $startTagOfWindowsOrIE.nextUntil($('img[src*="//support.microsoft.com/library/images/support/en-US/assets_"]')).filter('h3,h4,h5,b:first').appendTo($current_header_text);
                    }
                    else {
                        $startTagOfWindowsOrIE.nextUntil($('div:has(img[src*="//support.microsoft.com/library/images/support/en-US/assets_"])')).filter('h3,h4,h5,b:first').appendTo($current_header_text);
                    }
                    //$startTagOfWindowsOrIE.nextAll('h3,h4,h5,b').first().appendTo($current_header_text);
                    $collapseBox.append($current_header_text);

                    //detect language, if is en-us, display the tip
                    if ($.Assets_CIP.CurrentSiteLanguage == "en-us") {
                        $collapseBox.append($tip);
                    }

                    $collapseBox.append($current_text);
                    return $collapseBox;
                };
                var Mark32bitOr64bit = function ($allContents, $container) {
                    $.each(bitTag, function (key, value) {
                        if ($allContents.filter(value).length > 0) {
                            if (key == "ImgOf32Bit" || key == "TagOf32Bit") {
                                $container.addClass(Bit_class.bit32);
                            }
                            else if (key == "TagOf64Bit" || key == "TagOf64Bit") {
                                $container.addClass(Bit_class.bit64);
                            }
                        }
                    });
                };
                var VersionsOfWindows_IE = function () {
                    //Windows
                    $(windowsTag.StartImgOfWindows).each(function () {
                        var $startTagOfWindows = getTag(this);
                        var $collapseBox = GenerateCollapseBox($startTagOfWindows, true);
                        var $windows = $("<div>").addClass("windows CollapsedContent").css("display", "none");
                        var $windows8 = $("<div>").addClass(windows_class.windows8);
                        var $windows7 = $("<div>").addClass(windows_class.windows7);
                        var $windowsVista = $("<div>").addClass(windows_class.windowsVista);
                        var $windowsXP = $("<div>").addClass(windows_class.windowsXP);
                        $windows.append($windows8).append($windows7).append($windowsVista).append($windowsXP);
                        $collapseBox.append($windows);

                        var tempImgFlag = {
                            "StartImgOfWindows8": $windows8,
                            "StartImgOfWindows7": $windows7,
                            "StartImgOfWindowsVista": $windowsVista,
                            "StartImgOfWindowsXP": $windowsXP
                        };
                        var tempTagFlag = {
                            "StartTagOfWindows8": $windows8,
                            "StartTagOfWindows7": $windows7,
                            "StartTagOfWindowsVista": $windowsVista,
                            "StartTagOfWindowsXP": $windowsXP
                        };
                        var $allContents;
                        var $windows8_Content;
                        var $windows7_Content;
                        var $windowsVista_Content;
                        var $windowsXP_Content;
                        switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                            case "smc":
                            case "bemis":
                                $allContents = $startTagOfWindows.nextUntil($(windowsTag.EndTagOfWindows));
                                $windows8_Content = $allContents.filter($(windowsTag.StartTagOfWindows8)).nextUntil($(windowsTag.EndTagOfWindows8));
                                $windows7_Content = $allContents.filter($(windowsTag.StartTagOfWindows7)).nextUntil($(windowsTag.EndTagOfWindows7));
                                $windowsVista_Content = $allContents.filter($(windowsTag.StartTagOfWindowsVista)).nextUntil($(windowsTag.EndTagOfWindowsVista));
                                $windowsXP_Content = $allContents.filter($(windowsTag.StartTagOfWindowsXP)).nextUntil($(windowsTag.EndTagOfWindowsXP));
                                break;
                            case "vkb":
                                $allContents = $startTagOfWindows.nextUntil($(windowsTag.EndImgOfWindows));
                                $windows8_Content = $allContents.filter($(windowsTag.StartImgOfWindows8)).nextUntil($(windowsTag.EndImgOfWindows8));
                                $windows7_Content = $allContents.filter($(windowsTag.StartImgOfWindows7)).nextUntil($(windowsTag.EndImgOfWindows7));
                                $windowsVista_Content = $allContents.filter($(windowsTag.StartImgOfWindowsVista)).nextUntil($(windowsTag.EndImgOfWindowsVista));
                                $windowsXP_Content = $allContents.filter($(windowsTag.StartImgOfWindowsXP)).nextUntil($(windowsTag.EndImgOfWindowsXP));
                                break;
                        }
                        $.each([$windows8_Content, $windows7_Content, $windowsVista_Content, $windowsXP_Content], function (key, value) {
                            var currentKey = key;
                            var $current_Content = value;
                            var isFather = false;
                            var isVKBSite = $.Assets_CIP.CurrentSiteTypeFlag == "vkb" ? true : false;
                            $.each(isVKBSite ? tempImgFlag : tempTagFlag, function (currentFlag, $currentWindows) {
                                if ($current_Content.filter(windowsTag[currentFlag]).length > 0) {
                                    isFather = true;
                                    switch (currentKey) {
                                        case 0:
                                            $currentWindows.addClass(windows_class.windows8);
                                            $windows8.remove();
                                            break;
                                        case 1:
                                            $currentWindows.addClass(windows_class.windows7);
                                            $windows7.remove();
                                            break;
                                        case 2:
                                            $currentWindows.addClass(windows_class.windowsVista);
                                            $windowsVista.remove();
                                            break;
                                        case 3:
                                            $currentWindows.addClass(windows_class.windowsXP);
                                            $windowsXP.remove();
                                            break;
                                    }
                                }
                            });
                            if (!isFather) {
                                switch (currentKey) {
                                    case 0:
                                        $current_Content.appendTo($windows8);
                                        Mark32bitOr64bit($current_Content, $windows8);
                                        break;
                                    case 1:
                                        $current_Content.appendTo($windows7);
                                        Mark32bitOr64bit($current_Content, $windows7);
                                        break;
                                    case 2:
                                        $current_Content.appendTo($windowsVista);
                                        Mark32bitOr64bit($current_Content, $windowsVista);
                                        break;
                                    case 3:
                                        $current_Content.appendTo($windowsXP);
                                        Mark32bitOr64bit($current_Content, $windowsXP);
                                        break;
                                }
                            }
                        });
                        $.each([$windows8, $windows7, $windowsVista, $windowsXP], function () {
                            if ($(this).children().length == 0) {
                                $(this).remove();
                            }
                        });
                        $collapseBox.insertBefore($startTagOfWindows);
                    });

                    //IE
                    $(IETag.StartImgOfIE).each(function () {
                        var $startTagOfIE = getTag(this);
                        var $collapseBox = GenerateCollapseBox($startTagOfIE, false);
                        var $IEs = $("<div>").addClass("IE CollapsedContent").css("display", "none");
                        var $IE10 = $("<div>").addClass(IE_class.IE10);
                        var $IE9 = $("<div>").addClass(IE_class.IE9);
                        var $IE8 = $("<div>").addClass(IE_class.IE8);
                        var $IE7 = $("<div>").addClass(IE_class.IE7);
                        var $IE6 = $("<div>").addClass(IE_class.IE6);
                        $IEs.append($IE10).append($IE9).append($IE8).append($IE7).append($IE6);
                        $collapseBox.append($IEs);
                        var tempImgFlag = {
                            "StartImgOfIE10": $IE10,
                            "StartImgOfIE9": $IE9,
                            "StartImgOfIE8": $IE8,
                            "StartImgOfIE7": $IE7,
                            "StartImgOfIE6": $IE6
                        };
                        var tempTagFlag = {
                            "StartTagOfIE10": $IE10,
                            "StartTagOfIE9": $IE9,
                            "StartTagOfIE8": $IE8,
                            "StartTagOfIE7": $IE7,
                            "StartTagOfIE6": $IE6
                        };
                        var $allContents;
                        var $IE10_Content;
                        var $IE9_Content;
                        var $IE8_Content;
                        var $IE7_Content;
                        var $IE6_Content;
                        switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                            case "smc":
                            case "bemis":
                                $allContents = $startTagOfIE.nextUntil($(IETag.EndTagOfIE));
                                $IE10_Content = $allContents.filter($(IETag.StartTagOfIE10)).nextUntil($(IETag.EndTagOfIE10));
                                $IE9_Content = $allContents.filter($(IETag.StartTagOfIE9)).nextUntil($(IETag.EndTagOfIE9));
                                $IE8_Content = $allContents.filter($(IETag.StartTagOfIE8)).nextUntil($(IETag.EndTagOfIE8));
                                $IE7_Content = $allContents.filter($(IETag.StartTagOfIE7)).nextUntil($(IETag.EndTagOfIE7));
                                $IE6_Content = $allContents.filter($(IETag.StartTagOfIE6)).nextUntil($(IETag.EndTagOfIE6));
                                break;
                            case "vkb":
                                $allContents = $startTagOfIE.nextUntil($(IETag.EndImgOfIE));
                                $IE10_Content = $allContents.filter($(IETag.StartImgOfIE10)).nextUntil($(IETag.EndImgOfIE10));
                                $IE9_Content = $allContents.filter($(IETag.StartImgOfIE9)).nextUntil($(IETag.EndImgOfIE9));
                                $IE8_Content = $allContents.filter($(IETag.StartImgOfIE8)).nextUntil($(IETag.EndImgOfIE8));
                                $IE7_Content = $allContents.filter($(IETag.StartImgOfIE7)).nextUntil($(IETag.EndImgOfIE7));
                                $IE6_Content = $allContents.filter($(IETag.StartImgOfIE6)).nextUntil($(IETag.EndImgOfIE6));
                                break;
                        }
                        $.each([$IE10_Content, $IE9_Content, $IE8_Content, $IE7_Content, $IE6_Content], function (key, value) {
                            var currentKey = key;
                            var $current_Content = value;
                            var isFather = false;
                            var isVKBSite = $.Assets_CIP.CurrentSiteTypeFlag == "vkb" ? true : false;
                            $.each(isVKBSite ? tempImgFlag : tempTagFlag, function (currentFlag, $currentIE) {
                                if ($current_Content.filter(IETag[currentFlag]).length > 0) {
                                    isFather = true;
                                    switch (currentKey) {
                                        case 0:
                                            $currentIE.addClass(IE_class.IE10);
                                            $IE10.remove();
                                            break;
                                        case 1:
                                            $currentIE.addClass(IE_class.IE9);
                                            $IE9.remove();
                                            break;
                                        case 2:
                                            $currentIE.addClass(IE_class.IE8);
                                            $IE8.remove();
                                            break;
                                        case 3:
                                            $currentIE.addClass(IE_class.IE7);
                                            $IE7.remove();
                                            break;
                                        case 4:
                                            $currentIE.addClass(IE_class.IE6);
                                            $IE6.remove();
                                            break;
                                    }
                                }
                            });
                            if (!isFather) {
                                switch (currentKey) {
                                    case 0:
                                        $current_Content.appendTo($IE10);
                                        Mark32bitOr64bit($current_Content, $IE10);
                                        break;
                                    case 1:
                                        $current_Content.appendTo($IE9);
                                        Mark32bitOr64bit($current_Content, $IE9);
                                        break;
                                    case 2:
                                        $current_Content.appendTo($IE8);
                                        Mark32bitOr64bit($current_Content, $IE8);
                                        break;
                                    case 3:
                                        $current_Content.appendTo($IE7);
                                        Mark32bitOr64bit($current_Content, $IE7);
                                        break;
                                    case 4:
                                        $current_Content.appendTo($IE6);
                                        Mark32bitOr64bit($current_Content, $IE6);
                                        break;
                                }
                            }
                        });
                        $.each([$IE10, $IE9, $IE8, $IE7, $IE6], function () {
                            if ($(this).children().length == 0) {
                                $(this).remove();
                            }
                        });
                        $collapseBox.insertBefore($startTagOfIE);
                    });
                };

                VersionsOfWindows_IE();
            };

        })($);
        //*******************12.Versions of Windows_IE End**************************

        //*******************13.OS and browser detection Start**************************
        (function ($, window) {
            window.BrowserDetect = {
                init: function () {
                    this.browser = this.searchString(this.dataBrowser) || "An unknown browser";
                    this.version = this.searchVersion(navigator.userAgent)
                    || this.searchVersion(navigator.appVersion)
                    || "an unknown version";

                    if (this.browser == "Internet Explorer" && this.version == 7 && navigator.userAgent.indexOf('Trident/6.0') != -1)
                        this.version = 10;   /* Compatibility Mode */

                    this.OS = this.searchString(this.dataOS) || "an unknown OS";
                    this.OSVersion = this.searchString(this.dataOSVersion) || "an unknown OSVersion";
                },
                searchString: function (data) {
                    for (var i = 0; i < data.length; i++) {
                        var dataString = data[i].string;
                        var dataProp = data[i].prop;
                        this.versionSearchString = data[i].versionSearch || data[i].identity;
                        if (dataString) {
                            if (dataString.indexOf(data[i].subString) != -1)
                                return data[i].identity;
                        }
                        else if (dataProp)
                            return data[i].identity;
                    }
                },
                searchVersion: function (dataString) {
                    var index = dataString.indexOf(this.versionSearchString);
                    if (index == -1) return;
                    return parseFloat(dataString.substring(index + this.versionSearchString.length + 1));
                },
                dataBrowser: [
                {
                    string: navigator.userAgent,
                    subString: "Chrome",
                    identity: "Chrome"
                },
                {
                    string: navigator.userAgent,
                    subString: "OmniWeb",
                    versionSearch: "OmniWeb/",
                    identity: "OmniWeb"
                },
                {
                    string: navigator.vendor,
                    subString: "Apple",
                    identity: "Safari",
                    versionSearch: "Version"
                },
                {
                    prop: window.opera,
                    identity: "Opera"
                },
                {
                    string: navigator.vendor,
                    subString: "iCab",
                    identity: "iCab"
                },
                {
                    string: navigator.vendor,
                    subString: "KDE",
                    identity: "Konqueror"
                },
                {
                    string: navigator.userAgent,
                    subString: "Firefox",
                    identity: "Firefox"
                },
                {
                    string: navigator.vendor,
                    subString: "Camino",
                    identity: "Camino"
                },
                {		/* for newer Netscapes (6+)*/
                    string: navigator.userAgent,
                    subString: "Netscape",
                    identity: "Netscape"
                },
                {
                    string: navigator.userAgent,
                    subString: "MSIE",
                    identity: "Internet Explorer",
                    versionSearch: "MSIE"
                },
                {
                    string: navigator.userAgent,
                    subString: "Gecko",
                    identity: "Mozilla",
                    versionSearch: "rv"
                },
                { 	/* for older Netscapes (4-)*/
                    string: navigator.userAgent,
                    subString: "Mozilla",
                    identity: "Netscape",
                    versionSearch: "Mozilla"
                }
                ],
                dataOS: [
                {
                    string: navigator.platform,
                    subString: "Win",
                    identity: "Windows"
                },
                {
                    string: navigator.platform,
                    subString: "Mac",
                    identity: "Mac"
                },
                {
                    string: navigator.userAgent,
                    subString: "iPhone",
                    identity: "iPhone/iPod"
                },
                {
                    string: navigator.platform,
                    subString: "Linux",
                    identity: "Linux"
                }
                ],
                dataOSVersion: [
                {
                    string: navigator.userAgent,
                    subString: "Windows 95",
                    identity: "Windows 95 OSR2"
                },
                {
                    string: navigator.userAgent,
                    subString: "Windows 98",
                    identity: "Windows 98"
                },
                {
                    string: navigator.userAgent,
                    subString: "Windows NT 5.0",
                    identity: "Windows 2000 Professional"
                },
                {
                    string: navigator.userAgent,
                    subString: "Windows NT 5.1",
                    identity: "Windows XP Professional"
                },
                {
                    string: navigator.userAgent,
                    subString: "Windows NT 5.2",
                    identity: "Windows 2003 Server"
                },
                {
                    string: navigator.userAgent,
                    subString: "Windows NT 6.0",
                    identity: "Windows Vista"
                },
                {
                    string: navigator.userAgent,
                    subString: "Windows NT 6.1",
                    identity: "Windows 7"
                },
                {
                    string: navigator.userAgent,
                    subString: "Windows NT 6.2",
                    identity: "Windows 8"
                },
                {
                    string: navigator.userAgent,
                    subString: "Mac_PowerPC",
                    identity: "Mac OS 9.2"
                }
                ]
            };
        })($, window);
        //*******************13.OS and browser detection End**************************


        //*******************14.KB dynamic tabs start**************************
        (function (window) {
            var versionConfig = {
                non_ie: "non_ie",
                ie6: "ie6",
                ie7: "ie7",
                ie8: "ie8",
                ie9: "ie9",
                ie10: "ie10",
                ieElse: "ieElse",
                non_win: "non_win",
                win8: "windows8",
                win7: "windows7",
                vista: "vista",
                winxp: "winxp",
                winElse: "winElse"
            }
            BrowserDetect.init();

            var osVer = GetOSVer();
            var browserVer = GetBrowserVer();

            window.KB_Dynamic_Section_Init = function () {
                var $dynamicSections = $(".CollapsedBox");
                $dynamicSections.each(function () {
                    var $this = $(this);

                    var isWinSec = $this.children(".windows").length > 0 ? true : false;
                    var conSelector = isWinSec ? ".windows:eq(0)" : ".IE:eq(0)";
                    var version = isWinSec ? osVer : browserVer;

                    var $kb_dynamic_bar = $("<div>").addClass("kb_dynamic_bar");
                    var $current_text = $this.children(".current_text:eq(0)");
                    var $contents = $this.children(conSelector);

                    var $dynamicTabs = $("<div>").addClass("dynamicTabs");

                    var hasCurrentContent = false;//whether contain content for current version

                    $contents.children("div").each(function (i, o) {
                        var $o = $(o);
                        var contentClass = $o.attr("class") || "";

                        var $tabTitle = $(o).children("h3,h4,h5,b").first();
                        var contentTitle = $tabTitle.text() || "";
                        if (contentTitle == "") {
                            return;
                        }
                        $tabTitle.remove();
                        var tabClass = "dynamicTab";
                        if (contentClass.toLocaleLowerCase().indexOf(version) >= 0) {
                            tabClass = "dynamicTabActive";
                            hasCurrentContent = true;
                        }

                        var tSelector = GenerateClassSelector(contentClass);
                        var $tab = $("<div>").addClass(tabClass)
                            .attr("data-target", tSelector);
                        var $tabLink = $("<a>").text(contentTitle)
                            .attr({ "href": "javascript:void(0)" });

                        $tab.append($tabLink);
                        $dynamicTabs.append($tab);
                    });

                    var $tabs = $dynamicTabs.children("div")

                    if (hasCurrentContent == false) {//if section dosen't contain content for current version, choose fist tab.
                        $tabs.first().removeClass("dynamicTab").addClass("dynamicTabActive");
                    }

                    $tabs.click(function () {
                        var $source = $(this);

                        $tabs.removeClass("dynamicTabActive").addClass("dynamicTab");
                        $source.removeClass("dynamicTab").addClass("dynamicTabActive");

                        var target = $source.attr("data-target");
                        var $currentContent = $contents.find(target).clone(true, true);
                        $current_text.empty().append($currentContent);
                        if (BrowserDetect.version == 9)//fix IE9 ul ol bug
                            setTimeout(function () {
                                $current_text.find("ul,ol").hide().show();
                            }, 100);
                    });

                    $kb_dynamic_bar.append($dynamicTabs);
                    $current_text.before($kb_dynamic_bar);

                    //show content for current version
                    $dynamicTabs.children(".dynamicTabActive:eq(0)").click();
                });
            }

            function GetBrowserVer() {
                var version = "";
                if (BrowserDetect.browser != "Internet Explorer") {
                    version = versionConfig.non_ie;
                } else if (BrowserDetect.version == 10) {
                    version = versionConfig.ie10;
                } else if (BrowserDetect.version == 9) {
                    version = versionConfig.ie9;
                } else if (BrowserDetect.version == 8) {
                    version = versionConfig.ie8;
                } else if (BrowserDetect.version == 7) {
                    version = versionConfig.ie7;
                } else if (BrowserDetect.version == 6) {
                    version = versionConfig.ie6;
                } else {
                    version = versionConfig.ieElse;
                }
                return version;
            }

            function GetOSVer() {
                var version = "";
                if (BrowserDetect.OS != "Windows") {
                    version = versionConfig.non_win;
                } else if (BrowserDetect.OSVersion == "Windows 8") {
                    version = versionConfig.win8;
                } else if (BrowserDetect.OSVersion == "Windows 7") {
                    version = versionConfig.win7;
                } else if (BrowserDetect.OSVersion == "Windows Vista") {
                    version = versionConfig.vista;
                } else if (BrowserDetect.OSVersion == "Windows XP Professional") {
                    version = versionConfig.winxp;
                } else {
                    version = versionConfig.winElse;
                }
                return version;
            }

            function GenerateClassSelector(className) {
                var result = "";
                var names = className.split(" ");
                for (var i = 0; i < names.length; i++) {
                    if (className != " ")
                        result += "." + names[i];
                }
                result += ":eq(0)";
                return result;
            }
        })(window);
        //*******************14.KB dynamic tabs  End**************************

        //*******************14.detect os bit start**************************
        (function (window) {
            var images = {
                Bit64Start: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_64s.png"]:eq(0)',
                Bit64StartSrc: "//support.microsoft.com/library/images/support/en-US/assets_64s.png",
                Bit64End: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_64e.png"]:eq(0)',
                Bit64EndSrc: "//support.microsoft.com/library/images/support/en-US/assets_64e.png",
                Bit32Start: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_32s.png"]:eq(0)',
                Bit32StartSrc: "//support.microsoft.com/library/images/support/en-US/assets_32s.png",
                Bit32End: 'img[src*="//support.microsoft.com/library/images/support/en-US/assets_32e.png"]:eq(0)',
                Bit32EndSrc: "//support.microsoft.com/library/images/support/en-US/assets_32e.png"
            }

            function GetCurrentOSBit() {
                var agentStr = navigator.userAgent;
                var wowIndex = agentStr.indexOf("WOW64");
                var winIndex = agentStr.indexOf("Win64");
                if (wowIndex >= 0 || winIndex >= 0) {
                    return "64 bit";
                }
                else {
                    return "32 bit";
                }
                return "other os";
            };

            function GetSiteType() {
                var siteTypeFlag = "smc";
                var locationURL = window.location.href;
                var reg = new RegExp("http://bemis", "gi");
                if (reg.test(locationURL)) {
                    siteTypeFlag = "bemis";
                }
                else {
                    reg = new RegExp("(http://vkb)|(https://vkb)|(http://visualkb)|(https://visualkb)", "gi");
                    if (reg.test(locationURL)) {
                        siteTypeFlag = "vkb"
                    }
                    else {
                        reg = new RegExp("microsoft.com|(157.56.56.32)", "gi");
                        if (reg.test(locationURL)) {
                            siteTypeFlag = "smc";
                        }
                    }
                }
                return siteTypeFlag;
            }

            var osBit = GetCurrentOSBit();
            var site = GetSiteType();

            $(function () {
                var $startImg64 = $(images.Bit64Start);
                var $startImg32 = $(images.Bit32Start);
                var wrapper = null;
                var endImgSrc = "";
                switch (osBit) {
                    case "64 bit":
                        if ($startImg64.length > 0) {
                            wrapper = GetWrapper($startImg64[0], site);
                            endImgSrc = images.Bit64EndSrc;
                        }
                        break;
                    case "32 bit":
                        if ($startImg32.length > 0) {
                            wrapper = GetWrapper($startImg32[0], site);
                            endImgSrc = images.Bit32EndSrc;
                        }
                        break;
                }

                if (wrapper != null) {//create tip box
                    var middle = GetMiddle(wrapper, endImgSrc);
                    var closeNode = middle.closeNode;
                    var middleArr = middle.middleArr;

                    if (closeNode) {
                        var tipText = "";
                        for (var i = 0; i < middleArr.length; i++) {
                            tipText += $(middleArr[i]).text();
                        }
                        $(closeNode).after(CreateBitTip(tipText));
                    }
                }

                if ($startImg32.length > 0) {//delete 32 bit predefine images and text
                    var wrapper32 = GetWrapper($startImg32[0], site);
                    var middle = GetMiddle(wrapper32, images.Bit32EndSrc);
                    var closeNode = middle.closeNode;
                    var middleArr = middle.middleArr;
                    if (closeNode) {
                        for (var i = 0; i < middleArr.length; i++) {
                            $(middleArr[i]).remove();
                        }
                    }
                    $(wrapper32).remove();
                    $(closeNode).remove();
                }

                if ($startImg64.length > 0) {//delete 64 bit predefine iamges and text
                    var wrapper64 = GetWrapper($startImg64[0], site);
                    var middle = GetMiddle(wrapper64, images.Bit64EndSrc);
                    var closeNode = middle.closeNode;
                    var middleArr = middle.middleArr;
                    if (closeNode) {
                        for (var i = 0; i < middleArr.length; i++) {
                            $(middleArr[i]).remove();
                        }
                    }
                    $(wrapper64).remove();
                    $(closeNode).remove();
                }
            })

            function CreateBitTip(content) {
                var $container = $("<div>").addClass("cip-bit-box").css({
                    "border": "1px solid rgb(147, 147, 147)",
                    "border-radius": "4px",
                    "padding": "10px",
                    "font-weight": "bold",
                    "text-align": "center",
                    "color": "blue",
                    "box-shadow": "2px"
                });
                $container.text(content);
                return $container;
            }

            function GetMiddle(startNode, endImg) {
                var result = {
                    middleArr: [],
                    closeNode: null
                };
                var next = startNode.nextSibling;
                while (next) {
                    if (HasImg(next, endImg)) {
                        result.closeNode = next;
                        break;
                    } else {
                        result.middleArr.push(next);
                    }
                    next = next.nextSibling;
                }
                return result;
            }

            function GetWrapper(image, site) {
                switch (site) {
                    case "ssb":
                    case "smc":
                        return image.parentNode.parentNode;
                    case "bemis":
                        return image.parentNode;
                    case "vkb":
                        return image;
                }
            }

            function HasImg(wrapper, src) {
                if (wrapper.nodeType != 1) {
                    return false;
                }
                if (wrapper.tagName && wrapper.tagName.toLowerCase() == "img") {
                    var imgSrc = wrapper.getAttribute("src");
                    return imgSrc.indexOf(src) > 0 ? true : false;
                } else {
                    return $(wrapper).find('img[src*="' + src + '"]').length > 0 ? true : false;
                }
            }
        })(window);
        //*******************14.detect os bit end**************************

        //************ 15.Layout Image Title Abstract 1 Start ********************
        (function ($) {
            $.Assets_CIP.LayoutImageTitleAbstract1 = function () {
                var i = 0;

                $("a[href='#cfstartimghdabs']").each(function () {
                    var startTag = $.MSCOSCodeFree.GetBookmarkNode(this, 0);

                    var endTag = $.MSCOSCodeFree.GetBookmarkNode($("a[href='#cfendimghdabs']"), i);

                    if (startTag != undefined && endTag != undefined) {
                        $(startTag).parent().contents().filter(function () {
                            return this.nodeType == 3;
                        }).wrap("<span></span>");

                        var imghdabsLayout1 = $(startTag).nextUntil(endTag);
                        if (imghdabsLayout1 != undefined && $(imghdabsLayout1).length != 0) {
                            imghdabsLayout1.wrapAll('<div class="cfimghdabs1"></div>');
                        }
                        i++;
                    }
                });
                $('div.cfimghdabs1').each(function () {
                    var img = $(this).find('img');
                    var ul = $(this).find('ul');
                    if (img != undefined && img.length > 0) {
                        var dir = 'left';
                        if ($.Assets_CIP.CurrentSiteDir != 'ltr') {
                            dir = 'right';
                        }
                        $(img).attr("align", dir);
                        var ni = new Image();
                        ni.onload = function () {
                            var imgHeight = ni.height;
                            var imgWidth = ni.width + 35;
                            $(img).parents('div.cfimghdabs1').css('min-height', imgHeight);

                            $(ul).css('padding-' + dir, imgWidth);
                            $(ul).css('margin-' + dir, 0);
                        }
                        ni.src = img[0].src;
                    }
                });

            };
        })($);
        //************ 15.Layout Image Title Abstract 1 End ********************

        //************ 15.1.Layout Image Title Abstract 2 Start ********************
        (function ($) {
            $.Assets_CIP.LayoutImageTitleAbstract2 = function () {
                var i = 0;

                $("a[href='#cfstartimghdabs2']").each(function () {
                    var startTag = $.MSCOSCodeFree.GetBookmarkNode(this, 0);

                    var endTag = $.MSCOSCodeFree.GetBookmarkNode($("a[href='#cfendimghdabs2']"), i);

                    if (startTag != undefined && endTag != undefined) {
                        $(startTag).parent().contents().filter(function () {
                            return this.nodeType == 3;
                        }).wrap("<span></span>");

                        var imghdabsLayout1 = $(startTag).nextUntil(endTag);
                        if (imghdabsLayout1 != undefined && $(imghdabsLayout1).length != 0) {
                            imghdabsLayout1.wrapAll('<div class="cfimghdabs2 cfLTR"></div>');
                        }
                        i++;
                    }
                });
                $('div.cfimghdabs2').each(function () {
                    var layoutWidth = $(this).width();
                    var img = $(this).find('img');

                    if (img != undefined && img.length > 0) {
                        var ni = new Image();
                        ni.onload = function () {
                            var imgWidth = ni.width;
                            var imgNode = $.MSCOSCodeFree.GetImageNode(img);

                            if (imgNode != undefined) {
                                $(imgNode).width(imgWidth);
                                $(imgNode).addClass("cfimg");

                                var cfimghdabs1Width = layoutWidth - imgWidth - 10;
                                $(imgNode).nextAll().wrapAll('<div class="cftitleabs" style="width:' + cfimghdabs1Width + 'px"></div>');
                            }
                        }
                        ni.src = img[0].src;
                    }
                });

            };
        })($);
        //************ 15.1.Layout Image Title Abstract 2 End ********************        

        $.MSCOSCodeFree = {
            GetDir: function () {
                var SpecialLanguageArray = new Array("ar", "he", "ar-sa", "he-il");
                var URL = window.location.href;
                var pattern = new RegExp("kb/[0-9]+/(.+)", "gi");
                var matches = pattern.exec(URL);
                var language;
                if (matches != null) {
                    language = matches[1].toLocaleLowerCase();
                    for (var i = 0; i < SpecialLanguageArray.length; i++) {
                        pattern = new RegExp(SpecialLanguageArray[i], "gi");
                        matches = pattern.exec(language);
                        if (matches != null) {
                            return 'rtl';
                        }
                    }
                }
                return 'ltr';
            },
            GetBookmarkNode: function (node, i) {
                if (node == undefined) {
                    return undefined;
                }
                var returnNode;
                switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                    case "smc":
                    case "ssb":
                        returnNode = $(node).parent()[i];
                        break;
                    case "bemis": returnNode = $(node)[i];
                        break;
                    default:
                        returnNode = $(node).parent()[i];
                }
                return returnNode;
            },
            GetImageNode: function (node) {
                if (node == undefined) {
                    return undefined;
                }
                var returnNode;
                switch ($.Assets_CIP.CurrentSiteTypeFlag) {
                    case "smc":
                        if ($(node).parent()[0].children.length > 1) {
                            returnNode = $(node).parent()[0];
                        }
                        else {
                            returnNode = $(node).parent().parent()[0];
                        }
                        break;
                    case "bemis":
                    case "ssb":
                        returnNode = $(node).parent()[0];
                        break;
                    default:
                        returnNode = $(node).parent().parent()[0];
                }
                return returnNode;
            }
        }

        $(function () {
            $.Assets_CIP();
            if ($.Assets_CIP.CurrentSiteTypeFlag == 'smc' || $.Assets_CIP.CurrentSiteTypeFlag == 'ssb' || $.Assets_CIP.CurrentSiteTypeFlag == 'bemis') {
                $.Assets_CIP.LayoutImageTitleAbstract1();
                $.Assets_CIP.LayoutImageTitleAbstract2();
            }

            if ($("a[href='#cfstartslideshow']").length > 0 && $.Assets_CIP.CurrentSiteTypeFlag == 'smc') {
                downloadJS(prefix + '//support.microsoft.com/library/JavaScript/support/en-US/Assets_CIP_Slideshow.js');
            }

            $.Assets_CIP.FixOther();

            BrowserDetect.init();
            $.Assets_CIP.VersionsOfWindows_IE();
            KB_Dynamic_Section_Init();

            $.Assets_CIP.Tab();
            $.Assets_CIP.DeleteAllAssetsImages();
        });

    });
})();
