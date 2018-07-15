/// <reference path="jquery-1.7.1.js" />
/// <reference path="os_browser_detect_script.js" />

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

            var hasCurrentContent=false;//whether contain content for current version

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
