    /* Browser Detect Script */

    var BrowserDetect = {

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
		{ string: navigator.userAgent,
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