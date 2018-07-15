var Microsoft = Microsoft || {}; Microsoft.Support = Microsoft.Support || {}; Microsoft.Support.Ssb = Microsoft.Support.Ssb || {};

//(function ($) {
    Microsoft.Support.Ssb.DisplaySsbAd = function () {
        var muid = fetchcookieval("MUID");
        if (!muid) {
            return;
        }
        $(document).ready(
        function () {
            isSmallMediumBusiness(muid, function () { $("#divSsbAdContent").show(); });
        });
    }

    Microsoft.Support.Ssb.DisplaySsbBanner = function () {
        var muid = fetchcookieval("MUID");
        if (!muid) {
            return;
        }
        $(document).ready(
        function () {
            isSmallMediumBusiness(muid, function () { $("#divSsbBannerContent").show(); });
        });
    }

    var isSmallMediumBusiness = function (muid, callback) {

        var url = PDF.GetConfig("SsbSmcServiceUrl")
        $.get(url, null,
                function (data) {
                    if (containValue(data, ",", "IsSmallMediumBusiness")) {
                        callback();
                    }
                });
    }

    var containValue = function (longString, separator, value) {
        var list = longString.split(separator), result = false, i;
        for (i = 0; i < list.length; i += 1) {
            if (list[i] === value) {
                result = true;
                break;
            }
        }
        return result;
    };
//}(jQuery));