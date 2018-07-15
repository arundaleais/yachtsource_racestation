(function (window, $, ko, PDF, undefined) {
    "use strict";

    PDF.GetSupport = {
        PesId: "",
        ProductFamilyId: "",
        Locale: "",
        Lcid: "",
        Host: "",
        Context: [],
        PartnerId: "",
        Persist: false,
        SupportTopicId: "",
        add: function (key, value, addToQueryString) {
            var exist = false;
            if (key) {
                for (var i = 0; i < PDF.GetSupport.Context.length; i++) {

                    if (PDF.GetSupport.Context[i].key.toLowerCase() == key.toLowerCase()) {

                        PDF.GetSupport.Context[i].value = value;
                        PDF.GetSupport.Context[i].addToQueryString = addToQueryString || false;
                        break;
                    }

                }
                if (!exist) {

                    PDF.GetSupport.Context.push({ key: key, value: value, addToQueryString: addToQueryString || false });
                }
            }
        },
        get: function (key) {
            for (var i = 0; i < PDF.GetSupport.Context.length; i++) {

                if (PDF.GetSupport.Context[i].key.toLowerCase() == key.toLowerCase()) {

                    return PDF.GetSupport.Context[i].value;
                }

                return "";

            }
        },
        load: function (callback) {
            var needhelp = $('.needhelp');
            if (needhelp.attr('href') === 'javascript:void(0)') {
                $(".needhelp").on('click', function() {
                    PDF.GetSupport.requestHelp(callback);
                });
            }
        },
        requestHelp: function (callback) {
            if (typeof callback === 'function') {
                callback();
            }

            var getUrl = PDF.Utils.GetContextUrlFn("GetSupport", "dataContextBaseUrl");

            PDF.Utils.AjaxWithCredential({

                url: getUrl('api/pdf/GetSupport/SetContext'),
                type: 'POST',
                cache: false,
                data: JSON.stringify(PDF.GetSupport),
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                success: function (data) {
                    if (data) {
                        PDF.Events.Trigger("getsupport.success", data);

                        if (data.result === true) {
                            PDF.GetSupport.redirect(data.location);

                        }
                        else {
                            PDF.Utils.LogMessage("redirection was not successful result was false", PDF.LogLevels.ERROR, document.URL, 0);
                            PDF.Events.Trigger("getsupport.error", { componentName: "GetSupport", arguments: "redirection was not successful result was false" });
                        }
                    }
                },
                error: function (data) {
                    PDF.Utils.LogMessage(data.responseText, PDF.LogLevels.ERROR, document.URL, 0);
                    PDF.Events.Trigger("getsupport.error", { componentName: "GetSupport", arguments: data });

                },
                complete: function () {
                    PDF.Events.Trigger("getsupport.complete");
                }
            });

        },
        attach: function (element, eventType, callback) {
            if (element && $(element).length > 0) {
                if (eventType) {
                    $(element).bind(eventType, function () {
                        PDF.GetSupport.requestHelp(callback);
                    });
                } else {
                    $(element).click(function () {
                        PDF.GetSupport.requestHelp(callback);
                    });
                }
            }
        },
        redirect: function (url) {
            if (url) {
                window.location.assign(url);
            }
        }
    };

})(window, jQuery, ko, PDF);

$(document).ready(PDF.GetSupport.load());

