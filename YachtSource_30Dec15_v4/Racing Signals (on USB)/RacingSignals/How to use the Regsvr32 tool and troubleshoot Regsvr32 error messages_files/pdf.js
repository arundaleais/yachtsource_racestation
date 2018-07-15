///#source 1 1 /../../Components/PDF.Component.DynamicGPS/Scripts/pdf.comp.dgps-1.0.0.js
///#source 1 1 /../../Components/PDF.Component.DynamicGPS/Scripts/pdf.comp.dgps-1.0.0.js
(function (window, $, ko, PDF, undefined)
{
    "use strict";
    var componentIdentifier = "DynamicGPS";

    PDF.ViewModels[componentIdentifier] = function (promise, el, params)
    {
        var viewModel = function ($)
        {
            var self = Object.create(PDF.BaseViewModel);
            var cookieSeperator = '$%^';
            var parentSeperator = '^%$';
            var cookieName = 'dgpsHistory';
            var productSuggestSourceType = "ProductAutosuggestion";
            var taxonomySuggestSourceType = "TaxonomyAutosuggestion";
            var topSearchSuggestSourceType = "TopSearch";
            var searchHistorySuggestSourceType = "SearchHistory";
            var config = PDF.GetComponentConfig(componentIdentifier);
            var suggestionCount = config.suggestionCount || 8;
            var recentSearchesCount = config.recentSearchesCount || 5;
            var autoSuggestUrl = config.autoSuggestUrl || 'http://dgps.cloudapp.net/API/v1/AutoSuggestion';
            var searchPageUrl = params.searchPageUrl || '';
            self.currentSuggestions = {};
            self.searchImageUrl = ko.observable();
            self.searchImageUrl(params.searchImageUrl || PDF.Config.imageUrl + 'icon_search_white.png');
            self.searchTerm = ko.observable();
            $.support.cors = true;
            self.autoSuggestions = [];
            self.currentAutosuggestion = {};
            self.preAutoFillTerm = '';
            self.AppendCookie = function (cookieName, value)
            {
                if (value)
                {
                    var cookieValue = PDF.Utils.GetCookie(cookieName);
                    if (cookieValue)
                    {
                        // Last search should appear on top so append it in begenning
                        cookieValue = self.htmlEncode(value) + cookieSeperator + cookieValue;
                    }
                    else
                    {
                        cookieValue = self.htmlEncode(value);
                    }
                    PDF.Utils.SetCookie(cookieName, cookieValue);
                }
            };

            self.GetSearchHistory = function (name)
            {
                var cookieValue = PDF.Utils.GetCookie(name);
                if (cookieValue)
                {
                    return self.getUniqueArray(cookieValue.split(cookieSeperator));
                }
                else
                {
                    return null;
                }
            };

            PDF.Events.Bind("ComponentLoaded", function (e, data)
            {
                if (data && data.componentName === componentIdentifier)
                {
                    if (params.hasOwnProperty("searchTerm") && params.searchTerm)
                    {
                        $("#dgpsSearch").val(params.searchTerm);
                    }

                    if (params.hasOwnProperty("searchButtonClass") && params.searchButtonClass)
                    {
                        $(".dgpsSearchBtn").addClass(params.searchButtonClass);
                    }

                    $("#dgpsSearch").attr('maxlength', config.searchMaxLength || 250);

                    if (params.hasOwnProperty("searchBoxWidth") && params.searchBoxWidth)
                    {
                        $("#dgpsSearch").css('width', params.searchBoxWidth);
                    }

                    // If search box gets focus and is empty trigger autocomplete to execute for top searches
                    $("#dgpsSearch").focus(function ()
                    {
                        if ($('#dgpsSearch').val() === "")
                        {
                            $('#dgpsSearch').autocomplete("search");
                        }
                    });

                    $("#dgpsSearch").click(function ()
                    {
                        $('#dgpsSearch').autocomplete("search");
                    });

                    $(".dgpsSearchBtn").click(function ()
                    {
                        self.triggerSearch();
                    });

                    // Override render item to highlight terms based on need
                    $.ui.autocomplete.prototype._renderItem = function (ul, item)
                    {
                        // Escape any regex syntax inside this.term
                        var cleanTerm = this.term.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');

                        var output = '';
                        var returnVal = '';

                        var prefixText = cleanTerm;
                        var beginPrefixIndex = item.label.toUpperCase().indexOf(prefixText.toUpperCase());
                        if (beginPrefixIndex === 0)
                        {
                            var endPrefixIndex = item.label.length;

                            if (beginPrefixIndex !== -1 && prefixText != "")
                            {
                                endPrefixIndex = beginPrefixIndex + (prefixText.length);
                                returnVal += item.label.slice(0, endPrefixIndex);

                                if (endPrefixIndex < item.label.length)
                                {
                                    returnVal += "<span class=\"ui-menu-item-highlight\">{0}</span>".format(item.label.slice(endPrefixIndex));
                                }
                            }
                            else
                            {
                                returnVal = self.htmlEncode(item.label);
                            }
                            output = returnVal.toLowerCase();
                        }
                        else
                        {
                            var queryWords = prefixText.split(' ');
                            returnVal = item.label.toUpperCase();

                            queryWords = self.getUniqueArray(queryWords);

                            for (var i = 0; i < queryWords.length; i++)
                            {
                                returnVal = returnVal.replace(queryWords[i].toUpperCase(), "<span class=\"ui-menu-item-highlight\">{0}</span>".format(queryWords[i].toUpperCase()));
                            }

                            output = returnVal.toLowerCase();
                        }
                        var className = "";

                        if (item.hasOwnProperty('recentSearch') && item.recentSearch)
                        {
                            className = "recentSearch";
                        }

                        return $("<li>")
                           .append($('<a class=\"font3 ' + className + '\">').html(output))
                           .appendTo(ul);
                    };

                    $("#dgpsSearch").autocomplete({
                        source: function (request, response)
                        {
                            if (window.XDomainRequest)
                            {
                                var xdr = new XDomainRequest();
                                var url = autoSuggestUrl + '?searchterm=' + request.term + "&OSName=" + self.getOS() + '&Culture=' + PDF.Config.culture;
                                xdr.open("GET", url);
                                xdr.onload = function ()
                                {
                                    var data = $.parseJSON(xdr.responseText);
                                    if (data == null || typeof (data) == 'undefined')
                                    {
                                        data = $.parseJSON(data.firstChild.textContent);
                                    }
                                    var responseObj = self.dgpsSuccess(request.term, data);
                                    if (responseObj)
                                    {
                                        response(responseObj);
                                    }
                                    else
                                    {
                                        response();
                                    }
                                };
                                xdr.send();
                            }
                            else
                            {
                                $.ajax({
                                    url: autoSuggestUrl,
                                    type: "GET",
                                    data: { SearchTerm: $.trim(request.term), OSName: self.getOS(), Culture: PDF.Config.culture },
                                    dataType: "json",
                                    success: function (data)
                                    {
                                        var responseObj = self.dgpsSuccess(request.term, data);
                                        if (responseObj)
                                        {
                                            response(responseObj);
                                        }
                                        else
                                        {
                                            response();
                                        }
                                    },
                                    error: function (xhr, e, ex)
                                    {
                                        self.dgpsError(xhr, e, ex);
                                        response();
                                    }
                                });
                            }
                        },
                        select: function (event, ui)
                        {
                            var cookieVal = ui.item.label;
                            if (ui.item.hasOwnProperty('parentName') && ui.item.parentName)
                            {
                                cookieVal += parentSeperator + ui.item.parentName;
                            }

                            self.AppendCookie(cookieName, cookieVal);
                            self.triggerSearch(event, ui, true);
                        },
                        minLength: 0,
                        delay: 100,
                        hasSearchButton: true,
                        autocomplete: true,
                    }).keypress(function (e)
                    {
                        if (e.keyCode === 13)
                        {
                            self.triggerSearch(e);
                        }
                    });

                    self.dgpsSuccess = function (term, data)
                    {
                        if ($.trim(term))
                        {
                            $.ui.autocomplete.autocompleteText = $.trim(data.TextToInsert);
                            self.AutocompleteTerm();
                            if (data)
                            {
                                var count = 1;
                                var obj = $.map(
                                    data.AutoSuggestions,
                                    function (item)
                                    {
                                        return {
                                            label: $.trim(item.CompletionText).toLowerCase(),
                                            value: $.trim(item.CompletionText).toLowerCase(),
                                            inputText: term,
                                            parentName: self.getParentName(item).toLowerCase(),
                                            source: self.getSource(item),
                                            index: count++,
                                        };
                                    }).splice(0, suggestionCount);
                                self.currentSuggestions = obj;
                                return obj;

                            }
                            else
                            {
                                return null;
                            }
                        } else
                        {
                            var recentSearchesObj = {};
                            // Take 5 resutls from recent searches
                            var recentSearches = self.GetSearchHistory(cookieName);
                            if (recentSearches && recentSearches.length)
                            {
                                recentSearchesObj = $.map(recentSearches,
                                    function (item)
                                    {
                                        var suggestionValue = self.htmlDecode(item);
                                        var parent = "";
                                        if (item.indexOf(parentSeperator) !== -1)
                                        {
                                            var items = item.split(parentSeperator);
                                            suggestionValue = self.htmlDecode(items[0]);
                                            parent = items[1];
                                        }

                                        return {
                                            label: suggestionValue.toLowerCase(),
                                            value: suggestionValue.toLowerCase(),
                                            inputText: "",
                                            parentName: parent.toLowerCase(),
                                            source: searchHistorySuggestSourceType,
                                            recentSearch: true
                                        };
                                    });

                                recentSearchesObj = recentSearchesObj.splice(0, recentSearchesCount);
                            }

                            var responseObj = recentSearchesObj || {};
                            var topSearchesObj = {};
                            if (data)
                            {

                                topSearchesObj = $.map(
                                    data.AutoSuggestions,
                                    function (item)
                                    {
                                        return {
                                            label: $.trim(item.CompletionText).toLowerCase(),
                                            value: $.trim(item.CompletionText).toLowerCase(),
                                            inputText: term,
                                            parentName: self.getParentName(item).toLowerCase(),
                                            source: topSearchSuggestSourceType
                                        };
                                    });

                                if (recentSearchesObj && recentSearchesObj.length)
                                {
                                    responseObj = self.getUniqueSuggestions($.merge(recentSearchesObj, topSearchesObj));
                                } else
                                {
                                    responseObj = topSearchesObj;
                                }
                            }

                            //fill the bucket with 8 unique results
                            if (responseObj)
                            {
                                responseObj = responseObj.splice(0, suggestionCount);
                                var count = 1;
                                for (var i = 0; i < responseObj.length; i++)
                                {
                                    responseObj[i].index = count++;
                                }
                                return responseObj;
                            }
                            else
                            {
                                return null;
                            }
                        }
                    };

                    self.dgpsError = function (xhr, e, ex)
                    {
                        var message = "";
                        if (xhr && xhr.hasOwnProperty('responseText') && xhr.responseText)
                        {
                            message += 'ResponseText:' + xhr.responseText + '|';
                        }
                        if (ex && ex.hasOwnProperty('message') && ex.message)
                        {
                            message += 'message:' + ex.message + '|';
                        }
                        if (ex && ex.hasOwnProperty('stack') && ex.stack)
                        {
                            message += 'stack:' + ex.stack + '|';
                        }
                        if (e)
                        {
                            message += 'error:' + e + '|';
                        }

                        PDF.Utils.LogMessage(e + ' : ' + ex.message, PDF.LogLevels.ERROR, document.URL, 0, 'DGPSServiceException');
                    };

                    self.htmlEncode = function (value)
                    {
                        return $('<div/>').text(value).html();
                    };

                    self.htmlDecode = function (value)
                    {
                        return $('<div/>').html(value).text();
                    };

                    self.AutocompleteTerm = function ()
                    {
                        var text = $.trim($.ui.autocomplete.autocompleteText);
                        var $input = $("#dgpsSearch");
                        if (text)
                        {
                            var oVal = $input.val();
                            var oLength = oVal.length;
                            var inputText = $.ui.autocomplete.inputText;
                            if (text.toLowerCase().indexOf(oVal.toLowerCase()) == -1)
                            {
                                return;
                            }
                            if (inputText && inputText.length > oLength)
                            {
                                $.ui.autocomplete.inputText = oVal;
                                return;
                            }

                            $input.val(oVal + text.substr(oLength, text.length));
                            self.setInputSelection($input[0], oLength, text.length);
                            $.ui.autocomplete.inputText = oVal;

                            if (oLength != text.length)
                            {
                                self.preAutoFillTerm = oVal;
                            }

                            if (self.currentAutosuggestion.Term != text && oLength != text.length)
                            {
                                self.PushCurrentAutosuggestion(self.currentAutosuggestion.Used);
                                self.currentAutosuggestion = {};
                                self.currentAutosuggestion.Term = text;
                                self.currentAutosuggestion.Used = false;
                            }
                        }
                        else
                        {
                            self.preAutoFillTerm = '';
                            self.PushCurrentAutosuggestion(self.currentAutosuggestion.Used);
                            self.currentAutosuggestion = {};
                        }
                    };

                    self.setInputSelection = function (input, startPos, endPos)
                    {
                        input.focus();
                        if (typeof input.selectionStart != "undefined")
                        {
                            input.selectionStart = startPos;
                            input.selectionEnd = endPos;
                        } else if (document.selection && document.selection.createRange)
                        {
                            // IE branch
                            input.select();
                            var range = document.selection.createRange();
                            range.collapse(true);
                            range.moveEnd("character", endPos);
                            range.moveStart("character", startPos);
                            range.select();
                        }
                    };

                    // Debug Code
                    //$(document).bind('dgps.suggestionclick', function (e, data)
                    //{
                    //    alert('BI Trigger received: ' + JSON.stringify(data));
                    //});

                    self.getUniqueArray = function (array)
                    {
                        var item = array.concat();
                        for (var i = 0; i < item.length; ++i)
                        {
                            for (var j = i + 1; j < item.length; ++j)
                            {
                                if (item[i] === item[j])
                                {
                                    item.splice(j--, 1);
                                }
                            }
                        }

                        return item;
                    };

                    $(document).on('click', function ()
                    {
                        self.PushCurrentAutosuggestion(true);

                    });

                    $('#dgpsSearch').keydown(function (e)
                    {
                        switch (e.which)
                        {
                            case $.ui.keyCode.LEFT: // left
                                break;

                            case $.ui.keyCode.UP: // up
                                self.PushCurrentAutosuggestion(false);
                                self.preAutoFillTerm = '';
                                break;

                            case $.ui.keyCode.RIGHT: // right
                                self.PushCurrentAutosuggestion(true);
                                self.preAutoFillTerm = '';
                                break;

                            case $.ui.keyCode.DOWN: // down
                                self.PushCurrentAutosuggestion(false);
                                self.preAutoFillTerm = '';
                                break;

                            default: return; // exit this handler for other keys
                        }
                        e.preventDefault();
                    });

                    self.PushCurrentAutosuggestion = function (val)
                    {
                        if (self.currentAutosuggestion && self.currentAutosuggestion.Term)
                        {
                            self.currentAutosuggestion.Used = val;
                            self.autoSuggestions.push(self.currentAutosuggestion);
                        }
                        self.currentAutosuggestion = {};
                    };

                    self.getUniqueSuggestions = function (array)
                    {
                        var item = array.concat();
                        for (var i = 0; i < item.length; ++i)
                        {
                            for (var j = i + 1; j < item.length; ++j)
                            {
                                if (item[i].value === item[j].value)
                                {
                                    item.splice(j--, 1);
                                }
                            }
                        }

                        return item;
                    };

                    self.triggerSearch = function (event, ui, isSuggestionClick)
                    {
                        var searchRequest = {};

                        if (ui && ui.item && ui.item.value)
                        {
                            searchRequest = ui.item;
                        }
                        else
                        {
                            if (!self.searchTerm())
                            {
                                self.searchTerm($('#dgpsSearch').val());
                            }

                            //Check if we suggestion was present in last suggestion and has L0
                            var match = PDF.Utils.ArrayFirst(self.currentSuggestions, function (item)
                            {
                                return item.value.toLowerCase() == self.searchTerm().toLowerCase();
                            });

                            var parent = '';

                            if (match && match.parentName)
                            {
                                parent = match.parentName;
                            }

                            searchRequest = { label: self.searchTerm(), inputText: self.searchTerm(), value: self.searchTerm(), parentName: parent, index: -1 };

                            self.AppendCookie(cookieName, self.searchTerm());
                        }

                        self.PushCurrentAutosuggestion(true);

                        $(document).trigger('dgps.dgpsBITrigger', { searchterm: self.searchTerm(), isSuggestoinClick: isSuggestionClick, data: ui, autofillSuggestion: self.autoSuggestions, preAutoFillTerm: self.preAutoFillTerm });

                        if (searchPageUrl) //Search url
                        {
                            var urlsearch = window.location.protocol + "//" + window.location.host + "/";
                            window.location.assign(urlsearch + searchPageUrl.format(encodeURIComponent(searchRequest.value), encodeURIComponent(searchRequest.parentName)));
                        }
                        else
                        {
                            PDF.Events.Trigger("dgps.searchTriggered", searchRequest);
                        }
                    };

                    // Debug Code
                    //PDF.Events.Bind("dgps.searchTriggered", function (e, data)
                    //{
                    //    alert('searching....' + JSON.stringify(data));
                    //});

                    self.getParentName = function (item)
                    {
                        var parentName = '';
                        if (item.hasOwnProperty('L1L3Pair') && item.L1L3Pair)
                        {
                            if (item.L1L3Pair.hasOwnProperty('Product') && item.L1L3Pair.Product)
                            {
                                if (item.L1L3Pair.Product.hasOwnProperty('ParentName') && item.L1L3Pair.Product.ParentName)
                                {
                                    parentName = item.L1L3Pair.Product.ParentName;
                                }
                            }
                        }

                        return parentName;
                    };

                    self.getSource = function (item)
                    {
                        var source = "";
                        if (item.hasOwnProperty('AutoSuggestionType') && item.AutoSuggestionType)
                        {
                            switch (item.AutoSuggestionType)
                            {
                                case 1:
                                    source = productSuggestSourceType;
                                    break;
                                case 2:
                                    source = taxonomySuggestSourceType;
                                    break;
                                case 3:
                                    source = topSearchSuggestSourceType;
                                    break;
                                default:
                                    source = 'UNKNOWN';
                            }
                        }

                        return source;
                    };

                    self.getOS = function ()
                    {
                        var os = navigator.platform;
                        var userAgent = navigator.userAgent;
                        var OSinfo = "";
                        if (os.indexOf("Win") > -1)
                        {
                            if (userAgent.indexOf("Windows NT 5.0") > -1)
                            {
                                OSinfo = "Windows 2000";
                            }
                            else if (userAgent.indexOf("Windows NT 5.1") > -1)
                            {
                                OSinfo = "Windows XP";
                            }
                            else if (userAgent.indexOf("Windows NT 5.2") > -1)
                            {
                                OSinfo = "Windows 2003";
                            }
                            else if (userAgent.indexOf("Windows NT 6.0") > -1)
                            {
                                OSinfo = "WindowsVista";
                            }
                            else if (userAgent.indexOf("Windows NT 6.1") > -1 || userAgent.indexOf("Windows 7") > -1)
                            {
                                OSinfo = "Windows 7";
                            }
                            else if (userAgent.indexOf("Windows NT 6.2") > -1 || userAgent.indexOf("Windows 8") > -1)
                            {
                                OSinfo = "Windows 8";
                            }
                            else if (userAgent.indexOf("Windows NT 6.3") > -1 || userAgent.indexOf("Windows 8") > -1)
                            {
                                OSinfo = "Windows 8.1";
                            }
                            else
                            {
                                OSinfo += "Other";
                            }
                        }
                        else if (os.indexOf("Mac") > -1)
                        {
                            OSinfo = "Mac";
                        }
                        else if (os.indexOf("X11") > -1)
                        {
                            OSinfo = "Unix";
                        }
                        else if (os.indexOf("Linux") > -1)
                        {
                            OSinfo = "Linux";
                        }
                        else
                        {
                            OSinfo = "Other";
                        }

                        return OSinfo;
                    };
                }

                if (params.hasOwnProperty("defaultSearchBoxFocus") && params.defaultSearchBoxFocus)
                {
                    // Default focus searchBox and hide flyout for this specific instance
                    $("#dgpsSearch").focus();
                    window.setTimeout(function ()
                    {
                        $("#dgpsSearch").autocomplete("close");
                    }, 10);
                }
            });

            self.watermarkText = function ()
            {
                var watermarkText = '';
                if (params.hasOwnProperty("showWatermark") && params.showWatermark)
                {
                    watermarkText = self.UI().SearchWatermarkText;
                }
                return watermarkText;
            }

            return self;
        }($);

        promise.await([
            viewModel.ExtendUIStrings("PDF.DynamicGPS.UIStrings")
        ], viewModel);
    };
})(window, jQuery, ko, PDF);

