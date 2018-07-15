///#source 1 1 pdf.core.init.js
window.PDF = window.PDF || {};

(function (window, document, $, ko, PDF, undefined)
{
    "use strict";

    // Load PDF component to element.
    // Performance optimization: Load template and viewmodel in parallel, then apply bindings.
    PDF.LoadComponent = function (el, componentName, params)
    {
        var defaultParams = {
            enableLoadingHtml: false,
            enableErrorHtml: false
        };
        params = $.extend({}, defaultParams, params);

        var inlineTemplate = el.innerHTML;
        var $el = $(el);

        if (params.enableLoadingHtml)
        {
            var loadingHtmlTempalte = PDF.GetConfig("componentLoadingHtmlTempalteFn", componentName);
            if (loadingHtmlTempalte)
            {
                var loadingText = PDF.GetConfig("componentLoadingText", componentName);
                var loadingHtml = loadingHtmlTempalte().format(loadingText);

                $el.html(loadingHtml);
            }
        }

        var modelPromise = PDF.ViewModels.GetPromise(componentName, el, params); // resolve with viewModel
        var templatePromise;

        if (params.inlineTemplate && inlineTemplate.trim())
        {
            templatePromise = PDF.Promise();
            templatePromise.resolve(inlineTemplate);
        }
        else
        {
            templatePromise = PDF.TemplateManager.GetPromise(componentName); // resolve with html content
        }

        PDF.Promise.join([modelPromise, templatePromise])
            .done(function (viewModel, template)
            {
                // apply bindings after both template and viewmodel are loaded.
                $el.html(template);

                if (viewModel.hasOwnProperty("sharedTemplates"))
                {
                    var promiseArray = PDF.Utils.LoadSharedTemplates(viewModel.sharedTemplates, el, params);
                    promiseArray.done(function ()
                    {
                        ko.applyBindings(viewModel, el);
                    });
                }
                else
                {
                    ko.applyBindings(viewModel, el);
                }
                PDF.Events.Trigger("ComponentLoaded", { componentName: componentName, arguments: arguments });
            })
            .fail(function (xhr)
            {
                // BUG: when template fails, "xhr" is a string (error message), not a xhr object.
                PDF.Utils.TemplateLoadingErrorHandler(xhr, $el, componentName, params);
            });
    };

    // create pdf component and append it to 'appendTo' element
    PDF.CreateComponent = function (componentName, appendTo, params)
    {
        var element = document.createElement('div');

        appendTo = appendTo || document.body;
        appendTo.appendChild(element);

        PDF.LoadComponent(element, componentName, params);
    };

    // keep CreatePDFControl API for backward compatibility (v1.0)
    PDF.CreatePDFControl = PDF.CreateComponent;

    // This will be used to load components for and inside element, element is selector string
    // Component at selector level and child level will be loaded.
    PDF.LoadComponentsInElement = function (element)
    {
        var componentType = ['data-pdfcomponent', 'data-pdfControl'];
        PDF.LoadComponentsByType(element, componentType);
    };

    PDF.LoadComponentOnDemand = function (element, params)
    {
        var componentType = ['data-pdfdynamiccomponent'];
        PDF.LoadComponentsByType(element, componentType, params);
    }

    PDF.LoadComponentsByType = function (element, componentType, dynamicParams)
    {
        var defaultParams = { enableLoadingHtml: true, enableErrorHtml: true, inlineTemplate: true };

        if (!componentType) return;

        $.each(componentType, function (i, attrName)
        {
            var attrSelector = '';
            if (element)
            {
                attrSelector = "{0}[{1}], {0} > [{1}]".format(element, attrName);
            }
            else
            {
                attrSelector = "[{0}]".format(attrName);
            }

            $(attrSelector).each(function (i, el)
            {
                var componentName = el.getAttribute(attrName);

                var params = defaultParams;
                var paramsString = el.getAttribute('data-params');
                if (paramsString)
                {
                    // use eval instead of JSON.parse to have less constraint on the input param, i.e.: double quoted property name.
                    try
                    {
                        var parsedParams = eval('(' + paramsString + ')');
                        params = $.extend({}, defaultParams, parsedParams);
                    }
                    catch (ex)
                    {
                        params = $.extend({}, defaultParams);
                        PDF.Utils.LogMessage(ex, PDF.LogLevels.ERROR, document.URL, ex.lineNumber);
                    }
                }

                // overwrite parameters
                params = $.extend(params, dynamicParams);

                var resourceString = el.getAttribute('data-resources');

                if (resourceString)
                {
                    var resources = eval('(' + resourceString + ')');
                    resources = $.extend({}, resources, {
                        callback: function ()
                        {
                            PDF.LoadComponent(el, componentName, params);
                        }
                    });
                    PDF.ResourceLoader.Load(resources);
                }
                else
                {
                    PDF.LoadComponent(el, componentName, params);
                }
            });
        });
    };

    $(function ()
    {
        PDF.LoadComponentsInElement(null);
    });
})(window, document, jQuery, ko, PDF);
///#source 1 1 pdf.core.config.js
(function (window, document, $, ko, PDF, undefined) {
    "use strict";

    PDF.Config = {
        debugMode: "false",
        debugModeEnableAlert: "false",
        robustBindingMode: "true",
        clientErrorLoggingEnabled: "true",
        clientErrorLoggingLevel: "error",
        defaultAnimationDuration: "1000",
        defaultEasing: 'easeInOutExpo',
        imageUrl: 'content/pdf/css/img/',
        componentLoadingHtmlTempalte: null,

        componentLoadingText: 'Loading',
        componentLoadingErrorHtmlTemplate: '<div class="error"><h3>{0}</h3><p>{1}</p></div>',
        componentLoadingErrorText1: 'Sorry',
        componentLoadingErrorText2: 'This feature is temporarily unavailable. Please try again later.',
        dynamicComponentLoadingHtmlTemplate: null,

        bloodhoundJS: "http://pidcheckerdev.cloudapp.net/scripts/bloodhound.js",
        bloodhoundAPIVersion: "2.0",
        bloodhoundPartnerId: "PDF",
        closeDialogOnOverlayClick: "true",
        contentApiUrl: "api/pdf/pdfcontent/get",
        errorLogApiUrl: "api/pdf/PDFCommonUI/LogClientMessage",
        dateTimeFormatApiUrl: "api/pdf/pdfcontent/datetimeformats",
        templateUrl: "content/pdf/template/{0}.html",
        sharedTemplateUrl: "content/pdf/template/subtemplate/{0}.html",
        isWindowsRT: "false",
        isScanableDevice: "true",
        locale: "en-us",

        componentLoadingHtmlTempalteFn: function () {
            return PDF.Config.componentLoadingHtmlTempalte || '<div tabindex="0" style="margin: 40px 0;"><img class="vlb mr10" alt="Loading" src="' + PDF.Config.imageUrl + 'progressindicator.gif" height="20" width="20" />{0}</div>';
        },
        dynamicComponentLoadingHtmlTemplateFn: function () {
            return PDF.Config.dynamicComponentLoadingHtmlTemplate || '<script type="text/html" id="dynamicComponentLoadingHtmlTemplate"><img class="vlb mr10 ml10" alt="Loading" src="' + PDF.Config.imageUrl + 'progressindicator.gif" height="20" width="20" /></script>';
        },
        biLoggingEnabled: false,
    };

    PDF.LogLevels = { ERROR: 15, WARNING: 10, INFORMATION: 5 };

    PDF.ComponentConfig = {};

    PDF.GetConfig = function (key, componentType) {
        /// <summary>Get config value from PDF.Config</summary>
        /// <param name="key" type="string">config key</param>
        /// <param name="componentType" type="string"> component type, optional</param>

        if (componentType && PDF.ComponentConfig[componentType] && PDF.ComponentConfig[componentType][key]) {
            return PDF.ComponentConfig[componentType][key];
        }
        else {
            return PDF.Config[key];
        }
    };

    PDF.GetComponentConfig = function (componentType) {
        /// <summary>Get component config from PDF.Config</summary>
        /// <param name="componentType" type="string"> component type, optional</param>

        if (componentType && PDF.ComponentConfig[componentType]) {
            return PDF.ComponentConfig[componentType];
        }
        else {
            return '';
        }
    };

    PDF.GetConfigInInt = function (key, componentType, defaultValue) {
        /// <summary>Get config value in integer type from PDF.Config</summary>
        /// <param name="key" type="string">config key</param>
        /// <param name="componentType" type="string"> component type, optional</param>
        /// <param name="defaultValue" type="integer">default value</param>

        var configValue = PDF.GetConfig(key, componentType);
        try {
            return parseInt(configValue);
        }
        catch (e) {
            PDF.Utils.Alert(e);
            return defaultValue || 0;
        }
    };

    PDF.SetConfig = function (key, value, componentType) {
        /// <summary>Set config value into PDF.Config</summary>
        /// <param name="key" type="string">config key, support "componentName.key" format</param>
        /// <param name="value" type="Object">config value</param>
        /// <param name="componentType" type="string"> component type, optional</param>

        var keys = key.split('.');
        if (keys.length == 2) {
            if (!componentType || keys[0] === componentType) {
                componentType = keys[0];
                key = keys[1];
            }
        }

        if (componentType) {
            if (!PDF.ComponentConfig[componentType]) {
                PDF.ComponentConfig[componentType] = {};
                PDF.ComponentConfig[componentType].GetConfig = function (key) {
                    return PDF.GetConfig(key, componentType);
                };
            }

            PDF.ComponentConfig[componentType][key] = value;
        }
        else {
            PDF.Config[key] = value;
        }
    };
})(window, document, jQuery, ko, PDF);
///#source 1 1 pdf.core.customextension.js
// custom extensions
(function (window, document, $, ko, PDF, undefined)
{
    "use strict";
    if (typeof String.prototype.format === 'undefined')
    {
        String.prototype.format = function ()
        {
            var args = arguments;
            return this.replace(/{(\d+)}/g, function (match, number)
            {
                return typeof args[number] !== 'undefined'
                  ? args[number]
                  : match
                ;
            });
        };
    }

    // For IE 8 Compatibility
    if (!Array.prototype.filter)
    {
        Array.prototype.filter = function (fun)
        {
            if (this == null)
            {
                throw new TypeError();
            }

            var t = Object(this);
            var len = t.length >>> 0;
            if (typeof fun !== "function")
            {
                throw new TypeError();
            }

            var res = [];
            var thisp = arguments[1];
            for (var i = 0; i < len; i++)
            {
                if (i in t)
                {
                    var val = t[i]; // in case fun mutates this
                    if (fun.call(thisp, val, i, t))
                    {
                        res.push(val);
                    }
                }
            }

            return res;
        };
    }

    // For IE 8 Compatibility
    if (!Array.prototype.indexOf)
    {
        Array.prototype.indexOf = function (searchElement /*, fromIndex */)
        {
            if (this == null)
            {
                throw new TypeError();
            }
            var t = Object(this);
            var len = t.length >>> 0;
            if (len === 0)
            {
                return -1;
            }
            var n = 0;
            if (arguments.length > 1)
            {
                n = Number(arguments[1]);
                if (n != n)
                { // shortcut for verifying if it's NaN
                    n = 0;
                }
                else if (n != 0 && n != Infinity && n != -Infinity)
                {
                    n = (n > 0 || -1) * Math.floor(Math.abs(n));
                }
            }
            if (n >= len)
            {
                return -1;
            }
            var k = n >= 0 ? n : Math.max(len - Math.abs(n), 0);
            for (; k < len; k++)
            {
                if (k in t && t[k] === searchElement)
                {
                    return k;
                }
            }
            return -1;
        }
    }
    // For IE 8 & 9 Compatibility
    // IE 9 only accept mili-second in 3-digit format, should use first fromISO function for compatibility
    var D = new Date('2011-06-02T09:34:29.12+02:00');
    if (!D || +D !== 1307000069120)
    {
        Date.fromISO = function (s)
        {
            var day, tz,
            rx = /^(\d{4}\-\d\d\-\d\d([tT ][\d:\.]*)?)([zZ]|([+\-])(\d\d):(\d\d))?$/,
            p = rx.exec(s) || [];
            if (p[1])
            {
                day = p[1].split(/\D/);
                for (var i = 0, L = day.length; i < L; i++)
                {
                    day[i] = parseInt(day[i], 10) || 0;
                }
                day[1] -= 1;
                day = new Date(Date.UTC.apply(Date, day));
                if (!day.getDate()) return NaN;
                if (p[5])
                {
                    tz = (parseInt(p[5], 10) * 60);
                    if (p[6]) tz += parseInt(p[6], 10);
                    if (p[4] == '+') tz *= -1;
                    if (tz) day.setUTCMinutes(day.getUTCMinutes() + tz);
                }
                return day;
            }
            return NaN;
        };
    }
    else
    {
        Date.fromISO = function (s)
        {
            return new Date(s);
        };
    }

    if (typeof Object.create !== 'function')
    {
        Object.create = function (o)
        {
            function F() { }
            F.prototype = o;
            return new F();
        };
    }

    if (typeof Function.prototype.SingleCall !== 'function')
    {
        /// make the function only called once.
        Function.prototype.SingleCall = function ()
        {
            // save original function
            var f = this;
            // Set a flag to indicate if the function is called.
            var isCalled = false;

            return function ()
            {
                if (!isCalled)
                {
                    isCalled = true;
                    f.apply(this, arguments);
                }
            };
        };
    }

    if (!String.prototype.trim)
    {
        String.prototype.trim = function ()
        {
            return this.replace(/^\s+|\s+$/g, '');
        };
    }

    $.getJSONSync = function (url, data, callback)
    {
        return PDF.Utils.AjaxWithCredential({
            url: url,
            data: data,
            dataType: "json",
            async: false,
            success: callback
        });
    };

    // extend jQuery easing effect
    $.extend($.easing, {
        easeInOutExpo: function (x, t, b, c, d)
        {
            if (t == 0) return b;
            if (t == d) return b + c;
            if ((t /= d / 2) < 1) return c / 2 * Math.pow(2, 10 * (t - 1)) + b;
            return c / 2 * (-Math.pow(2, -10 * --t) + 2) + b;
        }
    });

    // Wrap the ko parseBindingsString to make code robust.
    var ko_parseBindingsString = ko.bindingProvider.prototype.parseBindingsString;
    ko.bindingProvider.prototype.parseBindingsString = function (bindingsString, bindingContext)
    {
        try
        {
            return ko_parseBindingsString.bind(this)(bindingsString, bindingContext);
        }
        catch (ex)
        {
            PDF.Utils.Alert(ex);
            if (PDF.Config.robustBindingMode === "false" || PDF.Config.robustBindingMode === false)
            {
                throw ex;
            }
            return null;
        }
    }

    // Override ko template engine's make template source method
    var oldKOMakeTemplateSourceFn = ko.templateEngine.prototype.makeTemplateSource;
    ko.templateEngine.prototype.makeTemplateSource = function (template, templateDocument)
    {
        // Named template
        if (typeof template == "string")
        {
            // 1. If template is already present in DOM return it
            var elem = document.getElementById(template);

            if (!elem)
            {
                if (!window.IsDynamicComponentLoadingHtmlTemplate)
                {
                    window.IsDynamicComponentLoadingHtmlTemplate = true;
                    $('body').append(PDF.Config.dynamicComponentLoadingHtmlTemplateFn());
                }

                var loadingContainer = document.getElementById('dynamicComponentLoadingHtmlTemplate');

                // 2. If template is not present in the Dom resolve file name where template belongs to from template repository
                // 3. Trigger template manager to load template and return loading container as response
                // 4. When promise is resolved trigger observable (by setting new value) so that dependency tracker will re-evaluate template and its binding and 
                //    downloaded template will be bind with available data context.

                var isTemplateLoaded = ko.observable(false);

                var el = new ko.templateSources.domElement(loadingContainer);
                el.text = function ()
                {
                    // call isTemplateLoaded(); to register the dependency
                    isTemplateLoaded();
                    return ko.templateSources.domElement.prototype['text'].apply(this, arguments);
                };


                PDF.Utils.LoadDynamicTemplate(template).done(function (templateElement)
                {
                    el.domElement = templateElement;
                    isTemplateLoaded(true);
                });

                return el;
            }
            else
            {
                return new ko.templateSources.domElement(elem);
            }
        }
        else
        {
            return oldKOMakeTemplateSourceFn.apply(this, arguments);
        }
    };

    // This will have a side effect that this will be executed at website level and not specifically for PDF component 
    window.onerror = function (err, url, line)
    {
        PDF.Utils.LogMessage(err, PDF.LogLevels.ERROR, url, line);
    }

    // window.onerror is not working as expected in previous versions of mozilla (V5 and earlier due to a bug) and safari (doesn't support window.onerror)
    // Wrap all functions with event handler and log the exception based on config value
    var jQueryBind = jQuery.fn.bind;
    jQuery.fn.bind = function (type, data, fn)
    {
        if (!PDF.Config.clientErrorLoggingEnabled)
        {
            return;
        }

        if (!fn && data && typeof data === 'function')
        {
            fn = data;
            data = null;
        }

        if (fn)
        {
            var baseFunction = fn;
            var wrapperFunction = function ()
            {
                try
                {
                    baseFunction.apply(this, arguments);
                }
                catch (ex)
                {
                    PDF.Utils.LogMessage(ex, PDF.LogLevels.ERROR, document.URL, ex.lineNumber);
                    throw ex;
                }
            };

            fn = wrapperFunction;
        }

        return jQueryBind.call(this, type, data, fn);
    };

    // lazy load observable which is only loaded when first bound
    ko.observableOnDemand = function (callback, target)
    {
        var value = ko.observable();  //private observable

        var result = ko.computed({
            read: function ()
            {
                //if it has not been loaded, execute callback to load it
                if (!result.loaded())
                {
                    callback.call(target);
                }
                //always return the current value
                return value();
            },
            write: function (newValue)
            {
                //indicate that the value is now loaded and set it
                result.loaded(true);
                value(newValue);
            },
            deferEvaluation: true  //do not evaluate immediately when created
        });

        //expose the current state, which can be bound against
        result.loaded = ko.observable();

        // refresh the data from server
        result.refresh = function ()
        {
            result.loaded(false);
        };

        return result;
    };

    // Create an observable with On/Off methods.
    ko.observableToggle = function (defaultValue)
    {
        var value = ko.observable(defaultValue);

        value.On = function ()
        {
            value(true);
        };

        value.Off = function ()
        {
            value(false);
        };

        return value;
    };

    //wrapper to an observable that requires accept/cancel
    ko.protectedObservable = function (initialValue)
    {
        var actualValue = ko.observable(initialValue),
            tempValue = initialValue;

        var result = ko.computed({
            read: function ()
            {
                return actualValue();
            },
            write: function (newValue)
            {
                tempValue = newValue;
            }
        });

        result.commit = function ()
        {
            if (tempValue !== actualValue())
            {
                actualValue(tempValue);
            }
        };

        result.reset = function ()
        {
            actualValue.valueHasMutated();
            tempValue = actualValue();
        };

        return result;
    };

    // add "required" validation
    ko.extenders.required = function (target, message)
    {
        target.hasError = ko.observable();
        target.validationMessage = ko.observable();

        function validate(newValue)
        {
            target.hasError(newValue == "" ? true : false);
            target.validationMessage(message);
        }

        target.subscribe(validate);

        return target;
    };

    // Wrap Promise for later extend to WinJS.
    PDF.Promise = function ()
    {
        return $.extend($.Deferred(), {
            // wait for multiple promises and resolve with data
            await: function (promiseArray, data)
            {
                var callerPromise = this;
                PDF.Promise.join(promiseArray)
                .done(function ()
                {
                    callerPromise.resolve(data);
                })
                .fail(callerPromise.reject);
            }
        });
    };

    // Provide a way to combine multiple promise for synchronization
    PDF.Promise.join = function (promiseArray)
    {
        return $.when.apply($, promiseArray);
    };

    PDF.HashHandlers = function ()
    {
        var prevFragmentIdentifierName = '';
        var components = {};
        function TriggerHashChangeHanlder(fragmentIdentifier)
        {
            if (!fragmentIdentifier.name && prevFragmentIdentifierName !== '')
            {
                $.each(components, function (key, value)
                {
                    components[key.toLowerCase()](null);
                });
                prevFragmentIdentifierName = '';
            }
            else if (fragmentIdentifier && fragmentIdentifier.name)
            {
                prevFragmentIdentifierName = fragmentIdentifier.name.toLowerCase();
                var handler = components[prevFragmentIdentifierName];
                if (typeof handler === 'function')
                {
                    handler(fragmentIdentifier.items);
                }
            }
        }

        $(window).bind('hashchange', function ()
        {
            var parameters = GetHashParameters();
            TriggerHashChangeHanlder(parameters);
        });

        function GetHashParameters()
        {
            var fragmentIdentifier = {
                name: null,
                items: []
            };

            var hash = decodeURIComponent(window.location.hash);

            if (!hash)
            {
                return fragmentIdentifier;
            }

            var expressions = hash.substring(1).split('/');

            if (expressions.length > 0)
            {
                fragmentIdentifier.name = expressions[0].toLocaleLowerCase();
            }

            for (var i = 1; i < expressions.length; i++)
            {
                fragmentIdentifier.items[i - 1] = expressions[i].toLocaleLowerCase();
            }

            return fragmentIdentifier;
        }

        var self = {
            GetHashUrlString: function (componentName, items)
            {
                if (!componentName)
                {
                    return;
                }

                items.unshift(componentName);

                items = $.map(items, function (item)
                {
                    return encodeURIComponent(item);
                });

                var hashString = '#' + items.join('/');

                return hashString.toLowerCase();
            },

            InitializeHashUrl: function (componentName, items)
            {
                var hashUrl = self.GetHashUrlString(componentName, items);
                window.location.hash = hashUrl;
            },

            AddComponent: function (name, handler)
            {
                var lName = name.toLowerCase();
                if (lName && handler)
                {
                    components[lName] = handler;
                    var fragmentIdentifier = GetHashParameters();

                    if (!fragmentIdentifier.name || fragmentIdentifier.name === lName)
                    {
                        components[lName](fragmentIdentifier.items);
                    }
                }
            }
        };

        return self;
    }();

    // PDF events
    PDF.Events = (function ()
    {
        var getEventType = function (type)
        {
            var eventPrefix = "PDF.";
            return eventPrefix + type;
        };

        return {
            Bind: function (type, data, fn)
            {
                var eventType = getEventType(type);
                $(document).bind(eventType, data, fn);
            },
            Trigger: function (type, data)
            {
                var eventType = getEventType(type);
                $(document).trigger(eventType, data);
            }
        };
    })();

    // Resource Loader
    PDF.ResourceLoader = function ()
    {
        var self = {

            LoadJavaScript: (function ()
            {
                var existingSrc = {};
                var isLoaded = function (srcPath)
                {
                    return existingSrc.hasOwnProperty(srcPath.toLowerCase());
                };

                return function (srcPath, callback, errorCallback)
                {
                    if (isLoaded(srcPath))
                    {
                        callback();
                        return;
                    }

                    existingSrc[srcPath] = true;

                    var script = document.createElement('script');
                    script.type = 'text/javascript';
                    script.src = srcPath;
                    if (typeof callback === 'function')
                    {
                        if (typeof (script.onload) !== 'undefined')
                        {
                            script.onload = callback;

                            if (typeof errorCallback === 'function' && typeof (script.onerror) !== 'undefined')
                            {
                                script.onerror = errorCallback;
                            }
                        }
                        else
                        {
                            script.onreadystatechange = function (e)
                            {
                                if (this.readyState === 'complete' || this.readyState === 'loaded')
                                {
                                    callback();
                                }
                                else
                                {
                                    /// need to handle IE8
                                    if (typeof errorCallback === 'function' && this.readyState !== 'loading')
                                    {
                                        errorCallback(e);
                                    }
                                }
                            };
                        }
                    }
                    document.getElementsByTagName("head")[0].appendChild(script);
                };
            })(),

            LoadStyleSheet: function (path, callback, errorCallback)
            {
                var l = $('link[href="' + path + '"]');

                if (!l.length)
                {

                    var head = document.getElementsByTagName('head')[0],
                        link = document.createElement('link');
                    link.setAttribute('href', path);
                    link.setAttribute('rel', 'stylesheet');
                    link.setAttribute('type', 'text/css');

                    var sheet, cssRules;
                    if ('sheet' in link)
                    {
                        sheet = 'sheet'; cssRules = 'cssRules';
                    }
                    else
                    {
                        sheet = 'styleSheet'; cssRules = 'rules';
                    }

                    var timeout_id = setInterval(function ()
                    {
                        try
                        {
                            if (link[sheet] && link[sheet][cssRules].length)
                            {
                                clearInterval(timeout_id);
                                clearTimeout(timeout_id);
                                callback.call();
                            }
                        } catch (e) { } finally { }
                    }, 10),
                        timeout_id = setTimeout(function ()
                        {
                            clearInterval(timeout_id);
                            clearTimeout(timeout_id);
                            head.removeChild(link);
                            errorCallback.call();
                        }, 15000);

                    head.appendChild(link);

                    return link;
                }

                return l;
            },

            // Example:
            //PDF.ResourceLoader.Load({
            //    css: ['/content/pdf/css/pdf.base.css'],
            //    //js: ['/Script/component1.js'],
            //    callback: function ()
            //    {
            //        alert('resources loaded');
            //    }
            //});

            Load: function (resources)
            {
                var count = 0;
                var scriptTag, linkTag;
                var jsResources = resources.js;
                var cssResources = resources.css;
                var head = document.getElementsByTagName('head')[0];
                var jsLoadCount = 0;
                for (var r in cssResources)
                {
                    var link = $('link[href="' + cssResources[r] + '"]');
                    if (!link.length)
                    {
                        linkTag = document.createElement('link');
                        linkTag.type = 'text/css';
                        linkTag.rel = 'stylesheet';
                        linkTag.href = cssResources[r];
                        head.appendChild(linkTag);

                        if (!jsResources)
                        {
                            linkTag.onload = function ()
                            {
                                resources.callback.call();
                            }
                        }
                        else if (!jsResources.length)
                        {
                            linkTag.onload = function ()
                            {
                                resources.callback.call();
                            }
                        }
                    }
                }

                for (var r in jsResources)
                {
                    var script = $('script[src="' + jsResources[r] + '"]');
                    if (!script.length)
                    {
                        scriptTag = document.createElement('script');
                        scriptTag.type = 'text/javascript';
                        jsLoadCount++;
                        if (typeof resources.callback == "function")
                        {
                            if (scriptTag.readyState)
                            {  //IE
                                scriptTag.onreadystatechange = function ()
                                {
                                    if (scriptTag.readyState == "loaded" || scriptTag.readyState == "complete")
                                    {
                                        count++;
                                        if (count == jsLoadCount)
                                        {
                                            resources.callback.call();
                                        }
                                    }
                                };
                            }
                            else
                            {
                                scriptTag.onload = function ()
                                {
                                    count++;
                                    if (count == jsLoadCount)
                                    {
                                        resources.callback.call();
                                    }
                                }
                            }
                        }
                        scriptTag.src = jsResources[r];
                        head.appendChild(scriptTag);
                    }
                }
            }
        }
        return self;
    }();
})(window, document, jQuery, ko, PDF);

///#source 1 1 pdf.core.baseviewmodel.js
(function (window, document, $, ko, PDF, undefined) {
    // ViewModels constructor
    var ViewModelManager = function () {
    };

    // promise data is view model obj.
    ViewModelManager.prototype.GetPromise = function (componentName, el, params) {
        var promise = PDF.Promise();

        // component code will register getViewModel function in to PDF.ViewModels
        var getViewModel = this[componentName];

        if (typeof getViewModel === 'function') {
            getViewModel(promise, el, params);
        }
        else {
            promise.reject("componentName: {0} is not defined".format(componentName));
        }
        return promise;
    };

    // Component will register their GetViewModel function into ViewModels
    PDF.ViewModels = new ViewModelManager();

    // Base View Model as prototype of all component view model.
    // Designed for sharing common method and data.
    PDF.BaseViewModel = {
        ExtendUIStrings: function (key, propertyName) {
            propertyName = propertyName || "UI";
            var url = PDF.Config.contentApiUrl || "api/pdf/pdfcontent/get";
            var self = this;
            self.UI = ko.observable();

            return $.getJSON(url, { key: key, locale: PDF.Config.locale }, function (data) {
                var uiStrings = {};
                if (data && data.Items) {
                    ko.utils.arrayForEach(data.Items, function (item) {
                        uiStrings[item.name] = item.value;
                    });
                }

                self.UI(uiStrings);
            });
        },
        ExtendContentModel: function (key, propertyName) {
            propertyName = propertyName || "Content";
            var url = PDF.Config.contentApiUrl || "api/pdf/pdfcontent/get";
            var self = this;
            self[propertyName] = ko.observable();

            function convert(data) {
                if (data && data.Items) {
                    ko.utils.arrayForEach(data.Items, function (item) {
                        data[item.name] = convert(item);
                    });
                }
                return data;
            }

            return $.getJSON(url, { key: key, locale: PDF.Config.locale }, function (data) {
                var res = convert(data);
                self[propertyName](res);
            });
        },
        ExtendBIContentModel: function (key, propertyName) {
            if (PDF.Config.biLoggingEnabled) {
                propertyName = propertyName || "BIContent";
                var url = PDF.Config.contentApiUrl || "api/pdf/pdfcontent/get";
                var self = this;
                self[propertyName] = ko.observable();

                function convert(data) {
                    if (data && data.Items) {
                        ko.utils.arrayForEach(data.Items, function (item) {
                            data[item.name] = convert(item);
                        });
                    }
                    return data;
                }

                return $.when(
                                $.getJSON(url, { key: PDF.Config.globalBIContentKey || 'GlobalBIContent' }, function (data) {
                                    var res = convert(data);
                                    self[propertyName]($.extend(self[propertyName](), res));
                                }),
                                $.getJSON(url, { key: key }, function (data) {
                                    var res = convert(data);
                                    self[propertyName]($.extend(self[propertyName](), res));
                                })
                            );
            }
        },
        ShowGrid: function () {
            var key = "pdfgrid";
            var val = PDF.Utils.UrlParam(key);
            if (val === "on" || val === "off") {
                PDF.Utils.SetCookie(key, val);
            }
            return PDF.Utils.GetCookie(key) === "on";
        },
        ExtendErrorSupport: function (getDefaultErrorMessage) {
            var self = this;
            self.error = ko.observableToggle(false);
            self.error.errorMessage = ko.observable();

            self.ShowError = function (errorMessage) {
                if (!errorMessage && typeof getDefaultErrorMessage === "function") {
                    errorMessage = getDefaultErrorMessage();
                }
                self.error.errorMessage(errorMessage);
                self.error.On();
            };
        },
        ExtendDateTimeFormats: function () {
            var self = this;
            self.DateTimeFormats = ko.observable();
            var getUrl = PDF.Utils.GetContextUrlFn("datetimeformats", "dataContextBaseUrl");

            return $.getJSON(getUrl("api/pdf/pdfcontent/datetimeformats"), function (data) {
                var dateTimeFormats = {};
                if (data) {
                    ko.utils.objectForEach(data, function (key, value) {
                        dateTimeFormats[key] = value;
                    });

                    self.DateTimeFormats(dateTimeFormats);
                }
            });
        }

    };

})(window, document, jQuery, ko, PDF);
///#source 1 1 pdf.core.custombinding.js
// custom extensions
(function (window, document, $, ko, PDF, undefined) {
    "use strict";
    // Use slide down/up to show/hide element
    ko.bindingHandlers.slideVisible = {
        init: function (element, valueAccessor) {
            ko.bindingHandlers.visible.update(element, valueAccessor);
        },
        update: function (element, valueAccessor, allBindingsAccessor) {
            var obj = ko.utils.unwrapObservable(valueAccessor());
            var duration = PDF.GetConfigInInt('defaultAnimationDuration', null, 1000);
            if (obj) {
                $(element).slideDown(duration, PDF.Config.defaultEasing);
            }
            else {
                $(element).slideUp(duration, PDF.Config.defaultEasing);
            }
        }
    };

    // For debugging purpose
    ko.bindingHandlers.alert = {
        update: function (element, valueAccessor, allBindingsAccessor) {
            var obj = ko.utils.unwrapObservable(valueAccessor());
            if (obj) {
                // "ko.toJSON" could serialize all observable properties
                PDF.Utils.Alert(ko.toJSON(obj));
            }
        }
    };

    ko.bindingHandlers.watermark = {
        init: function (element, valueAccessor) {
            var watermarkText = ko.utils.unwrapObservable(valueAccessor());
            var waterMarkedText = $(element);
            waterMarkedText.focus(function () {
                $(element).filter(function () {
                    return $(element).val() === "" || $(element).val() === watermarkText;
                }).removeClass("wm").val("").select();
            });

            waterMarkedText.blur(function () {
                $(element).filter(function () {
                    return $(element).val() === "";
                }).addClass("wm").val(watermarkText);
            });
        },
        update: function (element, valueAccessor) {
            var watermarkText = ko.utils.unwrapObservable(valueAccessor());
            var waterMarkedText = $(element);
            if (waterMarkedText.val() === "" || waterMarkedText.val() === watermarkText) {
                waterMarkedText.addClass("wm").val(watermarkText);
            }
        }
    };

    ko.bindingHandlers.carousel = {
        update: function (element, valueAccessor, allBindingsAccessor) {
            var init = ko.utils.unwrapObservable(allBindingsAccessor().init);

            if (init && (typeof init.length === 'undefined' || init.length >= 0)) {
                var defaultOptions = {
                    animation: PDF.GetConfigInInt('defaultAnimationDuration', null, 1000),
                    easing: PDF.Config.defaultEasing
                };

                var options = $.extend(defaultOptions, ko.toJS(valueAccessor()) || {});

                window.setTimeout(function () {
                    $(element).pdfCarousel(options);
                });
            }
        }
    };

    ko.bindingHandlers.dialog = {
        update: function (element, valueAccessor, allBindingsAccessor) {
            // all binding 
            var allBindings = allBindingsAccessor();

            var showDialog = ko.utils.unwrapObservable(allBindings.visible);

            var $element = $(element);
            var options = ko.utils.unwrapObservable(valueAccessor()) || {};
            options.close = function () {
                var visible = allBindings.visible;
                if (visible) {
                    visible(false);
                }

                if (typeof (options.closeWrapper) === 'function') {
                    options.closeWrapper();
                }
            };

            options.autoOpen = false;
            options.closeIconPath = options.closeIconPath || PDF.Config.imageUrl + "ui-dialog-titlebar-close.png";
            options.closeOnOverlayClick = PDF.Config.closeDialogOnOverlayClick;

            if ($element.pdfDialog) {
                if (showDialog) {
                    $element.pdfDialog(options);
                    // make the open action async to allow the dialog calculate the content height correctly before open it.
                    window.setTimeout(function () {
                        $element.pdfDialog("open");
                    }, 0);
                }
                else {
                    $element.pdfDialog(options).pdfDialog("close");
                }
            }
        }
    };

    ko.bindingHandlers.control = {
        init: function (element, valueAccessor, allBindingsAccessor, viewModel, bindingContext) {
            // Check if template is in DOM
            var template = document.getElementById(valueAccessor().name);

            if (!template) {
                // If template is not in DOM store current inner html in temporary storage
                var originalElementInnerHtml = element.innerHTML;

                // Render control template which will render maketemplate source override to load dynamic template
                ko.renderTemplate(valueAccessor().name, bindingContext, {
                    afterRender: function (nodes) {
                        // After render will be executed twice 
                        // 1. when make template source has issues async call to download template loading html will be appended to control wrapper
                        // 2. when actual template html is part of DOM at this moment replace wrapper html with template html and replace inner 
                        //    html from temporary storage.
                        var controlTemplate = document.getElementById(valueAccessor().name);
                        if (!controlTemplate) {
                            controlTemplate = PDF.Config.dynamicComponentLoadingHtmlTemplateFn();
                        }
                        else {
                            controlTemplate = controlTemplate.innerHTML.replace('{{$pdfinnerHtml}}', originalElementInnerHtml);
                        }
                        element.innerHTML = controlTemplate;

                        // Apply binding to descendants that were added at run time.
                        ko.applyBindingsToDescendants(bindingContext, element);
                    }
                }, element);

            }
            else {
                var controlTemplate = document.getElementById(valueAccessor().name).innerHTML;
                var updatedTemplate = controlTemplate.replace('{{$pdfinnerHtml}}', element.innerHTML);
                element.innerHTML = updatedTemplate;
            }

            return ko.bindingHandlers.template.init.call(this, element, ko.observable({ "if": true }), allBindingsAccessor, viewModel, bindingContext);
        },
        update: function (element, valueAccessor, allBindingsAccessor, viewModel, bindingContext) {
            return ko.bindingHandlers.template.update.call(this, element, ko.observable({ "if": true }), allBindingsAccessor, viewModel, bindingContext);
        }
    };

    ko.bindingHandlers.bi = {
        update: function (element, valueAccessor) {
            if (PDF.Config.biLoggingEnabled) {
                // sample code of how to use this binding: <a data-bind="bi:{title: $data.anything, test:'123'}"></a>
                var tags = ko.utils.unwrapObservable(valueAccessor()) || {};

                for (var key in tags) {
                    element.setAttribute("bi:" + key, ko.utils.unwrapObservable(tags[key]));
                }
            }
        }
    };
})(window, document, jQuery, ko, PDF);
///#source 1 1 pdf.core.templatemanager.js
(function (window, document, $, ko, PDF, undefined)
{
    "use strict";

    // allow consuming project customize template
    PDF.CustomTemplates = {};

    // Manage all template for PDF components.
    PDF.TemplateManager = {
        cache: {},
        GetPromise: function (componentName, GetTemplateURL)
        {
            var url;

            var customTemplate = PDF.CustomTemplates[componentName];
            if (customTemplate)
            {
                if (typeof (customTemplate) === 'object')
                {
                    url = customTemplate.url;
                }
                else
                {
                    url = customTemplate;
                }
            }
            else
            {
                if (GetTemplateURL && typeof GetTemplateURL === 'function')
                {
                    url = GetTemplateURL(componentName);
                }

                if (!url)
                {
                    var baseUrl = PDF.GetConfig('templateBaseUrl', componentName) || '';
                    if (baseUrl)
                    {
                        url = PDF.Utils.GetContextUrl(baseUrl, PDF.Config.templateUrl, "content/pdf/template/{0}.html").format(componentName);
                    }
                    else
                    {
                        url = (PDF.Config.templateUrl || "content/pdf/template/{0}.html").format(componentName);
                    }
                }
            }

            // cache the promise with the html data
            var promise = this.cache[url];

            if (promise)
            {
                return promise;
            }
            else
            {
                this.cache[url] = promise = PDF.Promise();
            }

            $.get(url)
            .done(function (data)
            {
                promise.resolve(data);
            })
            .fail(function (data)
            {
                PDF.Utils.Alert(data);
                promise.reject(data);
            });

            return promise;
        }
    };

    // Shared template repository
    PDF.SharedTemplateRepository = function ()
    {
        var sharedTemplateRepository = {};

        var self = {
            GetTemplateFileNameByTemplateName: function (templateName)
            {
                // If shared template is not in repository than likely it has its own dedicated physical file
                // return template name in that case and Template manager will find the html of template
                if (sharedTemplateRepository.hasOwnProperty(templateName.toLowerCase()))
                {
                    return sharedTemplateRepository[templateName.toLowerCase()];
                }
                else
                {
                    return templateName;
                }
            },

            RegisterTemplate: function (templateName, containerTemplateFileName)
            {
                var lName = templateName.toLowerCase();
                if (lName && containerTemplateFileName)
                {
                    sharedTemplateRepository[lName] = containerTemplateFileName;
                }
            }
        };

        return self;
    }();

    PDF.SharedTemplateRepository.RegisterTemplate('linkTemplate', 'PDF.UI.Common.Templates');
    PDF.SharedTemplateRepository.RegisterTemplate('linkTemplateInline', 'PDF.UI.Common.Templates');
    PDF.SharedTemplateRepository.RegisterTemplate('PDFProductSelector', 'PDF.UI.Common.Templates');
    PDF.SharedTemplateRepository.RegisterTemplate('mesgboxTemplate', 'PDF.UI.Common.Templates');
    
})(window, document, jQuery, ko, PDF);
///#source 1 1 pdf.core.utils.js
(function (window, document, $, ko, PDF, undefined)
{
    "use strict";

    PDF.Utils = {
        SetCookie: function (name, value, exdays)
        {
            var _value = value === null ? "" : escape(value);
            if (typeof exdays !== 'undefined')
            {
                var exdate = new Date();
                exdate.setDate(exdate.getDate() + exdays);
                _value = _value + '; expires=' + exdate.toUTCString();
            }
            document.cookie = name + "=" + _value + ';Path=/;';
        },

        GetCookie: function (name)
        {
            var i, x, y, cookies = document.cookie.split(";");
            for (i = 0; i < cookies.length; i++)
            {
                x = cookies[i].substr(0, cookies[i].indexOf("="));
                y = cookies[i].substr(cookies[i].indexOf("=") + 1);
                x = x.replace(/^\s+|\s+$/g, "");
                if (x === name)
                {
                    return unescape(y);
                }
            }
        },

        UrlParam: function (name)
        {
            return decodeURIComponent(
               (location.search.match(RegExp("[?|&]" + name + '=(.+?)(&|$)')) || [, null])[1]
           );
        },

        // check if any item in array matches predicate
        ArrayExists: function (array, predicate)
        {
            for (var i = 0, j = array.length; i < j; i++)
            {
                if (predicate(array[i]))
                {
                    return true;
                }
            }
            return false;
        },

        // find first item which matches predicate
        ArrayFirst: function (array, predicate)
        {
            for (var i = 0, j = array.length; i < j; i++)
            {
                if (predicate(array[i]))
                {
                    return array[i];
                }
            }
            return null;
        },

        // check if array contains the value
        ArrayContains: function (array, value)
        {
            for (var i = 0, j = array.length; i < j; i++)
            {
                if (array[i] == value) // id compare is done between string and int
                {
                    return true;
                }
            }
            return false;
        },

        // Only work in debug mode
        Alert: function (obj)
        {
            if (PDF.Config.debugMode === 'true' || PDF.Config.debugMode === true)
            {
                var msg = (typeof obj === 'string') ? obj : JSON.stringify(obj);
                if (console && console.log)
                {
                    console.log(msg);
                    PDF.Utils.LogMessage(msg, PDF.LogLevels.ERROR, document.URL, obj.lineNumber);
                }
                if (PDF.Config.debugModeEnableAlert === 'true' || PDF.Config.debugModeEnableAlert === true)
                {
                    alert(msg);
                }
            }
        },

        DateDiff: function (date1, date2)
        {
            var ONE_DAY = 1000 * 60 * 60 * 24;

            var d1 = new Date(date1);
            var d2 = new Date(date2);

            var difference_ms = (d1.getTime() - d2.getTime());

            return Math.round(difference_ms / ONE_DAY);
        },

        FormatDate: function (date, isDateInLocalFormat)
        {
            if (date == null)
            {
                return false;
            }

            var f = { "ar-SA": "dd/MM/yy", "bg-BG": "dd.M.yyyy", "ca-ES": "dd/MM/yyyy", "zh-TW": "yyyy/M/d", "cs-CZ": "d.M.yyyy", "da-DK": "dd-MM-yyyy", "de-DE": "dd.MM.yyyy", "el-GR": "d/M/yyyy", "en-US": "M/d/yyyy", "fi-FI": "d.M.yyyy", "fr-FR": "dd/MM/yyyy", "he-IL": "dd/MM/yyyy", "hu-HU": "yyyy. MM. dd.", "is-IS": "d.M.yyyy", "it-IT": "dd/MM/yyyy", "ja-JP": "yyyy/MM/dd", "ko-KR": "yyyy-MM-dd", "nl-NL": "d-M-yyyy", "nb-NO": "dd.MM.yyyy", "pl-PL": "yyyy-MM-dd", "pt-BR": "d/M/yyyy", "ro-RO": "dd.MM.yyyy", "ru-RU": "dd.MM.yyyy", "hr-HR": "d.M.yyyy", "sk-SK": "d. M. yyyy", "sq-AL": "yyyy-MM-dd", "sv-SE": "yyyy-MM-dd", "th-TH": "d/M/yyyy", "tr-TR": "dd.MM.yyyy", "ur-PK": "dd/MM/yyyy", "id-ID": "dd/MM/yyyy", "uk-UA": "dd.MM.yyyy", "be-BY": "dd.MM.yyyy", "sl-SI": "d.M.yyyy", "et-EE": "d.MM.yyyy", "lv-LV": "yyyy.MM.dd.", "lt-LT": "yyyy.MM.dd", "fa-IR": "MM/dd/yyyy", "vi-VN": "dd/MM/yyyy", "hy-AM": "dd.MM.yyyy", "az-Latn-AZ": "dd.MM.yyyy", "eu-ES": "yyyy/MM/dd", "mk-MK": "dd.MM.yyyy", "af-ZA": "yyyy/MM/dd", "ka-GE": "dd.MM.yyyy", "fo-FO": "dd-MM-yyyy", "hi-IN": "dd-MM-yyyy", "ms-MY": "dd/MM/yyyy", "kk-KZ": "dd.MM.yyyy", "ky-KG": "dd.MM.yy", "sw-KE": "M/d/yyyy", "uz-Latn-UZ": "dd/MM yyyy", "tt-RU": "dd.MM.yyyy", "pa-IN": "dd-MM-yy", "gu-IN": "dd-MM-yy", "ta-IN": "dd-MM-yyyy", "te-IN": "dd-MM-yy", "kn-IN": "dd-MM-yy", "mr-IN": "dd-MM-yyyy", "sa-IN": "dd-MM-yyyy", "mn-MN": "yy.MM.dd", "gl-ES": "dd/MM/yy", "kok-IN": "dd-MM-yyyy", "syr-SY": "dd/MM/yyyy", "dv-MV": "dd/MM/yy", "ar-IQ": "dd/MM/yyyy", "zh-CN": "yyyy/M/d", "de-CH": "dd.MM.yyyy", "en-GB": "dd/MM/yyyy", "es-MX": "dd/MM/yyyy", "fr-BE": "d/MM/yyyy", "it-CH": "dd.MM.yyyy", "nl-BE": "d/MM/yyyy", "nn-NO": "dd.MM.yyyy", "pt-PT": "dd-MM-yyyy", "sr-Latn-CS": "d.M.yyyy", "sv-FI": "d.M.yyyy", "az-Cyrl-AZ": "dd.MM.yyyy", "ms-BN": "dd/MM/yyyy", "uz-Cyrl-UZ": "dd.MM.yyyy", "ar-EG": "dd/MM/yyyy", "zh-HK": "d/M/yyyy", "de-AT": "dd.MM.yyyy", "en-AU": "d/MM/yyyy", "es-ES": "dd/MM/yyyy", "fr-CA": "yyyy-MM-dd", "sr-Cyrl-CS": "d.M.yyyy", "ar-LY": "dd/MM/yyyy", "zh-SG": "d/M/yyyy", "de-LU": "dd.MM.yyyy", "en-CA": "dd/MM/yyyy", "es-GT": "dd/MM/yyyy", "fr-CH": "dd.MM.yyyy", "ar-DZ": "dd-MM-yyyy", "zh-MO": "d/M/yyyy", "de-LI": "dd.MM.yyyy", "en-NZ": "d/MM/yyyy", "es-CR": "dd/MM/yyyy", "fr-LU": "dd/MM/yyyy", "ar-MA": "dd-MM-yyyy", "en-IE": "dd/MM/yyyy", "es-PA": "MM/dd/yyyy", "fr-MC": "dd/MM/yyyy", "ar-TN": "dd-MM-yyyy", "en-ZA": "yyyy/MM/dd", "es-DO": "dd/MM/yyyy", "ar-OM": "dd/MM/yyyy", "en-JM": "dd/MM/yyyy", "es-VE": "dd/MM/yyyy", "ar-YE": "dd/MM/yyyy", "en-029": "MM/dd/yyyy", "es-CO": "dd/MM/yyyy", "ar-SY": "dd/MM/yyyy", "en-BZ": "dd/MM/yyyy", "es-PE": "dd/MM/yyyy", "ar-JO": "dd/MM/yyyy", "en-TT": "dd/MM/yyyy", "es-AR": "dd/MM/yyyy", "ar-LB": "dd/MM/yyyy", "en-ZW": "M/d/yyyy", "es-EC": "dd/MM/yyyy", "ar-KW": "dd/MM/yyyy", "en-PH": "M/d/yyyy", "es-CL": "dd-MM-yyyy", "ar-AE": "dd/MM/yyyy", "es-UY": "dd/MM/yyyy", "ar-BH": "dd/MM/yyyy", "es-PY": "dd/MM/yyyy", "ar-QA": "dd/MM/yyyy", "es-BO": "dd/MM/yyyy", "es-SV": "dd/MM/yyyy", "es-HN": "dd/MM/yyyy", "es-NI": "dd/MM/yyyy", "es-PR": "dd/MM/yyyy", "am-ET": "d/M/yyyy", "tzm-Latn-DZ": "dd-MM-yyyy", "iu-Latn-CA": "d/MM/yyyy", "sma-NO": "dd.MM.yyyy", "mn-Mong-CN": "yyyy/M/d", "gd-GB": "dd/MM/yyyy", "en-MY": "d/M/yyyy", "prs-AF": "dd/MM/yy", "bn-BD": "dd-MM-yy", "wo-SN": "dd/MM/yyyy", "rw-RW": "M/d/yyyy", "qut-GT": "dd/MM/yyyy", "sah-RU": "MM.dd.yyyy", "gsw-FR": "dd/MM/yyyy", "co-FR": "dd/MM/yyyy", "oc-FR": "dd/MM/yyyy", "mi-NZ": "dd/MM/yyyy", "ga-IE": "dd/MM/yyyy", "se-SE": "yyyy-MM-dd", "br-FR": "dd/MM/yyyy", "smn-FI": "d.M.yyyy", "moh-CA": "M/d/yyyy", "arn-CL": "dd-MM-yyyy", "ii-CN": "yyyy/M/d", "dsb-DE": "d. M. yyyy", "ig-NG": "d/M/yyyy", "kl-GL": "dd-MM-yyyy", "lb-LU": "dd/MM/yyyy", "ba-RU": "dd.MM.yy", "nso-ZA": "yyyy/MM/dd", "quz-BO": "dd/MM/yyyy", "yo-NG": "d/M/yyyy", "ha-Latn-NG": "d/M/yyyy", "fil-PH": "M/d/yyyy", "ps-AF": "dd/MM/yy", "fy-NL": "d-M-yyyy", "ne-NP": "M/d/yyyy", "se-NO": "dd.MM.yyyy", "iu-Cans-CA": "d/M/yyyy", "sr-Latn-RS": "d.M.yyyy", "si-LK": "yyyy-MM-dd", "sr-Cyrl-RS": "d.M.yyyy", "lo-LA": "dd/MM/yyyy", "km-KH": "yyyy-MM-dd", "cy-GB": "dd/MM/yyyy", "bo-CN": "yyyy/M/d", "sms-FI": "d.M.yyyy", "as-IN": "dd-MM-yyyy", "ml-IN": "dd-MM-yy", "en-IN": "dd-MM-yyyy", "or-IN": "dd-MM-yy", "bn-IN": "dd-MM-yy", "tk-TM": "dd.MM.yy", "bs-Latn-BA": "d.M.yyyy", "mt-MT": "dd/MM/yyyy", "sr-Cyrl-ME": "d.M.yyyy", "se-FI": "d.M.yyyy", "zu-ZA": "yyyy/MM/dd", "xh-ZA": "yyyy/MM/dd", "tn-ZA": "yyyy/MM/dd", "hsb-DE": "d. M. yyyy", "bs-Cyrl-BA": "d.M.yyyy", "tg-Cyrl-TJ": "dd.MM.yy", "sr-Latn-BA": "d.M.yyyy", "smj-NO": "dd.MM.yyyy", "rm-CH": "dd/MM/yyyy", "smj-SE": "yyyy-MM-dd", "quz-EC": "dd/MM/yyyy", "quz-PE": "dd/MM/yyyy", "hr-BA": "d.M.yyyy.", "sr-Latn-ME": "d.M.yyyy", "sma-SE": "yyyy-MM-dd", "en-SG": "d/M/yyyy", "ug-CN": "yyyy-M-d", "sr-Cyrl-BA": "d.M.yyyy", "es-US": "M/d/yyyy" };

            var date1;
            if (isDateInLocalFormat)
            {
                date1 = new Date(date);
            }
            else
            {
                date1 = Date.fromISO(date);
            }

            var l = (PDF.Config.locale)
                        ? PDF.Config.locale
                        : (navigator.language ? navigator.language : navigator['userLanguage']),
                y = date1.getFullYear(), m = date1.getMonth() + 1, d = date1.getDate(),
                df = "MM/dd/yyyy";

            for (var locale in f)
            {
                // get the date format based on passed in locale
                if (locale.toLowerCase() === l.toLowerCase())
                {
                    df = f[locale];
                    break;
                }
            }

            function z(s)
            {
                s = '' + s;
                return s.length > 1 ? s : '0' + s;
            }

            df = df.replace(/yyyy/, y);
            df = df.replace(/yy/, String(y).substr(2));
            df = df.replace(/MM/, z(m));
            df = df.replace(/M/, m);
            df = df.replace(/dd/, z(d));
            df = df.replace(/d/, d);
            return df;
        },

        DaysToToday: function (date)
        {
            if (date == null)
            {
                return 0;
            }
            var oneDay = 1000 * 60 * 60 * 24;
            // Convert both dates to milliseconds
            var date1 = new Date.fromISO(date);
            var date2 = new Date();
            // Calculate the difference in milliseconds
            var difference = date1 - date2;

            // Convert back to days and return
            var numberOfDays = Math.round(difference / oneDay);
            return numberOfDays;
        },

        AddYearToDateOffset: function (date, yearsToAdd)
        {
            var newDate = new Date();
            newDate.setFullYear(date.getFullYear() + yearsToAdd, date.getMonth(), date.getDate());
            return PDF.Utils.FormatDate(newDate, true);
        },

        IsDateInBetween: function (fromDate, toDate, dateCheck)
        {
            var d1 = fromDate.split("/");
            var d2 = toDate.split("/");
            var c = dateCheck.split("/");

            var from = new Date(d1[2], d1[1] - 1, d1[0]);  // -1 because months are from 0 to 11
            var to = new Date(d2[2], d2[1] - 1, d2[0]);
            var check = new Date(c[2], c[1] - 1, c[0]);

            var retVal = ((check <= to && check >= from));
            return retVal;
        },

        LogMessage: function (message, messageTypeId, url, line, scenarioId)
        {
            if (PDF.Config.clientErrorLoggingEnabled && PDF.LogLevels[PDF.Config.clientErrorLoggingLevel.toUpperCase()] >= messageTypeId) // Int to string comapre
            {
                var messageTypeName;
                for (var property in PDF.LogLevels)
                {
                    if (PDF.LogLevels[property] === messageTypeId)
                    {
                        messageTypeName = property;
                        break;
                    }
                }

                try
                {
                    $.ajax(
                    {
                        url: PDF.Config.errorLogApiUrl,
                        type: 'POST',
                        data:
                        {
                            errorMessage: message,
                            messageType: messageTypeName,
                            errorLine: line,
                            queryString: document.location.search,
                            url: url,
                            referrer: document.referrer,
                            userAgent: navigator.userAgent,
                            scenarioId: scenarioId || ''
                        }
                    });
                }
                catch (e) { }
            }
        },

        LoadSharedTemplates: function (templates, el, options)
        {
            var promiseArray = [];
            var $el = $(el);
            ko.utils.arrayForEach(templates, function (name)
            {
                promiseArray.push(PDF.TemplateManager.GetPromise(name, PDF.Utils.GetSharedTemplateURL)
                .done(function (template)
                {
                    $el.append(template);
                })
               .fail(function (xhr)
               {
                   PDF.Utils.TemplateLoadingErrorHandler(xhr, $el, name, options);
               }));
            });

            return PDF.Promise.join(promiseArray)
        },

        TemplateLoadingErrorHandler: function (xhr, $el, componentName, options)
        {
            // Chrome unload fix: chrome will fire ajax error when unload page. 
            if (xhr && xhr.status === 0)
            {
                return;
            }

            var errorMessage = xhr.hasOwnProperty('error') ? xhr.error().responseText : xhr;
            PDF.Utils.LogMessage(errorMessage, PDF.LogLevels.ERROR, document.URL, 0);

            // error occurs when loading PDF component
            var errorHtml = "";
            if (options.enableErrorHtml)
            {
                // In future: get more information from error context and show different error
                var errorTemplate = PDF.GetConfig("componentLoadingErrorHtmlTemplate", componentName);
                if (errorTemplate)
                {
                    var errorText1 = PDF.GetConfig("componentLoadingErrorText1", componentName);
                    var errorText2 = PDF.GetConfig("componentLoadingErrorText2", componentName);

                    errorHtml = errorTemplate.format(errorText1, errorText2);
                }
            }
            $el.html(errorHtml);


            PDF.Events.Trigger("ComponentLoadingError", { componentName: componentName, arguments: arguments });

            // For debugging
            PDF.Utils.Alert(arguments);
        },

        LoadTabEvent: function (event) {
            var key;
            var isShift;
            if (window.event) {
                key = window.event.keyCode;
                isShift = window.event.shiftKey ? true : false;
            } else {
                key = ev.which;
                isShift = ev.shiftKey ? true : false;
            }
            if ((key == 16) && (isShift))
            { shiftpressed = "TRUE"; }

        },

        LoadDynamicTemplate: function (templateName)
        {
            var templateSourceKey = PDF.SharedTemplateRepository.GetTemplateFileNameByTemplateName(templateName);
            var promise = PDF.Promise();

            PDF.TemplateManager.GetPromise(templateSourceKey, PDF.Utils.GetSharedTemplateURL)
            .done(function (templateDocument)
            {
                // TODO: check if SourceKey-document is appended
                var templateElement = document.getElementById(templateName);
                if (!templateElement)
                {
                    // TODO: native dom api
                    $('body').append(templateDocument);
                    templateElement = document.getElementById(templateName);
                }

                promise.resolve(templateElement);
            })
            .fail(function (xhr)
            {
                throw new Error("Cannot find template with ID " + templateName);
            });

            return promise;
        },

        IsNullOrUndefined: function (item)
        {
            return (typeof item === undefined || item == null);
        },

        GetSharedTemplateURL: function (templateName)
        {
            return (PDF.Config.sharedTemplateUrl || "content/pdf/template/subtemplate/{0}.html").format(templateName);
        },

        GetContextUrl: function (baseURL, postUrl, defaultPostUrl)
        {
            baseURL = baseURL || '';
            postUrl = postUrl || defaultPostUrl;
            return baseURL + postUrl;
        },

        GetContextUrlFn: function (componentName, contextBaseUrlKey)
        {
            if (!componentName)
            {
                return;
            }

            contextBaseUrlKey = contextBaseUrlKey || "baseUrl";

            var baseURL = PDF.GetConfig(contextBaseUrlKey, componentName) || '';

            return function (contextUrlKey)
            {
                var contextUrl = PDF.GetConfig(contextUrlKey, componentName) || contextUrlKey;
                return baseURL + contextUrl;
            };
        },

        LoadJavaScript: PDF.ResourceLoader.LoadJavaScript,

        // Wrap jQuery.ajax with credentials flag
        AjaxWithCredential: function (url, options)
        {
            // shit parameters if url is omitted
            if (typeof url === "object")
            {
                options = url;
                url = undefined;
            }
            options = options || {};

            // set allow credential flag
            options.xhrFields = { withCredentials: true };

            if (PDF.Config.locale)
            {
                options = options || {};
                options.data = options.data || {};
                if (!options.data.hasOwnProperty('locale') && typeof(options.data) == 'object')
                {
                    options.data['locale'] = PDF.Config.locale;
                }
            }

            return $.ajax(url, options);
        }
    };

    //Get, Post, Delete with credential flags 
    $.each(['Get', 'Post', 'Delete'], function (i, method)
    {
        PDF.Utils[method] = function (url, data, callback, dataType)
        {
            // shift parameters if data was omitted
            if (typeof data === 'function')
            {
                dataType = dataType || callback;
                callback = data;
                data = undefined;
            }

            return PDF.Utils.AjaxWithCredential({
                url: url,
                type: method,
                dataType: dataType,
                data: data,
                success: callback
            });
        };
    });

    // Get JSON with credentials flag
    PDF.Utils.GetJSON = function (url, data, callback)
    {
        return PDF.Utils.Get(url, data, callback, "json");
    };

})(window, document, jQuery, ko, PDF);
