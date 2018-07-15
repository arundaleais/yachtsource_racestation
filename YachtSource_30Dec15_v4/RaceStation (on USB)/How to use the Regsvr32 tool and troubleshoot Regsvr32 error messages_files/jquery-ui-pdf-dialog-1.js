/* PDF UI Dialog - v1.10.0 - 2013-02-07 */
/* PDF UI: a customized version of jquery ui dialog with supporting Windows 8 styles */

/* Dependencies: jQuery 1.6+,  jQuery UI 1.10+, PDF theme css*/

(function ()
{
    // ensure dependencies loaded
    window.jQueryLibrarySrc = window.jQueryLibrarySrc || '//ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.1.min.js';
    window.jQueryUILibrarySrc = window.jQueryUILibrarySrc || '//ajax.aspnetcdn.com/ajax/jquery.ui/1.10.0/jquery-ui.min.js';

    EnsureJquery();

    function LoadJavaScript(srcPath, callback)
    {
        var script = document.createElement('script');
        script.type = 'text/javascript';
        script.src = srcPath;
        if (typeof callback === 'function')
        {
            if (typeof (script.onload) != 'undefined')
            {
                script.onload = callback;
            } else
            {
                script.onreadystatechange = function (e)
                {
                    if (this.readyState == 'complete' || this.readyState == 'loaded')
                        callback();
                };
            }
        }
        document.getElementsByTagName("head")[0].appendChild(script);
    }

    function EnsureJquery()
    {
        if (typeof (jQuery) === 'undefined')
        {
            LoadJavaScript(window.jQueryLibrarySrc, EnsureJqueryUI);
        }
        else
        {
            EnsureJqueryUI();
        }
    };

    function EnsureJqueryUI()
    {
        $.ui = $.ui || {};
        if (typeof ($.ui.version) === 'undefined')
        {
            LoadJavaScript(window.jQueryUILibrarySrc, Load_pdfDialog);

        } else
        {
            Load_pdfDialog();
        }
    };

    function Load_pdfDialog()
    {
        (function ($, undefined)
        {
            var sizeRelatedOptions = {
                buttons: true,
                height: true,
                maxHeight: true,
                maxWidth: true,
                minHeight: true,
                minWidth: true,
                width: true
            },
	        resizableRelatedOptions = {
	            maxHeight: true,
	            maxWidth: true,
	            minHeight: true,
	            minWidth: true
	        };

            $.widget("ui.pdfDialog", {
                version: "1.10.0",
                options: {
                    appendTo: "body",
                    autoOpen: true,
                    buttons: [],
                    closeOnEscape: true,
                    closeText: "close",
                    closeOnOverlayClick: true,
                    dialogClass: "",
                    draggable: false,
                    hide: null,
                    height: "auto",
                    maxHeight: null,
                    maxWidth: null,
                    minHeight: 150,
                    minWidth: 150,
                    modal: true,
                    defaultFocusElement: null,
                    defaultFocusCloseButton: false,
                    position: {
                        my: "center",
                        at: "center",
                        of: window,
                        collision: "fit",
                        // Ensure the titlebar is always visible
                        using: function (pos)
                        {
                            var topOffset = $(this).css(pos).offset().top;
                            if (topOffset < 0)
                            {
                                $(this).css("top", pos.top - topOffset);
                            }
                        }
                    },
                    resizable: false,
                    show: null,
                    title: null,
                    width: "100%",
                    closeIconPath: "/content/pdf/images/ui-dialog-titlebar-close.png",
                    titleId: null,

                    // callbacks
                    beforeClose: null,
                    close: null,
                    drag: null,
                    dragStart: null,
                    dragStop: null,
                    focus: null,
                    open: null,
                    resize: null,
                    resizeStart: null,
                    resizeStop: null
                },

                _create: function ()
                {
                    this.originalCss = {
                        display: this.element[0].style.display,
                        width: this.element[0].style.width,
                        minHeight: this.element[0].style.minHeight,
                        maxHeight: this.element[0].style.maxHeight,
                        height: this.element[0].style.height
                    };
                    this.originalPosition = {
                        parent: this.element.parent(),
                        index: this.element.parent().children().index(this.element)
                    };
                    this.originalTitle = this.element.attr("title");
                    this.options.title = this.options.title || this.originalTitle;
                    this._createWrapper();

                    this.element
                        .show()
                        .removeAttr("title")
                        .addClass("ui-dialog-content ui-widget-content")
                        .appendTo(this.uiDialog);

                    this._createTitlebar();
                    this._createButtonPane();

                    if (this.options.draggable && $.fn.draggable)
                    {
                        this._makeDraggable();
                    }
                    if (this.options.resizable && $.fn.resizable)
                    {
                        this._makeResizable();
                    }

                    this._isOpen = false;
                },

                _init: function ()
                {
                    if (this.options.autoOpen)
                    {
                        this.open();
                    }
                },

                _appendTo: function ()
                {
                    var element = this.options.appendTo;
                    if (element && (element.jquery || element.nodeType))
                    {
                        return $(element);
                    }
                    return this.document.find(element || "body").eq(0);
                },

                _destroy: function ()
                {
                    var next,
                        originalPosition = this.originalPosition;

                    this._destroyOverlay();

                    this.element
                        .removeUniqueId()
                        .removeClass("ui-dialog-content ui-widget-content")
                        .css(this.originalCss)
                        .detach();

                    this.uiDialog.stop(true, true).remove();

                    if (this.originalTitle)
                    {
                        this.element.attr("title", this.originalTitle);
                    }

                    next = originalPosition.parent.children().eq(originalPosition.index);
                    if (next.length && next[0] !== this.element[0])
                    {
                        next.before(this.element);
                    } else
                    {
                        originalPosition.parent.append(this.element);
                    }
                },

                widget: function ()
                {
                    return this.uiDialog;
                },

                disable: $.noop,
                enable: $.noop,

                close: function (event)
                {
                    var that = this;

                    if (!this._isOpen || this._trigger("beforeClose", event) === false)
                    {
                        return;
                    }

                    this._isOpen = false;
                    this._destroyOverlay();

                    if (!this.opener.filter(":focusable").focus().length)
                    {
                        $(this.document[0].activeElement).blur();
                    }

                    this._hide(this.uiDialog, this.options.hide, function ()
                    {
                        that._trigger("close", event);
                    });
                },

                isOpen: function ()
                {
                    return this._isOpen;
                },

                moveToTop: function ()
                {
                    this._moveToTop();
                },

                _moveToTop: function (event, silent)
                {
                    var moved = !!this.uiDialog.nextAll(":visible").insertBefore(this.uiDialog).length;
                    if (moved && !silent)
                    {
                        this._trigger("focus", event);
                    }
                    return moved;
                },

                open: function ()
                {
                    if (this._isOpen)
                    {
                        if (this._moveToTop())
                        {
                            this._focusTabbable();
                        }
                        return;
                    }

                    this.opener = $(this.document[0].activeElement);

                    this._size();
                    this._position();
                    this._createOverlay();
                    this._moveToTop(null, true);
                    this._show(this.uiDialog, this.options.show);

                    this._focusTabbable();

                    this._isOpen = true;
                    this._trigger("open");
                    this._trigger("focus");
                },

                _focusTabbable: function ()
                {
                    var hasFocus = {};

                    // First precedence is given to specific element passed in options
                    // Second precedence is given to close button if passes explicitly
                    // Continue to focus on first tabable element other wise.

                    if (this.options.defaultFocusElement !== null)
                    {
                        var element = this.element.find(this.options.defaultFocusElement);
                        if (element.length)
                        {
                            hasFocus = element;
                        }
                    }
                    if (!hasFocus.length && this.options.defaultFocusCloseButton)
                    {
                        hasFocus = $(this.uiDialogTitlebarClose);
                    }
                    if (!hasFocus.length)
                    {
                        hasFocus = this.element.find("[autofocus]");
                    }
                    if (!hasFocus.length)
                    {
                        hasFocus = this.element.find(":tabbable");
                    }
                    if (!hasFocus.length)
                    {
                        hasFocus = this.uiDialogButtonPane.find(":tabbable");
                    }
                    if (!hasFocus.length)
                    {
                        if (this.uiDialogTitlebarClose !== undefined)
                        {
                            hasFocus = this.uiDialogTitlebarClose.filter(":tabbable");
                        }
                    }
                    if (!hasFocus.length)
                    {
                        hasFocus = this.uiDialog;
                    }

                    hasFocus.eq(0).focus();
                },

                _keepFocus: function (event)
                {
                    function checkFocus()
                    {
                        var activeElement = this.document[0].activeElement,
                            isActive = this.uiDialog[0] === activeElement ||
                                $.contains(this.uiDialog[0], activeElement);
                        if (!isActive)
                        {
                            this._focusTabbable();
                        }
                    }
                    event.preventDefault();
                    checkFocus.call(this);
                    this._delay(checkFocus);
                },

                _createWrapper: function ()
                {
                    this.uiDialog = $("<div>")
                        .addClass("ui-dialog ui-widget ui-widget-content ui-corner-all ui-front " +
                            this.options.dialogClass)
                        .hide()
                        .attr({
                            tabIndex: -1,
                            role: "dialog"
                        })
                        .appendTo(this._appendTo());

                    this._on(this.uiDialog, {
                        keydown: function (event)
                        {
                            if (this.options.closeOnEscape && !event.isDefaultPrevented() && event.keyCode &&
                                    event.keyCode === $.ui.keyCode.ESCAPE)
                            {
                                event.preventDefault();
                                this.close(event);
                                return;
                            }

                            if (event.keyCode !== $.ui.keyCode.TAB)
                            {
                                return;
                            }
                            var tabbables = this.uiDialog.find(":tabbable"),
                                first = tabbables.filter(":first"),
                                last = tabbables.filter(":last");

                            if ((event.target === last[0] || event.target === this.uiDialog[0]) && !event.shiftKey)
                            {
                                first.focus(1);
                                event.preventDefault();
                            } else if ((event.target === first[0] || event.target === this.uiDialog[0]) && event.shiftKey)
                            {
                                last.focus(1);
                                event.preventDefault();
                            }
                        },
                        mousedown: function (event)
                        {
                            if (this._moveToTop(event))
                            {
                                this._focusTabbable();
                            }
                        }
                    });

                    if (!this.element.find("[aria-describedby]").length)
                    {
                        this.uiDialog.attr({
                            "aria-describedby": this.element.uniqueId().attr("id")
                        });
                    }
                },

                _createTitlebar: function ()
                {
                    var uiDialogTitle;

                    this.uiDialogTitlebar = $("<div>")
                        .addClass("ui-dialog-titlebar ui-widget-header ui-corner-all ui-helper-clearfix ui-front")
                        .prependTo(this.uiDialog);
                    this._on(this.uiDialogTitlebar, {
                        mousedown: function (event)
                        {
                            if (!$(event.target).closest(".ui-dialog-titlebar-close"))
                            {
                                this.uiDialog.focus();
                            }
                        }
                    });

                    this.uiDialogTitlebarClose = $("<button><img alt=\"" + this.options.closeText + "\" src=\"" + this.options.closeIconPath + "\" /></button>")
                        .attr('title', this.options.closeText)
                        .addClass("ui-dialog-titlebar-close ui-front")
                        .appendTo(this.uiDialogTitlebar);
                    this._on(this.uiDialogTitlebarClose, {
                        click: function (event)
                        {
                            event.preventDefault();
                            this.close(event);
                        }
                    });

                    if (this.options.title)
                    {
                        uiDialogTitle = $("<span>")
                            .uniqueId()
                            .addClass("ui-dialog-title")
                            .prependTo(this.uiDialogTitlebar);
                        this._title(uiDialogTitle);

                        this.uiDialog.attr({
                            "aria-labelledby": uiDialogTitle.attr("id")
                        });
                    }
                    else if (this.options.titleId)
                    {
                        this.uiDialog.attr({
                            "aria-labelledby": this.options.titleId
                        });
                    }
                },

                _title: function (title)
                {
                    if (!this.options.title)
                    {
                        title.html("&#160;");
                    }
                    title.text(this.options.title);
                },

                _createButtonPane: function ()
                {
                    this.uiDialogButtonPane = $("<div>")
                        .addClass("ui-dialog-buttonpane ui-widget-content ui-helper-clearfix");

                    this.uiButtonSet = $("<div>")
                        .addClass("ui-dialog-buttonset")
                        .appendTo(this.uiDialogButtonPane);

                    this._createButtons();
                },

                _createButtons: function ()
                {
                    var that = this,
                        buttons = this.options.buttons;

                    this.uiDialogButtonPane.remove();
                    this.uiButtonSet.empty();

                    if ($.isEmptyObject(buttons))
                    {
                        this.uiDialog.removeClass("ui-dialog-buttons");
                        return;
                    }

                    $.each(buttons, function (name, props)
                    {
                        var click, buttonOptions;
                        props = $.isFunction(props) ?
				{ click: props, text: name } :
                            props;
                        props = $.extend({ type: "button" }, props);
                        click = props.click;
                        props.click = function ()
                        {
                            click.apply(that.element[0], arguments);
                        };
                        buttonOptions = {
                            icons: props.icons,
                            text: props.showText
                        };
                        delete props.icons;
                        delete props.showText;
                        $("<button></button>", props)
                            .button(buttonOptions)
                            .appendTo(that.uiButtonSet);
                    });
                    this.uiDialog.addClass("ui-dialog-buttons");
                    this.uiDialogButtonPane.appendTo(this.uiDialog);
                },

                _makeDraggable: function ()
                {
                    var that = this,
                        options = this.options;

                    function filteredUi(ui)
                    {
                        return {
                            position: ui.position,
                            offset: ui.offset
                        };
                    }

                    this.uiDialog.draggable({
                        cancel: ".ui-dialog-content, .ui-dialog-titlebar-close",
                        handle: ".ui-dialog-titlebar",
                        containment: "document",
                        start: function (event, ui)
                        {
                            $(this).addClass("ui-dialog-dragging");
                            that._trigger("dragStart", event, filteredUi(ui));
                        },
                        drag: function (event, ui)
                        {
                            that._trigger("drag", event, filteredUi(ui));
                        },
                        stop: function (event, ui)
                        {
                            options.position = [
                                ui.position.left - that.document.scrollLeft(),
                                ui.position.top - that.document.scrollTop()
                            ];
                            $(this).removeClass("ui-dialog-dragging");
                            that._trigger("dragStop", event, filteredUi(ui));
                        }
                    });
                },

                _makeResizable: function ()
                {
                    var that = this,
                        options = this.options,
                        handles = options.resizable,
                        position = this.uiDialog.css("position"),
                        resizeHandles = typeof handles === "string" ?
                        handles :
                            "n,e,s,w,se,sw,ne,nw";

                    function filteredUi(ui)
                    {
                        return {
                            originalPosition: ui.originalPosition,
                            originalSize: ui.originalSize,
                            position: ui.position,
                            size: ui.size
                        };
                    }

                    this.uiDialog.resizable({
                        cancel: ".ui-dialog-content",
                        containment: "document",
                        alsoResize: this.element,
                        maxWidth: options.maxWidth,
                        maxHeight: options.maxHeight,
                        minWidth: options.minWidth,
                        minHeight: this._minHeight(),
                        handles: resizeHandles,
                        start: function (event, ui)
                        {
                            $(this).addClass("ui-dialog-resizing");
                            that._trigger("resizeStart", event, filteredUi(ui));
                        },
                        resize: function (event, ui)
                        {
                            that._trigger("resize", event, filteredUi(ui));
                        },
                        stop: function (event, ui)
                        {
                            options.height = $(this).height();
                            options.width = $(this).width();
                            $(this).removeClass("ui-dialog-resizing");
                            that._trigger("resizeStop", event, filteredUi(ui));
                        }
                    })
                    .css("position", position);
                },

                _minHeight: function ()
                {
                    var options = this.options;

                    return options.height === "auto" ?
                        options.minHeight :
                        Math.min(options.minHeight, options.height);
                },

                _position: function ()
                {
                    var isVisible = this.uiDialog.is(":visible");
                    if (!isVisible)
                    {
                        this.uiDialog.show();
                    }
                    this.uiDialog.position(this.options.position);
                    if (!isVisible)
                    {
                        this.uiDialog.hide();
                    }
                },

                _setOptions: function (options)
                {
                    var that = this,
                        resize = false,
                        resizableOptions = {};

                    $.each(options, function (key, value)
                    {
                        that._setOption(key, value);

                        if (key in sizeRelatedOptions)
                        {
                            resize = true;
                        }
                        if (key in resizableRelatedOptions)
                        {
                            resizableOptions[key] = value;
                        }
                    });

                    if (resize)
                    {
                        this._size();
                        this._position();
                    }
                    if (this.uiDialog.is(":data(ui-resizable)"))
                    {
                        this.uiDialog.resizable("option", resizableOptions);
                    }
                },

                _setOption: function (key, value)
                {
                    var isDraggable, isResizable,
                        uiDialog = this.uiDialog;

                    if (key === "dialogClass")
                    {
                        uiDialog
                            .removeClass(this.options.dialogClass)
                            .addClass(value);
                    }

                    if (key === "disabled")
                    {
                        return;
                    }

                    this._super(key, value);

                    if (key === "appendTo")
                    {
                        this.uiDialog.appendTo(this._appendTo());
                    }

                    if (key === "buttons")
                    {
                        this._createButtons();
                    }

                    if (key === "closeText")
                    {
                        this.uiDialogTitlebarClose.button({
                            label: "" + value
                        });
                    }

                    if (key === "draggable")
                    {
                        isDraggable = uiDialog.is(":data(ui-draggable)");
                        if (isDraggable && !value)
                        {
                            uiDialog.draggable("destroy");
                        }

                        if (!isDraggable && value)
                        {
                            this._makeDraggable();
                        }
                    }

                    if (key === "position")
                    {
                        this._position();
                    }

                    if (key === "resizable")
                    {
                        isResizable = uiDialog.is(":data(ui-resizable)");
                        if (isResizable && !value)
                        {
                            uiDialog.resizable("destroy");
                        }

                        if (isResizable && typeof value === "string")
                        {
                            uiDialog.resizable("option", "handles", value);
                        }

                        if (!isResizable && value !== false)
                        {
                            this._makeResizable();
                        }
                    }

                    if (key === "title")
                    {
                        this._title(this.uiDialogTitlebar.find(".ui-dialog-title"));
                    }
                },

                _size: function ()
                {
                    var nonContentHeight, minContentHeight, maxContentHeight,
                        options = this.options;

                    this.element.show().css({
                        width: "auto",
                        minHeight: 0,
                        maxHeight: "none",
                        height: 0
                    });

                    if (options.minWidth > options.width)
                    {
                        options.width = options.minWidth;
                    }

                    nonContentHeight = this.uiDialog.css({
                        height: "auto",
                        width: options.width
                    })
                        .outerHeight();
                    minContentHeight = Math.max(0, options.minHeight - nonContentHeight);
                    maxContentHeight = typeof options.maxHeight === "number" ?
                        Math.max(0, options.maxHeight - nonContentHeight) :
                        "none";

                    if (options.height === "auto")
                    {
                        this.element.css({
                            minHeight: minContentHeight,
                            maxHeight: maxContentHeight,
                            height: "auto"
                        });
                    } else
                    {
                        this.element.height(Math.max(0, options.height - nonContentHeight));
                    }

                    if (this.uiDialog.is(":data(ui-resizable)"))
                    {
                        this.uiDialog.resizable("option", "minHeight", this._minHeight());
                    }
                },

                _createOverlay: function ()
                {
                    if (!this.options.modal)
                    {
                        return;
                    }

                    if (!$.ui.pdfDialog.overlayInstances)
                    {
                        this._delay(function ()
                        {
                            if ($.ui.pdfDialog.overlayInstances)
                            {
                                this._on(this.document, {
                                    focusin: function (event)
                                    {
                                        if (!$(event.target).closest(".ui-dialog").length)
                                        {
                                            event.preventDefault();
                                            var dialog = $(".ui-dialog:visible:last .ui-dialog-content").data("ui-dialog");
                                            if (dialog && dialog._focusTabbable === "function")
                                            {
                                                dialog._focusTabbable();
                                            }
                                        }
                                    }
                                });
                            }
                        });
                    }

                    var that = this;
                    this.overlay = $("<div>")
                        .addClass("ui-widget-overlay ui-front")
                        .appendTo(this.document[0].body);
                    this._on(this.overlay, {
                        mousedown: that.options.closeOnOverlayClick ? "close" : "_keepFocus"
                    });
                    $.ui.pdfDialog.overlayInstances++;
                },

                _destroyOverlay: function ()
                {
                    if (!this.options.modal)
                    {
                        return;
                    }

                    $.ui.pdfDialog.overlayInstances--;
                    if (!$.ui.pdfDialog.overlayInstances)
                    {
                        this._off(this.document, "focusin");
                    }
                    this.overlay.remove();
                }
            });

            $.ui.pdfDialog.overlayInstances = 0;
        })(jQuery);
    }
})();