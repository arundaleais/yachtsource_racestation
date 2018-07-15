(function ($)
{
    var sections = $('.accordion .ac-detail');
    sections.hide();
   
    $(document).on("click", ".accordion .ac-title > a", function() {
        sections = $('.accordion .ac-detail');
        for (var i = 0; i < sections.length; i++) {
            if (sections[i] != $(this).parent().next()[0]) {
                $(sections[i]).slideUp();
            }
        }

        var image_icon = $('.accordion .ac-title > a > img').attr('src');
        if (!image_icon || image_icon.indexOf('/icon_collapse.png') || image_icon.indexOf('/icon_expand.png'))
        {
            $('.accordion .ac-title > a > img').attr('src', PDF.Config.imageUrl + 'icon_expand.png');
        }

        $(this).parent().next().slideDown();
        image_icon = $(this).children('img').attr('src');
        if (!image_icon || image_icon.indexOf('/icon_collapse.png') || image_icon.indexOf('/icon_expand.png'))
        {
            $(this).children('img').attr('src', PDF.Config.imageUrl + 'icon_collapse.png');
        }
        return false;
    });
})(jQuery);

// Webtrend Listener script
(function ($)
{
    MS.Support.WebTrendMetaDataTrigger = function (config)
    {
        var settings = config;

        $(document).ready(function ()
        {
            checkMetdata();
        });

        var retryCount = 0;

        checkMetdata = function ()
        {
            try
            {
                var webTrendIdMetadataValue = $('meta[name="' + settings.WebTrendIdMetadataName + '"]').attr('content');
                if (webTrendIdMetadataValue)
                {
                    PDF.Events.Trigger("WebTrendMetadataTrigger", { variant: settings.Config[webTrendIdMetadataValue.toUpperCase()] || 'control' });
                }
                else
                {
                    var webTrendErrorMetadataValue = $('meta[name="' + settings.WebTrendErrorMetadataName + '"]').attr('content');
                    retryCount++;
                    if ((retryCount == (settings.RetryCount || 3)) || webTrendErrorMetadataValue)
                    {
                        PDF.Events.Trigger("WebTrendMetadataTrigger", { variant: settings.defaultVariant || 'control' });
                    }
                    else
                    {
                        window.setTimeout(function ()
                        {
                            checkMetdata();
                        }, settings.RefreshInterval || 100);
                    }
                }
            }
            catch (e)
            {
                PDF.Events.Trigger("WebTrendMetadataTrigger", { variant: settings.defaultVariant || 'control' });
            }
        };
    };
})(jQuery);