(function (window, $, ko, PDF, undefined)
{
    "use strict";

    var componentIdentifier = "RelatedContent";

    PDF.ViewModels[componentIdentifier] = function (promise, el, params)
    {
        var viewModel = function ($)
        {
            var self = Object.create(PDF.BaseViewModel);
            self.relatedContentList = ko.observableArray();
            self.moreUrl = ko.observable(params.MoreUrl);
            self.componentKey = ko.observable(params.ComponentKey);
            self.getDataContextUrlFn = PDF.Utils.GetContextUrlFn(componentIdentifier, 'dataContextBaseUrl');

            var latestRequestSequence = 0;

            self.LoadRelatedContent = function ()
            {
                var url = self.getDataContextUrlFn('api/pdf/Search/GetRelatedContent');
                latestRequestSequence = latestRequestSequence + 1;
                var currentRequestSequence = latestRequestSequence;
                return $.getJSON(url, { query: params.Query, filter: params.Filter == null ? '' : params.Filter, culture: params.Culture, relatedTo: params.RelatedTo, componentKey: params.ComponentKey }, function (data)
                {
                    promise.done(function ()
                    {
                        if (currentRequestSequence >= latestRequestSequence)
                        {
                            self.relatedContentList(data);
                        }
                    });
                });
            };

            PDF.Events.Bind("changeContext", function (e, data)
            {
                params.Query = data.Query;
                params.filter = data.Filter;
                self.moreUrl(data.MoreUrl);
                self.LoadRelatedContent();
            });

            PDF.Events.Bind("FromCommunity.GetMoreCommunityUrl", function (e, data)
            {
                self.moreUrl(data.MoreUrl);
            });

            return self;
        }($);

        promise.await([
            viewModel.ExtendUIStrings("PDF.RelatedContent.UIStrings." + params.ComponentKey),
            viewModel.ExtendContentModel("PDF.RelatedContent.Configuration." + params.ComponentKey, "Config"),
            viewModel.ExtendBIContentModel("BiRelatedContent", "BIKeys"),
            viewModel.LoadRelatedContent()
        ], viewModel);
    };
})(window, jQuery, ko, PDF);