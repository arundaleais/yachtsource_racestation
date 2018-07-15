(function (window, $, ko, PDF, undefined)
{
    "use strict";

    var componentIdentifier = "ProductSelector";

    PDF.ViewModels[componentIdentifier] = function (promise, el, params)
    {
        var viewModel = function ($)
        {
            var self = Object.create(PDF.BaseViewModel);
            self.currentProductId = ko.observable();
            self.productSelector = ko.observable(params.productSelector);

            //start : PDF Product selector for home page code 
            self.currentProduct = function ()
            {
                var product = null;

                if (self.currentProductId())
                {
                    product = PDF.Utils.ArrayFirst(self.Products(), function (product)
                    {
                        return product.Value.ProductId === self.currentProductId();
                    });
                }

                return product;
            }

            self.GetProductUrl = function (data)
            {
                if (data.Value.SelfhelpUrl)
                {
                    if (data.Value.SelfhelpUrl != "")
                    {
                        PDF.Events.Trigger("LoadProductPage", { PID: data.Value.ProductId, PName: data.Value.Key, URL: data.Value.SelfhelpUrl });
                    }
                    else
                    {
                        PDF.Events.Trigger("LoadProductPage", { PID: data.Value.ProductId, PName: data.Value.Key, URL: null, ProductObj: data });
                    }
                }
                else
                {
                    PDF.Events.Trigger("LoadProductPage", { PID: data.Value.ProductId, PName: data.Value.Key, URL: null, ProductObj: data });
                }
            }

            self.UpdateAllProductsURL = function (data)
            {
                var parentProduct;
                ko.utils.arrayForEach(self.AllProductsList, function (parent)
                {
                    parentProduct = parent;
                    ko.utils.arrayForEach(parent.Products, function (child)
                    {
                        if (data.Value.ProductId === parentProduct.GsaId && data.Value.ProductName == null)
                        {
                            data.Value.ProductName = parentProduct.DisplayName;
                            data.Value.SelfhelpUrl = parentProduct.SelfHelpUrl;
                            data.Value.CommunityUrl = parentProduct.CommunityUrl;
                            data.Value.LivehelpUrl = parentProduct.LiveHelpUrl;
                            data.Value.RegisterUrl = parentProduct.RegisterUrl;
                            data.Value.RenewUrl = parentProduct.RenewUrl;
                        }

                        //As we have duplicate GsaIds in AllProducts.xml, i am picking up first one
                        if (data.Value.ProductId === child.GsaId && data.Value.ProductName == null)
                        {
                            data.Value.ProductName = child.DisplayName;
                            data.Value.SelfhelpUrl = child.SelfHelpUrl;
                            data.Value.CommunityUrl = child.CommunityUrl;
                            data.Value.LivehelpUrl = child.LiveHelpUrl;
                            data.Value.RegisterUrl = child.RegisterUrl;
                            data.Value.RenewUrl = child.RenewUrl;
                        }
                        ko.utils.arrayForEach(data.Value.Productlist, function (product)
                        {

                            if (product.ProductName == null)
                            {
                                if (child.GsaId != null && child.GsaId === product.ProductId)
                                {
                                    product.SelfhelpUrl = child.SelfHelpUrl;
                                    product.ProductName = child.DisplayName;
                                    product.CommunityUrl = child.CommunityUrl;
                                    product.LivehelpUrl = child.LiveHelpUrl;
                                    product.RegisterUrl = child.RegisterUrl;
                                    product.RenewUrl = child.RenewUrl;
                                }
                            }
                        });
                    });
                });
            }

            self.Products = function ()
            {
                if (!productlist)
                {
                    var productlist = self.ProductList.filter(function (product)
                    {
                        if (product.Value.hasOwnProperty('HideOnProductsTile') == false)
                        {
                            {
                                self.UpdateAllProductsURL(product);
                                return product.Value.hasOwnProperty('ProductId');
                            }
                        }
                        else
                        {
                            if (product.Value.HideOnProductsTile == null || product.Value.HideOnProductsTile.value == "false")
                            {
                                self.UpdateAllProductsURL(product);
                                return product.Value.hasOwnProperty('ProductId');
                            }
                            else
                            {
                                return;
                            }
                        }
                    });
                }

                if (productlist && PDF.Config.cdnServerHost) {
                    ko.utils.arrayForEach(productlist, function (product) {
                        var iconPath = product.Value.IconPath;
                        if (iconPath && iconPath.indexOf("//") != 0 && iconPath.indexOf("http") != 0) {
                            product.Value.IconPath = PDF.Config.cdnServerHost + iconPath.toLowerCase();
                        }
                    })
                }

                return productlist;
            };

            self.currentProductId.subscribe(function ()
            {
                var product = self.currentProduct();
                if (product)
                {
                    self.currentProduct(product);
                    self.currentProductId(product.Value.ProductId);

                }

            });

            self.ProductList = null;

            self.LoadProduct = function ()
            {
                var getUrl = PDF.Utils.GetContextUrlFn(componentIdentifier, "dataContextBaseUrl");
                return PDF.Utils.GetJSON(getUrl("/api/pdf/ProductSupportRouter/get"), { key: 'ProductSupportRouter', locale: PDF.Config.locale }, function (data)
                {
                    if (data)
                    {
                        self.ProductList = data;
                    }
                });
            };


            self.AllProductsList = null;

            self.LoadAllProducts = function ()
            {
                var getUrl = PDF.Utils.GetContextUrlFn(componentIdentifier, "dataContextBaseUrl");
                return PDF.Utils.GetJSON(getUrl("/api/pdf/product/get"), { key: 'ProductFamilies', context: '' }, function (data)
                {
                    if (data)
                    {
                        self.AllProductsList = data;
                    }
                });
            };

            return self;
        }($)
        promise.await([
           viewModel.LoadProduct(),
           viewModel.LoadAllProducts()

        ], viewModel);
        promise.done(function ()
        {

        });

        //end product selector code. 
    }
})(window, jQuery, ko, PDF);
