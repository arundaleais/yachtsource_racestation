function ReceiveServerDataForFeedBack(n){var r=$(n).find(".feedBack-flyout-body"),i;r.hide(),i=$(n).find(".feedBack-successmessage-flyout-body"),i.show()}window.Mst===undefined&&(window.Mst={}),Mst.FeedBackFlyout===undefined&&(Mst.FeedBackFlyout={}),Mst.SeoMenu=function(n,t,i,r,u,f,e,o){var s=this;s.Control=$("#"+n),s._menuClass="."+t,s._itemClass="."+i,s._flyoutLinkClass="."+r,s._flyoutClass="."+u,s._flyoutRegionClass="."+f,s._stageExpand=!1,s._onOpenClick=e,s._onCloseEvent=o,s.Control.find(s._flyoutLinkClass).click($.proxy(s.ItemClick,s)).keydown($.proxy(s.ItemKeyDown,s));var h=s.Control.find(s._flyoutClass).toArray(),c;for(c in h)$(h[c]).find("a:last").keydown($.proxy(s.LastLinkKeyDown,s));$(document).click($.proxy(s.DocClick,s))},Mst.SeoMenu.prototype={DocClick:function(n){var i=this.Control.find(".selected"),t;i.size()>0&&(t=$.contains(i.get(0),n.target),t||this.HideFlyouts())},ItemClick:function(n){var t=this;$(n.target).hasClass("mstLcp_DualLangLink")||(n.preventDefault(),t.IsFlyoutVisible(n)?t.HideFlyouts():(t.HideFlyouts(),t.ShowFlyout(n),t._onOpenClick!=null&&t._onOpenClick(n)))},ItemKeyDown:function(n){n.which==9&&n.shiftKey&&this.HideFlyouts()},LastLinkKeyDown:function(n){n.which==9&&(n.shiftKey||this.HideFlyouts())},IsFlyoutVisible:function(n){return $(n.target).parents(this._itemClass).find(this._flyoutClass).css("display")!="none"},ShowFlyout:function(n){var r=".stage-content",t=this,i=$(n.target).parents(t._itemClass).addClass("selected").find(t._flyoutClass);t.PositionFlyout(i),t._stageExpand||($(r).animate({height:$(r).outerHeight()+$(t._flyoutClass).outerHeight()},200),$("html").animate({scrollTop:$(document).height()},200),t._stageExpand=!0),jQuery.browser.opera||jQuery.browser.msie&&document.documentMode<=7?i.show():i.slideDown(200)},HideFlyouts:function(){var r=".stage-content",t=this,i=t;t._stageExpand&&($(r).animate({height:$(r).outerHeight()-$(t._flyoutClass).outerHeight()},200),t._stageExpand=!1),jQuery.browser.opera||jQuery.browser.msie&&document.documentMode<=7?$(t._flyoutClass,t.Control).hide():$(t._flyoutClass,t.Control).slideUp(200,function(){i.fixFlyoutContentOldIE()}),t._onCloseEvent!=null&&$(t._itemClass).filter(".selected").each(function(){i._onCloseEvent($(this).find(i._flyoutLinkClass))}),$(t._itemClass).removeClass("selected")},fixFlyoutContentOldIE:function(){var t,i,r,n;$.browser.msie&&parseInt($.browser.version)<=7&&(t=this,setTimeout(function(){var i="display";img=t.Control.find(".bg-img"),n=img.css(i),img.css(i,""),img.css(i,n)},1))},isIE6:function(){return navigator.appVersion.indexOf("MSIE 6.")!=-1},PositionFlyout:function(n){var s="rtl",f=this,o="left",r=n.parents(f._itemClass).width(),c=n.parents(f._itemClass).position().left,e=n.outerWidth(),h=n.parents(f._menuClass).width(),a=n.parents(f._menuClass).position().left,l,u,i,v,t;h=920,a=0,l=h,n.parents(f._flyoutRegionClass).size()>0&&(l=n.parents(f._flyoutRegionClass).innerWidth()),u=c+a,i=h-c-r,document.documentElement.dir=="ltr"&&(i+=18),document.documentElement.dir==s&&(o="right",r-=2,i=l-c-r,v=u,u=i,i=v),i<0&&(i=0),u<0&&(u=0),t=0,e>r+i?e<r+u?(t=document.documentElement.dir==s?-e+r+2:-e+r,n.css(o,t),n.addClass("dock-right")):e<u+i+r?(t=-e+(i+r),document.documentElement.dir==s&&(t-=23,f.isIE6()&&(t+=r)),n.css(o,t)):(t=document.documentElement.dir==s?-u-1:-u-6,n.css(o,t)):(n.css(o,t),n.addClass("dock-left"))}},Mst.FeedBackFlyout=function(n,t,i,r,u,f,e,o,s,h,c,l,a,v,y){var w=".",p=this;p.Id="#"+n,p.FeedBack=$(p.Id),p.Direction=c,p.FeedBackFlyoutClass=w+t,p.FeedBackFlyoutBodyClass=w+i,p.FeebackHighButtonClass=w+f,p.FeebackLowButton="#"+r,p.FeebackHighButton="#"+u,p.FeebackHighButtonClose="#"+e,p.FeedBackCloseButtonClass=w+o,p.FeedBackCancelButtonClass=w+s,p.FeedBackSubmitButtonClass=w+h,p.FeedbackBackgroundClass=w+y,p.IsFormSubmitted=!1,p.FeedBackSuccessMessageFlyoutClass=w+l,p.FeedBackSuccessMessageFormCloseClass=w+a,p.FeedBackSuccessMessageCloseClass=w+v,p.FeedBack.find(p.FeebackHighButton).click($.proxy(p.ShowFeedBackFlyout,p)),p.FeedBack.find(p.FeebackLowButton).click($.proxy(p.ShowFeedBackFlyout,p)),p.FeedBack.find(p.FeebackHighButtonClose).click($.proxy(p.HideFeedBackHighButton,p)),p.FeedBack.find(p.FeedBackCloseButtonClass).click($.proxy(p.HideFeedBackFlyout,p)),p.FeedBack.find(p.FeedBackCancelButtonClass).click($.proxy(p.HideFeedBackFlyout,p)),p.FeedBack.find(p.FeedbackBackgroundClass).click($.proxy(p.HideFeedBackFlyout,p)),p.FeedBack.find(p.FeedBackSuccessMessageFormCloseClass).click($.proxy(p.HideFeedBackFlyoutOnSucess,p)),p.FeedBack.find(p.FeedBackSuccessMessageCloseClass).click($.proxy(p.HideFeedBackFlyoutOnSucess,p)),$(document).keypress($.proxy(p.FeedBackFlyoutKeyPress,p)),p.FeedBack.find(p.FeedBackSubmitButtonClass).click($.proxy(p.SumitForm,p)),p.FeedBack.css("display","inline-block")},Mst.FeedBackFlyout.prototype={CenterAlignPopup:function(){var r="px",t=this,s=$(t.FeedBack.find(t.FeedBackFlyoutBodyClass)).height(),o=$(t.FeedBack.find(t.FeedBackSuccessMessageFlyoutClass)).height(),l,c,f,i,e,u,h;$.browser.msie&&parseInt($.browser.version)==6&&(s=parseInt($(window).height())*4,o=parseInt($(window).height())*4),l=parseInt($(window).height())/2-parseInt(s)/2,$(t.FeedBack.find(t.FeedBackFlyoutBodyClass)).css("top",l),c=parseInt($(window).height())/2-parseInt(o)/2,$(t.FeedBack.find(t.FeedBackSuccessMessageFlyoutClass)).css("top",c),$.browser.msie&&parseInt($.browser.version)<=7&&(f=$(t.FeedBack.find(t.FeedBackFlyoutBodyClass)).width(),$.browser.msie&&parseInt($.browser.version)==6&&(f=parseInt($(window).width())+f),i=parseInt($(window).width())/2-parseInt(f)/2,$.trim(t.Direction.toLowerCase())=="ltr"?($(t.FeedBack.find(t.FeedBackFlyoutBodyClass)).css("left",i+r),$(t.FeedBack.find(t.FeedBackSuccessMessageFlyoutClass)).css("left",i+r)):($(t.FeedBack.find(t.FeedBackFlyoutBodyClass)).css("right",i+r),$(t.FeedBack.find(t.FeedBackSuccessMessageFlyoutClass)).css("right",i+r)),$.browser.msie&&parseInt($.browser.version)<=7&&(e=t,setTimeout(function(){var n="display";u=e.FeedBack.find(e.FeedBackFlyoutBodyClass),h=u.css(n),u.css(n,""),u.css(n,h)},1)))},ShowFeedBackFlyout:function(n){n&&n.preventDefault();var t=$(this.FeedBack.find(this.FeedBackFlyoutClass));t.slideDown(200),this.CenterAlignPopup()},HideFeedBackFlyout:function(n){var t=this;n&&n.preventDefault();var i=$(t.FeedBack.find(t.FeedBackFlyoutClass));i.hide(),t.BiTrack(2,10,"."+$(n.currentTarget).attr("class")),t.IsFormSubmitted==!0&&t.DisableFeeback(n)},HideFeedBackFlyoutOnSucess:function(n){this.HideFeedBackFlyout(n),this.DisableFeeback(n)},DisableFeeback:function(){var t=this;t.FeedBack.find(t.FeebackLowButton).addClass("feedBack-low-link-disabled"),t.FeedBack.find(t.FeebackLowButton).unbind("click");var i=$(t.FeedBack.find(t.FeebackHighButtonClass));i.hide()},ShowFeedBackFlyoutOnSucess:function(){var t=this;t.FeedBack.find(t.FeedBackFlyoutBodyClass).hide(),t.FeedBack.find(t.FeedBackSuccessMessageFlyoutClass).show()},FeedBackFlyoutKeyPress:function(n){var t=this;n.keyCode==27&&(t.IsFormSubmitted==!1?t.HideFeedBackFlyout(n):t.HideFeedBackFlyoutOnSucess(n)),t.HandleKeyPress(n)},HandleKeyPress:function(n){this.FeedBack.find(".feedback-item").each(function(){$(this).find(".feedback-element-control").each(function(){var r=$(this).attr("type"),i;r=="textarea"&&(i=$(this).attr("maxlength"),$(this).val().length+1>i&&n.preventDefault())})})},HideFeedBackHighButton:function(n){var t=this;n&&n.preventDefault();var i=$(t.FeedBack.find(t.FeebackHighButtonClass));i.hide(),t.BiTrack(2,10,t.FeebackHighButtonClose)},SumitForm:function(n){var t=this;n&&n.preventDefault(),t.FeedBack.find(t.FeedBackSubmitButtonClass).unbind("click");var i="";t.FeedBack.find(".feedback-item").each(function(){$(this).find(".feedback-element-control").each(function(){var u=" ± ",r="name",t=this,o=$(t).attr("type"),f,e;switch(o){case"radio":t.checked&&(i+="≤"+$(t).attr(r)+u+$(t).attr("id")+"≥");break;case"checkbox":t.checked&&(i+="≤"+$(t).attr(r)+u+$(t).attr("id")+"≥");break;case"textarea":f=$.trim($(t).val()),e=$(t).attr("maxlength"),f.length>0&&(i+=$(t).val().length>e?"≤"+$(t).attr(r)+u+f.substr(0,e)+"≥":"≤"+$(t).attr(r)+u+f+"≥")}})}),t.IsFormSubmitted=!0,i.length>0?(i+="≠"+t.Id.toString(),SendFeedbackDataToServer(i,"")):t.ShowFeedBackFlyoutOnSucess(n),t.BiTrack(2,10,t.FeedBackSubmitButtonClass)},BiTrack:function(n,t,i){var u,r,f;if($.bi&&$.bi.dataRetrievers.structure){u=this.FeedBack.find(i),r={title:$.trim($(u).text())};try{$.extend(r,$.bi.baseData(),$.bi.dataRetrievers.structure.getData(u))}catch(n){}f=$.extend({},r,{interactiontype:t,cot:5,parenttypestructure:r.parenttypestructure}),$.bi.record(f)}return!0}},Mst.FooterV3===undefined&&(Mst.FooterV3={}),Mst.FooterV3=function(n,t,i,r){var u=this;u.Control=$("#"+n+"_mstFooterV3Ctl"),u.Id=n,u.StageWidth=t,u.StagePadding=i,u.IsResponsive=r,$($.proxy(u.Ready,u)),u.FlyoutHeight=u.Control.find(".mstFooterLocFlyoutContainer").height()},Mst.FooterV3.prototype={Ready:function(){var t=".mstFooterMsLinkItemLink",i=".mstLcpSearchText",n=this;n.Control.find(".mstFooterLocale").click($.proxy(n.OpenLocalePicker,n)),n.Control.find(".mstLcpClose1").click($.proxy(n.CloseLocalePicker,n)),n.Control.find(".mstLcpClose2").click($.proxy(n.CloseLocalePicker,n)),$(i).bind("keyup",$.proxy(n.FilterList,n)),$(document).click($.proxy(n.DoClick,n)),n.Control.find(i).keypress($.proxy(n.DoNothing,n)),n.Control.find(t).mouseenter($.proxy(n.OnMouseEnterMsLinkImage,n)),n.Control.find(t).mouseleave($.proxy(n.OnMouseLeaveMsLinkImage,n)),n.Control.find("#mstLocPickerCtl").append("<div class='cssClear'></div>"),$.browser.msie&&parseInt($.browser.version.substr(0,2))<7&&$(window).resize($.proxy(n.SetFlyoutBackgroundForIe6,n)),$.browser.msie&&(n.SetFooterBackgroundForIE(),$(window).resize($.proxy(n.SetFooterBackgroundForIE,n)))},SetFooterBackgroundForIE:function(){$(".mstFooterV3Backround").css("height",$(".mstFooterTop").height()+28+$(".mstFooterBottom").height()+6)},SetFlyoutBackgroundForIe6:function(){var n=$(".mstFooterLocFlyoutContainer").css("width",$(document).width()-16)},DoClick:function(n){var t=this.Control.find(".selected");if(t.size()>0){var u=$.contains(t.get(0),n.target),r=this.Control.find(".mstFooterLocFlyoutContainer"),i=$.contains(r.get(0),n.target);u||i||this.CloseLocalePicker()}},DoNothing:function(n){if(n.keyCode==13)return!1},OpenLocalePicker:function(n){var t=this;if(t.IsResponsive=="true")return;n&&n.preventDefault();if(t.Control.find(".selected").size()>0){t.CloseLocalePicker();return}var i=t.Control.find(".mstFooterLocFlyoutContainer");t.Control.find(".mstFooterLocale").addClass("selected"),i.height(t.FlyoutHeight),t.SetMinFlyoutHeight(),i.slideDown(200),window.setTimeout(function(){$("html").animate({scrollTop:i.offset().top-20},"200")},500),$(".mstLcpSearchText").focus(),t.BiTrack(n,9)},CloseLocalePicker:function(n){n&&n.preventDefault();var t=this.Control.find(".mstFooterLocFlyoutContainer");this.Control.find(".mstFooterLocale").removeClass("selected"),t.slideUp(200),window.setTimeout(function(){$(".mstLcpSearchText").val(""),$(".mstLcpAllSitesLinks ul li").each(function(){$(this).css("display","block")})},500),$.browser.webkit&&window.setTimeout(function(){t.css("display","none")},100),this.BiTrack(n,10)},SetMinFlyoutHeight:function(){var t=".mstFooterBottom",n=this,i=n.Control.find(".mstFooterLocFlyoutContainer"),r=n.Control.height();i.height()<r-n.Control.find(t).height()&&i.height(r-n.Control.find(t).height())},FilterList:function(){var i=".mstFooterLocFlyoutContainer";$(i).css("height","");var t=$(".mstLcpSearchText").val().toLowerCase();$(".mstLcpAllSitesLinks ul li").each(function(){var u=$(this).find("span").text(),n=u.split("-"),r=$.trim(n[0]).toLowerCase(),i=$.trim(n[1]).toLowerCase();r.substr(0,t.length)==t||i.substr(0,t.length)==t?$(this).css("display","block"):$(this).css("display","none")}),this.SetMinFlyoutHeight(),$(document).scrollTop($(i).offset().top-20)},OnMouseEnterMsLinkImage:function(n){var t=$(n.currentTarget).parents(".mstFooterMsLinkItemLi").find(".mstFooterMsLinkItemText").css("text-decoration","underline")},OnMouseLeaveMsLinkImage:function(n){var t=$(n.currentTarget).parents(".mstFooterMsLinkItemLi").find(".mstFooterMsLinkItemText").css("text-decoration","none")},BiTrack:function(n,t){return $.bi&&$.bi.dataRetrievers.structure&&(params={},params.interactiontype=t,params.cot=5,params.parenttypestructure=$.bi.dataRetrievers.structure.getTypeStructure(this.Control.find(".mstFooterLocaleMenu")),$.bi.record(params)),!0}}