/*jshint browser: true*//*global browserid_common, jQuery*/(function(){function i(a){return a||"login"}function j(a){b=a;var c={siteName:browserid_common.siteName||"",siteLogo:browserid_common.siteLogo||"",backgroundColor:browserid_common.backgroundColor||"",termsOfService:browserid_common.termsOfService||"",privacyPolicy:browserid_common.privacyPolicy||""};b==="comment"?c.returnTo=u("#submit_comment"):b==="register"&&(c.returnTo=u("#submit_registration")),navigator.id.request(c)}function k(b){var c=document.getElementById("rememberme");c!==null&&(c=c.checked);var d=document.createElement("form");d.setAttribute("style","display: none;"),d.method="POST",d.action=browserid_common.urlLoginSubmit;var e={browserid_assertion:b,rememberme:c};browserid_common.urlLoginRedirect!==null&&(e.redirect_to=browserid_common.urlLoginRedirect),v(d,e),a("body").append(d),d.submit()}function l(){d=!0,n(),j("comment")}function m(b){var e=o();if(!e&&!c)return p();var g=a("#commentform"),h=a("#comment_post_ID").val();v(g,{browserid_comment:h,browserid_assertion:b}),localStorage.removeItem("comment_hash"),sessionStorage.setItem("submitting_comment","true"),browserid_common.loggedInUser||(d=!0,navigator.id.logout()),f=!0,a("#submit").click()}function n(){var b={author:a("#author").val(),url:a("#url").val(),comment:a("#comment").val(),comment_parent:a("#comment_parent").val()};localStorage.setItem("comment_state",JSON.stringify(b))}function o(){var b=localStorage.getItem("comment_state");return b&&(b=JSON.parse(b),a("#author").val(b.author),a("#url").val(b.url),a("#comment").val(b.comment),a("#comment_parent").val(b.comment_parent),localStorage.removeItem("comment_state")),b}function p(){var a=localStorage.getItem("comment_hash");a?(localStorage.removeItem("comment_hash"),document.location.hash=a,document.location.reload(!0)):setTimeout(p,100)}function q(b){var d=s();if(!d&&!c)return t();sessionStorage.setItem("submitting_registration","true"),a("#browserid_assertion").val(b),e=!0,a("#wp-submit").click()}function r(){var b={user_login:a("#user_login").val()};localStorage.setItem("registration_state",JSON.stringify(b))}function s(){var b=localStorage.getItem("registration_state");return b&&(b=JSON.parse(b),a("#user_login").val(b.user_login),localStorage.removeItem("registration_state")),b}function t(){var a=localStorage.getItem("registration_complete");a?(localStorage.removeItem("registration_complete"),document.location=browserid_common.urlRegistrationRedirect):setTimeout(t,100)}function u(a){return document.location.href.replace(/http(s)?:\/\//,"").replace(document.location.host,"").replace(/#.*$/,"")+a}function v(b,c){b=a(b);for(var d in c){var e=document.createElement("input");e.type="hidden",e.name=d,e.value=c[d],b.append(e)}}function w(){var b=a("<div class='persona__submit'><div class='persona__submit_spinner'></div></div>");a("body").append(b)}function x(b,c){function d(){var d=a(b).val();d&&d.trim().length?a(c).removeClass("disabled"):a(c).addClass("disabled")}a(c).addClass("disabled"),a(b).keyup(d),a(b).change(d)}function y(b,c,d){a("body")[typeof a.fn.on=="function"?"on":"delegate"](c,b,d)}"use strict";var a=jQuery,b,c,d=!1,e=!1,f=browserid_common.loggedInUser||!1,g;a(".js-persona__login").click(function(a){a.preventDefault(),d=!1,j("login")}),y(".js-persona__logout","click",function(a){a.preventDefault(),d=!1,navigator.id.logout()}),browserid_common.isPersonaUsedWithComments&&(a("body").addClass("persona--comments"),a(".js-persona__submit-comment").click(function(b){b.preventDefault(),a("#commentform").submit()}),a("#commentform").submit(function(b){if(a("#comment").hasClass("disabled")){b.preventDefault();return}f||(b.preventDefault(),l())})),browserid_common.isPersonaOnlyAuth&&(a("body").addClass("persona--persona-only-auth"),x("#user_login",".js-persona__register"),a(".js-persona__register").click(function(b){b.preventDefault();if(a(b.target).hasClass("disabled"))return;d=!1,r(),j("register")}),a("#registerform").submit(function(b){if(e)return;b.preventDefault();if(a("#user_login").val().length===0)return;d=!1,r(),j("register")}));if(document.location.hash==="#submit_comment"){d=!0,w(),g=o();if(!g)return p();b="comment",c=!0}else if(document.location.hash==="#submit_registration"){d=!0,w(),g=s();if(!g)return t();b="register",c=!0}else document.location.href===browserid_common.urlRegistrationRedirect&&sessionStorage.getItem("submitting_registration")?localStorage.setItem("registration_complete","true"):sessionStorage.getItem("submitting_comment")&&(d=!0,sessionStorage.removeItem("submitting_comment"),localStorage.setItem("comment_hash",document.location.hash));if(browserid_common.msgError||a("#login_error").length)d=!0,navigator.id.logout();var h={login:k,register:q,comment:m};navigator.id.watch({loggedInUser:browserid_common.loggedInUser||null,onlogin:function(a){b=i(b);var c=h[b];c&&c(a)},onlogout:function(){if(d)return;browserid_common.loggedInUser&&(document.location=browserid_common.urlLogoutRedirect)}})})()