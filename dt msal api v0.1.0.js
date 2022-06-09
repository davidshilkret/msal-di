// Copyrights belong to Discover Technologies LLC. © 2019 Discover Technologies LLC
// Script Continued Below
var DocIntMSAL = (function () {
    var _instance = null;

    const LOGOUT_REDIRECT = 'docint_logout_redirect';
    const CACHED_CONFIG = 'docint_cached_config';

    const ACTION = {
        can_copy: "can_copy"
    };

    DocIntMSAL = function () {
        return DocIntMSAL.prototype._instance;
    };

    DocIntMSAL = function (clientId, tenant, spUri, noPopup) {
        var configPopup = {
            'tenant': tenant,
            'clientId': clientId,
            'endpoints': {
                'graphApiUri': 'https://graph.microsoft.com',
                'sharePointUri': spUri
            },
            'cacheLocation': 'localStorage',
            popUp: true,
            displayCall: function (urlNavigate) {

                var loginUrl = urlNavigate;
                var actx = this;
                //$sce.trustAsResourceUrl(loginUrl);
                //$scope.url = loginUrl;
                var popupWindow = window.open(loginUrl, "login", 'width=483, height=600');

                //per cert.
                //popupWindow.opener = null;
                //roll back
                if (popupWindow && popupWindow.focus)
                    popupWindow.focus();
                var registeredRedirectUri = this.redirectUri;

                var pollTimer = window.setInterval(function () {
                    if (!popupWindow || popupWindow.closed || popupWindow.closed === undefined) {
                        window.clearInterval(pollTimer);
                        authWait = false;
                    }

                    try {
                        if (popupWindow.document.URL.indexOf(registeredRedirectUri) != -1) {
                            window.clearInterval(pollTimer);

                            var dta = new DocIntMSAL();

                            if (dta.authContext.isCallback(popupWindow.location.hash)) {
                                var reqInfo = dta.authContext.getRequestInfo(popupWindow.location.hash);

                                dta.authContext.saveTokenFromHash(reqInfo);

                                //window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
                                //authContext.login();
                                //authContext.saveTokenFromHash(popupWindow.location.hash);
                            }

                            var userctx = dta.authContext.getCachedUser();

                            //?????? 
                            window.location.hash = popupWindow.location.hash;

                            //added to test
                            window.location.reload();

                            window.postMessage(popupWindow.location.hash, window.location.origin);

                            popupWindow.close();
                        }
                    } catch (e) {
                        //alert('error $$$ ' + e);
                    }
                }, 20);
            }
        };

        var configNoPopup = {
            'tenant': tenant,
            'clientId': clientId,
            'postLogoutRedirectUri': window.top.location.origin + "/x_dtll2_contentint_LogoutNoPopUpRedirect.do",
            'redirectUri': window.top.location.origin + "/x_dtll2_contentint_LoginNoPopUpRedirect.do",
            'endpoints': {
                'graphApiUri': 'https://graph.microsoft.com',
                'sharePointUri': spUri
            },
            'cacheLocation': 'localStorage',
            displayCall: function (urlNavigate) {
                //address iframe x-frame-option deny issue, use window.top
                window.top.location.replace(urlNavigate);
            },
            displayLogout: function (urlNavigate) {
                window.top.location.replace(urlNavigate);
            }
        };


        if (noPopup) {

            this.config = configNoPopup;

        } else {
            this.config = configPopup;
        }

        if (DocIntMSAL.prototype._instance) {
            return DocIntMSAL.prototype._instance;
        }

        this.authContext = new AuthenticationContext(this.config);

        //always save config
        this.authContext._saveItem(CACHED_CONFIG, JSON.stringify(this.config));

        DocIntMSAL.prototype._instance = this;

        return DocIntMSAL.prototype._instance;
    };

    DocIntMSAL.prototype.context = new function () {
        return DocIntMSAL.prototype._instance;
    };

    DocIntMSAL.prototype.log = function (msg) {};

    DocIntMSAL.prototype.loginUser = function () {
        var user = this.authContext.getCachedUser();

        return user;
    };

    DocIntMSAL.prototype.login = function () {

        // this.CONSTANTS = {
        //     CACHED_CONFIG: 'docint_cached_config',
        //     LOGOUT_REDIRECT: 'docint_logout_redirect'
        // };
        //persist this config
        this.authContext._saveItem(CACHED_CONFIG, JSON.stringify(this.config));

        this.authContext.login();

        //next line will next be run
        //user = this.authContext.getCachedUser();
    };

    DocIntMSAL.prototype.logout = function () {

        this.authContext._saveItem(LOGOUT_REDIRECT, window.top.location.href);
        this.authContext.logOut();
    };

    DocIntMSAL.prototype.validateUser = function () {
        var user = this.authContext.getCachedUser();

        // no user found
        if (user == null) {
            this.authContext.login();

            // todo: come up with a cleaner auth check
            //aync
            user = this.authContext.getCachedUser();
        }

        return user;
    };

    DocIntMSAL.prototype.getToken = function (callback, graph) {
        var user = this.validateUser();
        // var cachedToken = this.adal.getCachedToken(client_id_goes_here);
        // if (cachedToken) {
        //    this.adal.acquireToken("https://graph.microsoft.com", function(error, token) {
        //         jslog(error);
        //         jslog(token);
        //     });
        // }

        // no user found
        if (user != null) {
            //var cachedToken = this.authContext.getCachedToken (this.authContext.config.clientId);
            // ensure we have made the token request

            //var resource = this.authContext.config.clientId;
            var resource = this.authContext.config.endpoints.sharePointUri;

            if (graph) {
                //resource = this.authContext.config.clientId;
                // resource = this.authContext.config.endpoints.graphApiUri;
                //resource = "https://graph.microsoft.com";
            }

            this.authContext.acquireToken(resource, function (error, token, err, tokenType) {

                //this.authContext.acquireToken(this.authContext.config.endpoints.sharePointUri, function (error, token) {
                // this.authContext.acquireToken(this.authContext.config.clientId, function (error, token) {

                //manually testing error, remove after testing
                //error = 'Token renewal operation failed due to timeout';

                if (error) {
                    console.debug("Acquire Token Error " + err + " tokenType" + tokenType);

                    // todo: ref a singletone instance
                    var authContext = new AuthenticationContext();

                    // cheesy
                    if (error == 'Token renewal operation failed due to timeout') {
                        //authContext.acquireTokenPopup(authContext.config.clientId, null, null, function (error, token) {

                        if (authContext.config.popUp) {
                            authContext.acquireTokenPopup(resource, null, null, function (error, token) {
                                //authContext.acquireTokenPopup(authContext.config.endpoints.sharePointUri, null, null, function (error, token) {
                                if (error) {
                                    //alert('error === ' + error);
                                } else {
                                    // authorization
                                    var bearer = 'Bearer ' + token;

                                    if (callback != null) {
                                        callback(bearer);
                                    }
                                }
                            });
                        } else {
                            authContext.acquireTokenRedirect(resource, null, null);
                        }
                    }
                } else {
                    // authorization
                    var bearer = 'Bearer ' + token;

                    if (callback != null) {
                        callback(bearer);
                    }
                }
            });
        }

        return user;
    };


    return DocIntMSAL;
}());


/* JSONPath 0.8.0 - XPath for JSON
 *
 * Copyright (c) 2007 Stefan Goessner (goessner.net)
 * Licensed under the MIT (MIT-LICENSE.txt) licence.
 */
function jsonPath(obj, expr, arg) {
    var P = {
        resultType: arg && arg.resultType || "VALUE",
        result: [],
        normalize: function (expr) {
            var subx = [];
            return expr.replace(/[\['](\??\(.*?\))[\]']/g, function ($0, $1) {
                    return "[#" + (subx.push($1) - 1) + "]";
                })
                .replace(/'?\.'?|\['?/g, ";")
                .replace(/;;;|;;/g, ";..;")
                .replace(/;$|'?\]|'$/g, "")
                .replace(/#([0-9]+)/g, function ($0, $1) {
                    return subx[$1];
                });
        },
        asPath: function (path) {
            var x = path.split(";"),
                p = "$";
            for (var i = 1, n = x.length; i < n; i++)
                p += /^[0-9*]+$/.test(x[i]) ? ("[" + x[i] + "]") : ("['" + x[i] + "']");
            return p;
        },
        store: function (p, v) {
            if (p) P.result[P.result.length] = P.resultType == "PATH" ? P.asPath(p) : v;
            return !!p;
        },
        trace: function (expr, val, path) {
            if (expr) {
                var x = expr.split(";"),
                    loc = x.shift();
                x = x.join(";");
                if (val && val.hasOwnProperty(loc))
                    P.trace(x, val[loc], path + ";" + loc);
                else if (loc === "*")
                    P.walk(loc, x, val, path, function (m, l, x, v, p) {
                        P.trace(m + ";" + x, v, p);
                    });
                else if (loc === "..") {
                    P.trace(x, val, path);
                    P.walk(loc, x, val, path, function (m, l, x, v, p) {
                        typeof v[m] === "object" && P.trace("..;" + x, v[m], p + ";" + m);
                    });
                } else if (/,/.test(loc)) { // [name1,name2,...]
                    for (var s = loc.split(/'?,'?/), i = 0, n = s.length; i < n; i++)
                        P.trace(s[i] + ";" + x, val, path);
                } else if (/^\(.*?\)$/.test(loc)) // [(expr)]
                    P.trace(P.eval(loc, val, path.substr(path.lastIndexOf(";") + 1)) + ";" + x, val, path);
                else if (/^\?\(.*?\)$/.test(loc)) // [?(expr)]
                    P.walk(loc, x, val, path, function (m, l, x, v, p) {
                        if (P.eval(l.replace(/^\?\((.*?)\)$/, "$1"), v[m], m)) P.trace(m + ";" + x, v, p);
                    });
                else if (/^(-?[0-9]*):(-?[0-9]*):?([0-9]*)$/.test(loc)) // [start:end:step]  phyton slice syntax
                    P.slice(loc, x, val, path);
            } else
                P.store(path, val);
        },
        walk: function (loc, expr, val, path, f) {
            if (val instanceof Array) {
                for (var i = 0, n = val.length; i < n; i++)
                    if (i in val)
                        f(i, loc, expr, val, path);
            } else if (typeof val === "object") {
                for (var m in val)
                    if (val.hasOwnProperty(m))
                        f(m, loc, expr, val, path);
            }
        },
        slice: function (loc, expr, val, path) {
            if (val instanceof Array) {
                var len = val.length,
                    start = 0,
                    end = len,
                    step = 1;
                loc.replace(/^(-?[0-9]*):(-?[0-9]*):?(-?[0-9]*)$/g, function ($0, $1, $2, $3) {
                    start = parseInt($1 || start);
                    end = parseInt($2 || end);
                    step = parseInt($3 || step);
                });
                start = (start < 0) ? Math.max(0, start + len) : Math.min(len, start);
                end = (end < 0) ? Math.max(0, end + len) : Math.min(len, end);
                for (var i = start; i < end; i += step)
                    P.trace(i + ";" + expr, val, path);
            }
        },
        eval: function (x, _v, _vname) {
            try {
                return $ && _v && eval(x.replace(/@/g, "_v"));
            } catch (e) {
                throw new SyntaxError("jsonPath: " + e.message + ": " + x.replace(/@/g, "_v").replace(/\^/g, "_a"));
            }
        }
    };

    var $ = obj;
    if (expr && obj && (P.resultType == "VALUE" || P.resultType == "PATH")) {
        P.trace(P.normalize(expr).replace(/^\$;/, ""), obj, "$");
        return P.result.length ? P.result : false;
    }
}


/**
 ************************
 Validate Site 
 ************************
 */
function doSiteValidationProcess(caller, isSubmit) {
    //
    g_scratchpad.isSubmit = isSubmit;

    var site_name = g_form.getValue('site_name');

    var is_root = g_form.getValue('is_root');
    var managed_path = g_form.getValue('managed_path');

    var connectionType = caller.conn_type;
    var conn_base_url = caller.base_url;

    //var relative_url = (is_root == "true") ? site_name : ((site_name[0] != '/') ? "/sites/" + site_name : "/sites" + site_name);

    var relative_url;

    //default
    var final_managed_path = "sites";

    if (managed_path && managed_path != "") {
        final_managed_path = managed_path;
    }

    if (final_managed_path[0] == '/') {
        //if leading slash, drop it
        final_managed_path = final_managed_path.substring(1);
    }

    g_form.setValue('managed_path', final_managed_path);

    if (is_root == "true") {
        relative_url = site_name;
    } else {
        // //default
        // var final_managed_path = "sites";

        // if (managed_path && managed_path != "") {
        //     final_managed_path = managed_path;
        // }

        // if (final_managed_path[0] == '/') {
        //     //if leading slash, drop it
        //     final_managed_path = final_managed_path.substring(1);
        // }

        if (site_name[0] != '/') {
            relative_url = "/" + final_managed_path + "/" + site_name;
        } else {
            relative_url = "/" + final_managed_path + site_name;
        }
    }

  
    var dtsp = new DocIntSharePoint(conn_base_url, relative_url, '', connectionType);

    if (connectionType != SP_ONLINE_TYPE) {
        if (connectionType.toUpperCase() == 'SP2016-CDL') {
            //
            var cdl_addin = g_form.getValue("cross_domain_add_in_web");
            var cdl_host = g_form.getValue("cdl_host_web_url");

            if (!cdl_addin || !cdl_host) {

                var msgObj = JSON.parse(getMessage("-1031"));
                g_form.addErrorMessage(msgObj.errorMessage);

            } else {

                var conn = {};
                conn.baseURL = conn_base_url;
                conn.relative_url = relative_url;
                conn.list_title = '';
                conn.connType = connectionType;
                conn.libraryID = '';
                conn.cdl_addin = cdl_addin;
                conn.cdl_host = cdl_host;
                //
                conn.managed_path = managed_path;
                conn.is_root = is_root;

                //var sp2 = new DocIntSharePointV3(conn_base_url, relative_url, '', connectionType, '', cdl_addin, cdl_host);
                var sp2 = new DocIntSharePointV3(conn);

                g_scratchpad.hostname = conn_base_url.replace('http://', '').replace('https://', '');
                g_scratchpad.relative_url = relative_url;

                sp2.validateSPSiteCDL(ValidateSPSiteSuccessHandler, ValidateSPSiteErrorHandler);
            }
            return;
        } else {
            var resp = dtsp.validateSPSite(null, relative_url);
            processsResult(resp);
        }
    } else {
        var client_id = caller.client_id;
        var tenant_name = caller.tenant_name;

        var spo_login_mode = caller.spo_login_mode;

        var login_no_popup = (spo_login_mode == "redirect_page");

        var _dtadlal = new DocIntMSAL(client_id, tenant_name, conn_base_url, login_no_popup);

        _dtadlal.getToken(function (token) {
            var resp = dtsp.validateSPSite(token, relative_url);
            processsResult(resp);
        }, _dtadlal.authContext);
    }
}

function ValidateSPSiteSuccessHandler(data) {
    var siteWebInfo = JSON.parse(data.body);
    var siteCollectionInfo = g_scratchpad.siteCollectionInfo;

    var result = {};

    var siteId = g_scratchpad.hostname + "," + siteCollectionInfo.d.Id + "," + siteWebInfo.d.Id;

    result.siteId = siteId;
    result.siteTitle = siteWebInfo.d.Title;
    result.siteUrl = siteWebInfo.d.Url;
    result.siteDescription = siteWebInfo.d.Description;

    result.relative_url = g_scratchpad.relative_url;

    processsResult(result);
}

function ValidateSPSiteErrorHandler(data, errorCode, errorMessage) {
    g_form.clearMessages();

    var msgObj = JSON.parse(getMessage("-1030"));
    g_form.addErrorMessage("Error: " + msgObj.errorMessage + ", Status [" + data.statusCode + "], Details: " + data.body);
}

function showDisplayReadOnly(g_form, field_name, field_value) {
    g_form.setValue(field_name, field_value);
    g_form.setReadOnly(field_name, true);
    g_form.setDisplay(field_name, true);
}

function processsResult(resp) {
    if (resp && !resp.error) {
        g_form.setValue('site_id', resp.siteId);
        g_form.setValue('site_relative_url', resp.relative_url);
        //
        g_form.getElement("validation").value = true;

        g_form.clearMessages();
        g_form.addInfoMessage('Site Information is valid');

        ////////////// AUTO CONFIG ////////////////////
        var today_date = new Date();
        var today_date_str = formatDate(today_date, g_user_date_time_format);

        showDisplayReadOnly(g_form, "site_description", resp.siteDescription);
        showDisplayReadOnly(g_form, "site_title", resp.siteTitle);
        showDisplayReadOnly(g_form, "last_validated", today_date_str);

        if (!g_form.getValue("display_name")) {
            g_form.setValue("display_name", resp.siteTitle);
        }

        if (!g_form.getValue("site_usage")) {
            g_form.setValue("site_usage", resp.siteDescription);
        }

        //////////////////////////////////////////////	
        //use isSubmit to control if submit/update or just validate
        if (g_scratchpad.isSubmit) {
            setTimeout(function () {
                //sysverb_insert_bottom or sysverb_update_bottom
                var updateElement = g_form.getElement('sysverb_update_bottom');

                if (updateElement) {
                    //default is update
                    g_form.submit();
                } else {
                    //sysverb_update vs sysverb_insert
                    g_form.submit("sysverb_insert");
                }
            }, 1000);
        }
    } else {
        g_form.clearMessages();
        var msgObj = JSON.parse(getMessage("-1032"));
        g_form.addErrorMessage(msgObj.errorMessage + "<br> Error" + resp.error);
        //g_form.addErrorMessage('Site name entered is not valid <br>Error: ' + resp.error);
    }
}
/**
 ************************
 Validate Site 
 ************************
 */




function validateLibrary(g_form, callback) {

    var isValid = false;

    var currentLibraryId = g_form.getUniqueValue();

    //sys_id of connection
    var connection = g_form.getValue('connectionid');
    var conn_base_url = g_scratchpad.conn_base_url;
    var instrumentationKey = g_scratchpad.conn_instrumentation_key;
    var relative_url = g_form.getValue('relative_url');
    var list_title = g_form.getValue('list_title');

    var one_drive_id = g_form.getValue('one_drive_id');
    var content_type = g_form.getValue('content_type');
    var sp_site_id = g_form.getValue('sp_site_id');
    var sp_library_id = g_form.getValue('library_id');

    var connectionType = g_scratchpad.conn_type;

    //cdl

    var cdl_addin = g_scratchpad.site_cdl_addin;
    var cdl_host = g_scratchpad.site_cdl_host;
    //??   var caller_site = g_form.getReference("site", populateLibraryListFromSite);

    //var caller_site = g_form.getReference("site", validateLibraryFromSite);

    //function validateLibraryFromSite(caller_site) {
    //var relative_url = caller_site.site_relative_url;

    if (connectionType.toUpperCase() != SP_ONLINE_TYPE) {

        if (connectionType.toUpperCase() == 'SP2016-CDL') {

            //var cdl_addin = caller_site.cross_domain_add_in_web;
            //var cdl_host = caller_site.cdl_host_web_url;
            var conn = {};
            conn.baseURL = conn_base_url;
            conn.relative_url = relative_url;
            conn.list_title = list_title;
            conn.connType = connectionType;
            conn.libraryID = sp_library_id;
            conn.cdl_addin = cdl_addin;
            conn.cdl_host = cdl_host;

            //var sp2 = new DocIntSharePointV3(conn_base_url, relative_url, list_title, connectionType, sp_library_id, cdl_addin, cdl_host);
            var sp2 = new DocIntSharePointV3(conn);

            sp2.testLibraryNoJQueryV2();
        } else {
            var sp = new DocIntSharePoint(conn_base_url, relative_url, list_title, connectionType, sp_library_id);
            var result = sp.testLibraryNoJQuery(null, instrumentationKey);

            if (result && !result.error && result.data) {

                var responseDataObj = JSON.parse(result.data);

                if (responseDataObj.d) {
                    g_form.addInfoMessage(DOCINT_CONSTANT_MSG_TEST_LIB_OK);
                    isValid = true;

                    var splibraryId = responseDataObj.d.Id;
                    var itemType = responseDataObj.d.ListItemEntityTypeFullName;

                    //         libaryResult.sp_library_id = resultD.Id;
                    //         libaryResult.sp_item_type = resultD.ListItemEntityTypeFullName;
                    //update additional meta-data from test library call
                    var ga = new GlideAjax('ValidateConnectionAjaxService');
                    ga.addParam('sysparm_name', 'updateSPLibrary');

                    ga.addParam('sysparm_currentLibrarySysId', currentLibraryId);
                    ga.addParam('sysparm_sp_library_id', splibraryId);
                    ga.addParam('sysparm_sp_item_type', itemType);

                    ga.getXML(function () {});
                } else {
                    //g_form.addErrorMessage("Failed retrieving library list from SharePoint response \n" + result.data);
                    var msgObj = JSON.parse(getMessage("-1044"));

                    g_form.addErrorMessage(msgObj.errorMessage + "\n" + result.data);
                }
            } else {
                g_form.addErrorMessage(result.error);
            }
        }
    } else {

        var ga = new GlideAjax('ValidateConnectionAjaxService');
        ga.addParam('sysparm_name', 'validateSPOLibrary');

        ga.addParam('sysparm_currentConnectionId', connection);
        ga.addParam('sysparm_relative_url', relative_url);
        ga.addParam('sysparm_list_title', list_title);
        ga.addParam('sysparm_content_type', content_type);
        ga.addParam('sysparm_one_drive_id', one_drive_id);
        ga.addParam('sysparm_sp_site_id', sp_site_id);
        ga.addParam('sysparm_sp_library_id', sp_library_id);

        ga.getXMLAnswer(callback);

        return;
    }
    //}
}

// function checkCanCopyCanLink(canCopyValue, canLinkValue) {
//     var canCopy = canCopyValue == 'true';
//     var canCreateLink = canLinkValue == 'true';

//     if (canCreateLink && !canCopy) {
//         g_form.showFieldMsg("can_create_link", "'Can Copy' is unchecked for this rule, any backend events for copying documents from SharePoint to SharePoint will create a link, instead of copies.", "warning", false);
//     } else {
//         g_form.hideFieldMsg("can_create_link");
//     }
// }

// function checkPermission(permission, ruleType) {

//     if (ruleType == 'Open (Exclusion Rules)') {
//         if (permission == 'Include') {

//             g_form.showFieldMsg("permission", "The 'rule type' is currently set to  'Open (Exclusion Rules)'. Inclusion rules will be ignored. use exclusion rules instead.", "warning", false);

//         }
//     } else {
//         if (permission == 'Exclude') {
//             g_form.showFieldMsg("permission", "The 'rule type' is currently set to  'Strict (Inclusion Rules)'. Exclusion rules will be ignored. use inclusion rules instead.", "warning", false);
//         }
//     }
// }

function getContentType(g_form, expr, arg) {
    var currentLibraryId = g_form.getUniqueValue();

    //sys_id of connection
    var connection = g_form.getValue('connectionid');
    var conn_base_url = g_scratchpad.conn_base_url;

    var library_name = g_form.getValue('name');
    var relative_url = g_form.getValue('relative_url');
    var list_title = g_form.getValue('list_title');

    var one_drive_id = g_form.getValue('one_drive_url');
    var content_type = g_form.getValue('content_type');
    var sp_site_id = g_form.getValue('sp_site_id');
    var sp_library_id = g_form.getValue('library_id');

    var connectionType = g_scratchpad.conn_type;

    var client_id = g_scratchpad.conn_client_id;
    var tenant_name = g_scratchpad.conn_tenant_name;

    var caller = g_form.getReference("connectionid", doGetConnection);

    function getSPContentTypeSuccessHandler(data) {
        var result = data.body;
        processContentTypes(result);
    }

    function getSPContentTypeErrorHandler(data, errorCode, errorMessage) {
        alert("" + errorMessage);
    }

    function doGetConnection(caller) {
        //alert(" curent content_type  " + content_type);

        connectionType = caller.conn_type;
        var contentTypes = null;

        if (connectionType != SP_ONLINE_TYPE) {

            if (connectionType.toUpperCase() == 'SP2016-CDL') {

                var cdl_addin = g_scratchpad.site_cdl_addin;
                var cdl_host = g_scratchpad.site_cdl_host;

                var conn = {};
                conn.baseURL = conn_base_url;
                conn.relative_url = relative_url;
                conn.list_title = list_title;
                conn.connType = connectionType;
                conn.libraryID = sp_library_id;
                conn.cdl_addin = cdl_addin;
                conn.cdl_host = cdl_host;

                //var sp2 = new DocIntSharePointV3(conn_base_url, relative_url, list_title, connectionType, sp_library_id, cdl_addin, cdl_host);
                var sp2 = new DocIntSharePointV3(conn);

                sp2.getSPLibraryContentTypesNoJQueryCDL(getSPContentTypeSuccessHandler, getSPContentTypeErrorHandler);

                return;
            } else {
                var sp = new DocIntSharePoint(conn_base_url, relative_url, list_title, connectionType, sp_library_id);

                contentTypes = sp.getSPContentTypeNoJQuery(null);
            }
        } else {

            var spo_login_mode = caller.spo_login_mode;

            var login_no_popup = (spo_login_mode == "redirect_page");

            var _dtadlal = new DocIntMSAL(client_id, tenant_name, conn_base_url, login_no_popup);


            //var _dtadlal = new DocIntMSAL(client_id, tenant_name, conn_base_url);

            var dtsp = new DocIntSharePoint(conn_base_url, relative_url, list_title, connectionType, sp_library_id);

            _dtadlal.getToken(function (token) {

                contentTypes = dtsp.getSPContentTypeNoJQuery(token);

            }, _dtadlal.authContext);
        }

        //sync call based on contentTypes

        processContentTypes(contentTypes);


    }

    function processContentTypes(contentTypes) {
        var spContentTypeResponseObject = JSON.parse(contentTypes);

        var contentTypesBlob = [];

        var currValue = g_form.getValue('content_type');
        g_form.clearOptions('content_type');

        g_form.addOption('content_type', 'ALL', 'ALL');
        //if (spContentTypeResponseObject) g_form.clearOptions('content_type');

        for (var j = 0; j < spContentTypeResponseObject.d.results.length; j++) {

            var curr = spContentTypeResponseObject.d.results[j];

            var contentTypeElement = {};

            //Mapping Rules
            contentTypeElement.Name = curr.Name;
            contentTypeElement.ContentTypeId = curr.Id.StringValue;
            contentTypeElement.Selected = false;

            contentTypesBlob.push(contentTypeElement);

            //alway set
            g_form.addOption('content_type', contentTypeElement.Name, contentTypeElement.Name);

            if (currValue != null) {
                if (currValue != contentTypeElement.Name) {
                    //g_form.addOption('content_type', contentTypeElement.Name, contentTypeElement.Name);
                } else {
                    // g_form.addOption('content_type', contentTypeElement.Name, contentTypeElement.Name);
                    g_form.setValue('content_type', currValue);
                }
            } else {
                //null, never set before
                //g_form.addOption('content_type', contentTypeElement.Name, contentTypeElement.Name);
            }
        }

        if (g_form.isNewRecord()) {
            //new record, database record is not committed yet to update
            g_form.setValue('content_type_blob', JSON.stringify(contentTypesBlob, null, 2));

            var answer = {
                error: false,
                message: "Content Types Added"
            };

            processGformAnswerWrapper(JSON.stringify(answer, null, 2));
        } else {

            var ga = new GlideAjax('ContentTypeAjaxService');
            ga.addParam('sysparm_name', 'insertContentTypes');

            ga.addParam('sysparm_libSysId', currentLibraryId);
            ga.addParam('sysparm_content_type_blob', JSON.stringify(contentTypesBlob, null, 2));

            ga.getXML(processGetContentTypeAjaxResponse);
        }
    }
}

function processGetContentTypeAjaxResponse(validationResponse) {
    var answer = validationResponse.responseXML.documentElement.getAttribute("answer");
    // g_form.clearMessages();
    // g_form.addInfoMessage(" " + answer);

    processGformAnswerWrapper(answer);
}

function populateDropdownFromJSONBlob(g_form, blob_field, dropdown_field, json_field) {

    var blob = g_form.getValue(blob_field);

    var currentValue = g_form.getValue(dropdown_field);

    if (blob) {
        var jsonBlob = JSON.parse(blob);

        for (var key in jsonBlob) {
            if (jsonBlob.hasOwnProperty(key)) {
                var obj = jsonBlob[key];

                var f = "Name";

                if (json_field) {
                    f = json_field;
                }

                if (currentValue != obj[f]) {
                    //don't add the current value
                    g_form.addOption(dropdown_field, obj[f], obj[f]);
                }
            }
        }
    }

}

function userHasRole(role) {
    var ga = new GlideAjax('PlatIntAjaxService');
    ga.addParam('sysparm_name', 'u_hasRole');
    ga.addParam('sysparm_role', role);
    ga.getXMLWait();
    return ga.getAnswer();
}

function renderMessage(message, document) {

    setTimeout(function () {
        var divElement = $j("#" + "notes");

        //document.getElementById('notes');

        divElement.html(message);
    }, 500);
}

//called by UI Macro, CommonMacro
function renderIframe(anchor, document) {

    // do not remove this empty script,  it is used to find the formatter.
    //var anchor = "${jvar_var_anchor}";

    var var_anchor = $j("#" + anchor);

    var closeSpan = var_anchor.closest("span")[0];
    var tabName = closeSpan.children[0].children[0].innerText;
    var sectionId = closeSpan.id;
    var sectionSysId = sectionId.substring(8);

    //alert('DocInt Common Macro, Section ' + tabName + " section id " + sectionSysId );

    var ga = new GlideAjax('PlatIntAjaxService');
    ga.addParam('sysparm_name', 'getSectionRegistrationInfo');
    ga.addParam('sysparm_form_section_sysid', sectionSysId);

    //need getXMLWait synchronouse here to ensure iframe is correctly loaded
    ga.getXMLWait();

    var configItems = ga.getAnswer();


    var itemObj = JSON.parse(configItems);

    var platIntItem = itemObj[0];
    //Inject logic to handle scenarios where no Platform Integration Exists. Patch 4 Story: STRY0086733 
    if (platIntItem) {
        var isActive = (platIntItem.active == 1);

        var target = anchor + "_iframe1";

        var widget = platIntItem.widget_name;
        var view_id = platIntItem.view_id;

        var page_id = platIntItem.page;
        var portal_url_suffix = platIntItem.portal;

        var timeout_time = 1000;

        var parentURL;
        var table;
        var sys_id;

        parentURL = location.href;

        var url = new URL(parentURL);
        sys_id = url.searchParams.get('sys_id') || '';
        table = url.searchParams.get('sysparm_record_target') || '';

        if (widget == "list_view_widget") {
            var map_sys_id;

            //Entering  List Macro
            var timeout = setTimeout(function () {
                var iframeElement = document.getElementById(target);

                var src_prefix = "/" + portal_url_suffix + "?id=" + page_id + "&native=true&sysparm_domain_restore=false&sysparm_stack=no";

                if (isActive) {
                    iframeElement.src = src_prefix + "&table=" + table + "&map_sys_id=" + sys_id + "&vw=" + view_id;
                } else {
                    iframeElement.style.display = "none";
                }

            }, timeout_time);
        } else if (widget == "search_page_widget") {
            var recTarget;
            var query;
            var sys_id;

            var search_source = platIntItem.search_source;
            var search_field = platIntItem.search_field;

            //Entering  Search Macro
            var interval = setInterval(function () {
                if (document.readyState === 'complete') {

                    clearInterval(interval);

                    var iframeElement = document.getElementById(target);

                    var src_prefix = "/" + portal_url_suffix + "?id=" + page_id + "&native=true&t=" + search_source + "&q=";

                    if (isActive) {

                        query = document.getElementById(table + '.' + search_field).value;

                        console("query: " + query);

                        var appendStr = query + "&sys_id=" + sys_id + "&vw=" + view_id;;

                        iframeElement.src = src_prefix + appendStr;
                    } else {
                        iframeElement.style.display = "none";
                    }
                }
            }, 1000);
        }
    } else {
        var isAdmin = userHasRole('admin');
        jslog(isAdmin + typeof (isAdmin));
        if (isAdmin === 'true') {

            var url = new URL(document.URL);
            var params = new URLSearchParams(url.search);
            var sysid = params.get('sys_id');
            var msg = getMessage('-0101');
            var pMsg = JSON.parse(msg);
            var errorMessage = pMsg.errorMessage;
            var errorCode = pMsg.errorCode;
            jslog(JSON.parse(msg));
            var urlLocation = document.location.origin;
            var table = url.searchParams.get('sysparm_record_target') || '';
            var newRecord = urlLocation + "/nav_to.do?uri=%2Fx_dtll2_contentint_platform_integration.do%3Fsys_id%3D-1%26sysparm_query%3Ddescription%3DPlatform Integration item based on the " + table + " table.";
            jslog(newRecord);

            var msgDiv = document.getElementsByTagName("div");
            var href = "<a href='" + newRecord + "' target='" + "_blank" + "'>Create new Platform Integration item</a>";
            msgDiv = document.writeln("<div>" + "error code: " + errorCode + " " + errorMessage + " Use the following URL to setup a new Platform Integration item:<div><br>" + href);
            //alert(newRecord);
        } else {
            var msgDiv = document.getElementsByTagName("div");
            msgDiv = document.writeln("<div>You need to be an admin user to setup a new platform integration.<div><br>");

        }
    }
}

function processGformAnswerWrapper(answer) {
    //var answer = response.responseXML.documentElement.getAttribute("answer");

    var isError = false;

    if (answer) {
        g_form.clearMessages();
        //g_form.addInfoMessage(" " + answer);

        var answerObj = JSON.parse(answer);

        if (answerObj.error) {
            isError = true;
            g_form.addErrorMessage(answerObj.message);
        } else {
            g_form.addInfoMessage(answerObj.message);
        }
        //location.reload();
    } else {
        isError = true;
        g_form.addErrorMessage("Server error ");
    }

    return isError;
}

const DOCINT_CONSTANT_PREV_KEY = "sn.library.form.list.title.prev";
const DOCINT_CONSTANT_MSG_TEST_LIB_OK = "Test Library Succeeds!";


//supported file format
//https://docs.microsoft.com/en-us/officeonlineserver/office-online-server-overview
const OWA_SUPPORTED_FILE_FORMATS = "pdf,doc,docx,dotx,dot,dotm,xls,xlsx,xlsm,xlm,xlsb,ppt,pptx,pps,ppsx,potx,pot,pptm,potm,ppsm";

function setInstKey(docint_ikey) {

    if (docint_ikey) {
        //var IKey = document.getElementsByName('IKey')[0].value;
        ! function (T, l, y) {
            var S = T.location,
                u = "script",
                k = "instrumentationKey",
                D = "ingestionendpoint",
                C = "disableExceptionTracking",
                E = "ai.device.",
                I = "toLowerCase",
                b = "crossOrigin",
                w = "POST",
                e = "appInsightsSDK",
                t = y.name || "appInsights";
            (y.name || T[e]) && (T[e] = t);
            var n = T[t] || function (d) {
                var g = !1,
                    f = !1,
                    m = {
                        initialize: !0,
                        queue: [],
                        sv: "4",
                        version: 2,
                        config: d
                    };

                function v(e, t) {
                    var n = {},
                        a = "Browser";
                    return n[E + "id"] = a[I](), n[E + "type"] = a, n["ai.operation.name"] = S && S.pathname || "_unknown_", n["ai.internal.sdkVersion"] = "javascript:snippet_" + (m.sv || m.version), {
                        time: function () {
                            var e = new Date;

                            function t(e) {
                                var t = "" + e;
                                return 1 === t.length && (t = "0" + t), t
                            }
                            return e.getUTCFullYear() + "-" + t(1 + e.getUTCMonth()) + "-" + t(e.getUTCDate()) + "T" + t(e.getUTCHours()) + ":" + t(e.getUTCMinutes()) + ":" + t(e.getUTCSeconds()) + "." + ((e.getUTCMilliseconds() / 1e3).toFixed(3) + "").slice(2, 5) + "Z"
                        }(),
                        iKey: e,
                        name: "Microsoft.ApplicationInsights." + e.replace(/-/g, "") + "." + t,
                        sampleRate: 100,
                        tags: n,
                        data: {
                            baseData: {
                                ver: 2
                            }
                        }
                    }
                }
                var h = d.url || y.src;
                if (h) {
                    function a(e) {
                        var t, n, a, i, r, o, s, c, p, l, u;
                        g = !0, m.queue = [], f || (f = !0, t = h, s = function () {
                            var e = {},
                                t = d.connectionString;
                            if (t)
                                for (var n = t.split(";"), a = 0; a < n.length; a++) {
                                    var i = n[a].split("=");
                                    2 === i.length && (e[i[0][I]()] = i[1])
                                }
                            if (!e[D]) {
                                var r = e.endpointsuffix,
                                    o = r ? e.location : null;
                                e[D] = "https://" + (o ? o + "." : "") + "dc." + (r || "services.visualstudio.com")
                            }
                            return e
                        }(), c = s[k] || d[k] || "", p = s[D], l = p ? p + "/v2/track" : config.endpointUrl, (u = []).push((n = "SDK LOAD Failure: Failed to load Application Insights SDK script (See stack for details)", a = t, i = l, (o = (r = v(c, "Exception")).data).baseType = "ExceptionData", o.baseData.exceptions = [{
                            typeName: "SDKLoadFailed",
                            message: n.replace(/\./g, "-"),
                            hasFullStack: !1,
                            stack: n + "\nSnippet failed to load [" + a + "] -- Telemetry is disabled\nHelp Link: https://go.microsoft.com/fwlink/?linkid=2128109\nHost: " + (S && S.pathname || "_unknown_") + "\nEndpoint: " + i,
                            parsedStack: []
                        }], r)), u.push(function (e, t, n, a) {
                            var i = v(c, "Message"),
                                r = i.data;
                            r.baseType = "MessageData";
                            var o = r.baseData;
                            return o.message = 'AI (Internal): 99 message:"' + ("SDK LOAD Failure: Failed to load Application Insights SDK script (See stack for details) (" + n + ")").replace(/\"/g, "") + '"', o.properties = {
                                endpoint: a
                            }, i
                        }(0, 0, t, l)), function (e, t) {
                            if (JSON) {
                                var n = T.fetch;
                                if (n && !y.useXhr) n(t, {
                                    method: w,
                                    body: JSON.stringify(e),
                                    mode: "cors"
                                });
                                else if (XMLHttpRequest) {
                                    var a = new XMLHttpRequest;
                                    a.open(w, t), a.setRequestHeader("Content-type", "application/json"), a.send(JSON.stringify(e))
                                }
                            }
                        }(u, l))
                    }

                    function i(e, t) {
                        f || setTimeout(function () {
                            !t && m.core || a()
                        }, 500)
                    }
                    var e = function () {
                        var n = l.createElement(u);
                        n.src = h;
                        var e = y[b];
                        return !e && "" !== e || "undefined" == n[b] || (n[b] = e), n.onload = i, n.onerror = a, n.onreadystatechange = function (e, t) {
                            "loaded" !== n.readyState && "complete" !== n.readyState || i(0, t)
                        }, n
                    }();
                    y.ld < 0 ? l.getElementsByTagName("head")[0].appendChild(e) : setTimeout(function () {
                        l.getElementsByTagName(u)[0].parentNode.appendChild(e)
                    }, y.ld || 0)
                }
                try {
                    m.cookie = l.cookie
                } catch (p) {}

                function t(e) {
                    for (; e.length;) ! function (t) {
                        m[t] = function () {
                            var e = arguments;
                            g || m.queue.push(function () {
                                m[t].apply(m, e)
                            })
                        }
                    }(e.pop())
                }
                var n = "track",
                    r = "TrackPage",
                    o = "TrackEvent";
                t([n + "Event", n + "PageView", n + "Exception", n + "Trace", n + "DependencyData", n + "Metric", n + "PageViewPerformance", "start" + r, "stop" + r, "start" + o, "stop" + o, "addTelemetryInitializer", "setAuthenticatedUserContext", "clearAuthenticatedUserContext", "flush"]), m.SeverityLevel = {
                    Verbose: 0,
                    Information: 1,
                    Warning: 2,
                    Error: 3,
                    Critical: 4
                };
                var s = (d.extensionConfig || {}).ApplicationInsightsAnalytics || {};
                if (!0 !== d[C] && !0 !== s[C]) {
                    method = "onerror", t(["_" + method]);
                    var c = T[method];
                    T[method] = function (e, t, n, a, i) {
                        var r = c && c(e, t, n, a, i);
                        return !0 !== r && m["_" + method]({
                            message: e,
                            url: t,
                            lineNumber: n,
                            columnNumber: a,
                            error: i
                        }), r
                    }, d.autoExceptionInstrumented = !0
                }
                return m
            }(y.cfg);
            (T[t] = n).queue && 0 === n.queue.length && n.trackPageView({})
        }(window, document, {
            src: "https://az416426.vo.msecnd.net/scripts/b/ai.2.min.js", // The SDK URL Source
            //name: "appInsights", // Global SDK Instance name defaults to "appInsights" when not supplied
            //ld: 0, // Defines the load delay (in ms) before attempting to load the sdk. -1 = block page load and add to head. (default) = 0ms load after timeout,
            //useXhr: 1, // Use XHR instead of fetch to report failures (if available),
            //crossOrigin: "anonymous", // When supplied this will add the provided value as the cross origin attribute on the script tag
            cfg: { // Application Insights Configuration
                instrumentationKey: docint_ikey
                /* ...Other Configuration Options... */
            }
        });
    }
};

function allowAjaxCallDownload() {
    // use this transport for "binary" data type
    $.ajaxTransport("+binary", function (options, originalOptions, jqXHR) {
        // check for conditions and support for blob / arraybuffer response type
        if (window.FormData && ((options.dataType && (options.dataType == 'binary')) || (options.data && ((window.ArrayBuffer && options.data instanceof ArrayBuffer) || (window.Blob && options.data instanceof Blob))))) {
            return {
                // create new XMLHttpRequest
                send: function (headers, callback) {
                    // setup all variables
                    var xhr = new XMLHttpRequest(),
                        url = options.url,
                        type = options.type,
                        async = options.async || true,
                            // blob or arraybuffer. Default is blob
                            dataType = options.responseType || "blob",
                            data = options.data || null,
                            username = options.username || null,
                            password = options.password || null;

                    xhr.addEventListener('load', function () {
                        var data = {};
                        data[options.dataType] = xhr.response;
                        // make callback and send data
                        callback(xhr.status, xhr.statusText, data, xhr.getAllResponseHeaders());
                    });
                    xhr.addEventListener('error', function () {
                        var data = {};
                        data[options.dataType] = xhr.response;
                        // make callback and send data
                        callback(xhr.status, xhr.statusText, data, xhr.getAllResponseHeaders());
                    });

                    //important to make SP2013 CORS work
                    if (options.xhrFields) {
                        xhr.withCredentials = options.xhrFields.withCredentials;
                    }

                    xhr.open(type, url, async, username, password);

                    // setup custom headers
                    for (var i in headers) {
                        xhr.setRequestHeader(i, headers[i]);
                    }

                    xhr.responseType = dataType;
                    xhr.send(data);
                },
                abort: function () {}
            };
        }
    });
}


function openInMS(fileFormat, pdfReady) {
    if (fileFormat) {
        var protocol;
        var ext = fileFormat;
        //if "pdf" and "Open in native App." set to true. Next do string replace "https:" with "docint".
        if (ext == "pdf" && pdfReady == true) {
            protocol = "docint";
            return protocol;
        }
        //start MS Office block. prepend url with result of "protocol".
        switch (ext) {
            case "doc":
                protocol = "ms-word";
                break;
            case "docx":
                protocol = "ms-word";
                break;
            case "vsdx":
                protocol = "ms-visio";
                break;
            case "xlsx":
                protocol = "ms-excel";
                break;
            case "xls":
                protocol = "ms-excel";
                break;
            case "csv":
                protocol = "ms-excel";
                break;
            case "pptx":
                protocol = "ms-powerpoint";
                break;
            case "ppt":
                protocol = "ms-powerpoint";
                break;
            case "accdb":
                protocol = "ms-access";
                break;
            case "pub":
                protocol = "ms-publisher";
                break;
            default:
                protocol = "unsupported";
        }
        if (protocol == 'unsupported') {
            return protocol;
        } else {
            return protocol + ':';
        }
    }

};

var SP_ONLINE_TYPE = "O365";