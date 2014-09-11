/* Excel specific API library */
/* Version: 15.0.4615.1000 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/
/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/
var OSF = OSF || {};
OSF.OUtil = function () {
    var e = "on",
        h = "configurable",
        g = "writable",
        d = "enumerable",
        f = "undefined",
        c = true,
        b = null,
        a = false,
        k = -1,
        n = "&_xdm_Info=",
        m = "_xdm_",
        i = "#",
        j = {},
        p = 3e4,
        l = a;

    function o() {
        return Math.floor(100000001 * Math.random()).toString()
    }
    return {
        extend: function (b, a) {
            var c = function () {};
            c.prototype = a.prototype;
            b.prototype = new c;
            b.prototype.constructor = b;
            b.uber = a.prototype;
            if (a.prototype.constructor === Object.prototype.constructor) a.prototype.constructor = a
        },
        setNamespace: function (b, a) {
            if (a && b && !a[b]) a[b] = {}
        },
        unsetNamespace: function (b, a) {
            if (a && b && a[b]) delete a[b]
        },
        loadScript: function (f, g, h) {
            if (f && g) {
                var l = window.document,
                    d = j[f];
                if (!d) {
                    var e = l.createElement("script");
                    e.type = "text/javascript";
                    d = {
                        loaded: a,
                        pendingCallbacks: [g],
                        timer: b
                    };
                    j[f] = d;
                    var i = function () {
                            if (d.timer != b) {
                                clearTimeout(d.timer);
                                delete d.timer
                            }
                            d.loaded = c;
                            for (var e = d.pendingCallbacks.length, a = 0; a < e; a++) {
                                var f = d.pendingCallbacks.shift();
                                f()
                            }
                        },
                        k = function () {
                            delete j[f];
                            if (d.timer != b) {
                                clearTimeout(d.timer);
                                delete d.timer
                            }
                            for (var c = d.pendingCallbacks.length, a = 0; a < c; a++) {
                                var e = d.pendingCallbacks.shift();
                                e()
                            }
                        };
                    if (e.readyState) e.onreadystatechange = function () {
                        if (e.readyState == "loaded" || e.readyState == "complete") {
                            e.onreadystatechange = b;
                            i()
                        }
                    };
                    else e.onload = i;
                    e.onerror = k;
                    h = h || p;
                    d.timer = setTimeout(k, h);
                    e.src = f;
                    l.getElementsByTagName("head")[0].appendChild(e)
                } else if (d.loaded) g();
                else d.pendingCallbacks.push(g)
            }
        },
        loadCSS: function (c) {
            if (c) {
                var b = window.document,
                    a = b.createElement("link");
                a.type = "text/css";
                a.rel = "stylesheet";
                a.href = c;
                b.getElementsByTagName("head")[0].appendChild(a)
            }
        },
        parseEnum: function (b, c) {
            var a = c[b.trim()];
            if (typeof a == f) {
                Sys.Debug.trace("invalid enumeration string:" + b);
                throw Error.argument("str")
            }
            return a
        },
        delayExecutionAndCache: function () {
            var a = {
                calc: arguments[0]
            };
            return function () {
                if (a.calc) {
                    a.val = a.calc.apply(this, arguments);
                    delete a.calc
                }
                return a.val
            }
        },
        getUniqueId: function () {
            k = k + 1;
            return k.toString()
        },
        formatString: function () {
            var a = arguments,
                b = a[0];
            return b.replace(/{(\d+)}/gm, function (d, b) {
                var c = parseInt(b, 10) + 1;
                return a[c] === undefined ? "{" + b + "}" : a[c]
            })
        },
        generateConversationId: function () {
            return [o(), o(), (new Date).getTime().toString()].join("_")
        },
        getFrameNameAndConversationId: function (b, c) {
            var a = m + b + this.generateConversationId();
            c.setAttribute("name", a);
            return this.generateConversationId()
        },
        addXdmInfoAsHash: function (a, d) {
            a = a.trim() || "";
            var b = a.split(i),
                c = b.shift(),
                e = b.join(i);
            return [c, i, e, n, d].join("")
        },
        parseXdmInfo: function () {
            var g = window.location.hash,
                d = g.split(n),
                a = d.length > 1 ? d[d.length - 1] : b;
            if (window.sessionStorage) {
                var c = window.name.indexOf(m);
                if (c > -1) {
                    var e = window.name.indexOf(";", c);
                    if (e == -1) e = window.name.length;
                    var f = window.name.substring(c, e);
                    if (a) window.sessionStorage.setItem(f, a);
                    else a = window.sessionStorage.getItem(f)
                }
            }
            return a
        },
        getConversationId: function () {
            var c = window.location.search,
                a = b;
            if (c) {
                var d = c.indexOf("&");
                a = d > 0 ? c.substring(1, d) : c.substr(1);
                if (a && a.charAt(a.length - 1) === "=") {
                    a = a.substring(0, a.length - 1);
                    if (a) a = decodeURIComponent(a)
                }
            }
            return a
        },
        validateParamObject: function (f, e) {
            var b = Function._validateParams(arguments, [{
                name: "params",
                type: Object,
                mayBeNull: a
            }, {
                name: "expectedProperties",
                type: Object,
                mayBeNull: a
            }, {
                name: "callback",
                type: Function,
                mayBeNull: c
            }]);
            if (b) throw b;
            for (var d in e) {
                b = Function._validateParameter(f[d], e[d], d);
                if (b) throw b
            }
        },
        writeProfilerMark: function (a) {
            if (window.msWriteProfilerMark) {
                window.msWriteProfilerMark(a);
                typeof Sys !== f && Sys && Sys.Debug && Sys.Debug.trace(a)
            }
        },
        defineNondefaultProperty: function (e, f, a, b) {
            a = a || {};
            for (var g in b) {
                var d = b[g];
                if (a[d] == undefined) a[d] = c
            }
            Object.defineProperty(e, f, a);
            return e
        },
        defineNondefaultProperties: function (c, a, d) {
            a = a || {};
            for (var b in a) OSF.OUtil.defineNondefaultProperty(c, b, a[b], d);
            return c
        },
        defineEnumerableProperty: function (c, b, a) {
            return OSF.OUtil.defineNondefaultProperty(c, b, a, [d])
        },
        defineEnumerableProperties: function (b, a) {
            return OSF.OUtil.defineNondefaultProperties(b, a, [d])
        },
        defineMutableProperty: function (c, b, a) {
            return OSF.OUtil.defineNondefaultProperty(c, b, a, [g, d, h])
        },
        defineMutableProperties: function (b, a) {
            return OSF.OUtil.defineNondefaultProperties(b, a, [g, d, h])
        },
        finalizeProperties: function (e, d) {
            d = d || {};
            for (var g = Object.getOwnPropertyNames(e), i = g.length, f = 0; f < i; f++) {
                var h = g[f],
                    b = Object.getOwnPropertyDescriptor(e, h);
                if (!b.get && !b.set) b.writable = d.writable || a;
                b.configurable = d.configurable || a;
                b.enumerable = d.enumerable || c;
                Object.defineProperty(e, h, b)
            }
            return e
        },
        mapList: function (a, c) {
            var b = [];
            if (a)
                for (var d in a) b.push(c(a[d]));
            return b
        },
        listContainsKey: function (d, e) {
            for (var b in d)
                if (e == b) return c;
            return a
        },
        listContainsValue: function (b, d) {
            for (var e in b)
                if (d == b[e]) return c;
            return a
        },
        augmentList: function (a, b) {
            var d = a.push ? function (c, b) {
                a.push(b)
            } : function (c, b) {
                a[c] = b
            };
            for (var c in b) d(c, b[c])
        },
        redefineList: function (a, b) {
            for (var d in a) delete a[d];
            for (var c in b) a[c] = b[c]
        },
        isArray: function (a) {
            return Object.prototype.toString.apply(a) === "[object Array]"
        },
        isFunction: function (a) {
            return Object.prototype.toString.apply(a) === "[object Function]"
        },
        isDate: function (a) {
            return Object.prototype.toString.apply(a) === "[object Date]"
        },
        addEventListener: function (b, c, d) {
            if (b.attachEvent) b.attachEvent(e + c, d);
            else if (b.addEventListener) b.addEventListener(c, d, a);
            else b[e + c] = d
        },
        removeEventListener: function (c, d, f) {
            if (c.detachEvent) c.detachEvent(e + d, f);
            else if (c.removeEventListener) c.removeEventListener(d, f, a);
            else c[e + d] = b
        },
        encodeBase64: function (c) {
            var j = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",
                h = [],
                a = [],
                b = 0,
                f, d, e, i = c.length;
            do {
                f = c[b++];
                d = c[b++];
                e = c[b++];
                a[0] = f >> 2;
                a[1] = (f & 3) << 4 | d >> 4;
                a[2] = (d & 15) << 2 | e >> 6;
                a[3] = e & 63;
                if (isNaN(d)) a[2] = a[3] = 64;
                else if (isNaN(e)) a[3] = 64;
                for (var g = 0; g < 4; g++) h.push(j.charAt(a[g]))
            } while (b < i);
            return h.join("")
        },
        getLocalStorage: function () {
            var a = b;
            if (!l) try {
                if (window.localStorage) a = window.localStorage
            } catch (d) {
                l = c
            }
            return a
        },
        splitStringToList: function (j, h) {
            for (var e = a, i = -1, g = [], f = a, d = h + j, b = 0; b < d.length; b++)
                if (d[b] == "\\" && !e) e = c;
                else {
                    if (d[b] == h && !f) {
                        g.push("");
                        i++
                    } else if (d[b] == '"' && !e) f = !f;
                    else g[i] += d[b];
                    e = a
                }
            return g
        },
        convertIntToHex: function (b) {
            var a = "#" + (Number(b) + 16777216).toString(16).slice(-6);
            return a
        }
    }
}();
window.OSF = OSF;
OSF.OUtil.setNamespace("OSF", window);
OSF.InternalPerfMarker = {
    DataCoercionBegin: "Agave.HostCall.CoerceDataStart",
    DataCoercionEnd: "Agave.HostCall.CoerceDataEnd"
};
OSF.HostCallPerfMarker = {
    IssueCall: "Agave.HostCall.IssueCall",
    ReceiveResponse: "Agave.HostCall.ReceiveResponse",
    RuntimeExceptionRaised: "Agave.HostCall.RuntimeExecptionRaised"
};
OSF.AgaveHostAction = {
    Select: 0,
    UnSelect: 1
};
OSF.SharedConstants = {
    NotificationConversationIdSuffix: "_ntf"
};
OSF.OfficeAppContext = function (m, i, e, d, g, j, f, h, l, b, k, c) {
    var a = this;
    a._id = m;
    a._appName = i;
    a._appVersion = e;
    a._appUILocale = d;
    a._dataLocale = g;
    a._docUrl = j;
    a._clientMode = f;
    a._settings = h;
    a._reason = l;
    a._osfControlType = b;
    a._eToken = k;
    a._correlationId = c;
    a.get_id = function () {
        return this._id
    };
    a.get_appName = function () {
        return this._appName
    };
    a.get_appVersion = function () {
        return this._appVersion
    };
    a.get_appUILocale = function () {
        return this._appUILocale
    };
    a.get_dataLocale = function () {
        return this._dataLocale
    };
    a.get_docUrl = function () {
        return this._docUrl
    };
    a.get_clientMode = function () {
        return this._clientMode
    };
    a.get_bindings = function () {
        return this._bindings
    };
    a.get_settings = function () {
        return this._settings
    };
    a.get_reason = function () {
        return this._reason
    };
    a.get_osfControlType = function () {
        return this._osfControlType
    };
    a.get_eToken = function () {
        return this._eToken
    };
    a.get_correlationId = function () {
        return this._correlationId
    }
};
OSF.AppName = {
    Unsupported: 0,
    Excel: 1,
    Word: 2,
    PowerPoint: 4,
    Outlook: 8,
    ExcelWebApp: 16,
    WordWebApp: 32,
    OutlookWebApp: 64,
    Project: 128
};
OSF.OsfControlType = {
    DocumentLevel: 0,
    ContainerLevel: 1
};
OSF.ClientMode = {
    ReadOnly: 0,
    ReadWrite: 1
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Client", Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension", Microsoft.Office);
OSF.OUtil.setNamespace("Internal", Microsoft.Office);
OSF.NamespaceManager = function () {
    var b, a = false;
    return {
        enableShortcut: function () {
            if (!a) {
                if (window.Office) b = window.Office;
                else OSF.OUtil.setNamespace("Office", window);
                window.Office = Microsoft.Office.WebExtension;
                a = true
            }
        },
        disableShortcut: function () {
            if (a) {
                if (b) window.Office = b;
                else OSF.OUtil.unsetNamespace("Office", window);
                a = false
            }
        }
    }
}();
OSF.NamespaceManager.enableShortcut();
Microsoft.Office.WebExtension.InitializationReason = {
    Inserted: "inserted",
    DocumentOpened: "documentOpened"
};
Microsoft.Office.WebExtension.ApplicationMode = {
    WebEditor: "webEditor",
    WebViewer: "webViewer",
    Client: "client"
};
Microsoft.Office.WebExtension.DocumentMode = {
    ReadOnly: "readOnly",
    ReadWrite: "readWrite"
};
Microsoft.Office.WebExtension.CoercionType = {
    Text: "text",
    Matrix: "matrix",
    Table: "table"
};
Microsoft.Office.WebExtension.ValueFormat = {
    Unformatted: "unformatted",
    Formatted: "formatted"
};
Microsoft.Office.WebExtension.FilterType = {
    All: "all"
};
Microsoft.Office.WebExtension.BindingType = {
    Text: "text",
    Matrix: "matrix",
    Table: "table"
};
Microsoft.Office.WebExtension.GoToType = {
    Binding: "binding",
    NamedItem: "namedItem",
    Slide: "slide",
    Index: "index"
};
Microsoft.Office.WebExtension.SelectionMode = {
    Default: "default",
    Selected: "selected",
    None: "none"
};
Microsoft.Office.WebExtension.EventType = {
    DocumentSelectionChanged: "documentSelectionChanged",
    BindingSelectionChanged: "bindingSelectionChanged",
    BindingDataChanged: "bindingDataChanged"
};
Microsoft.Office.Internal.EventType = {
    OfficeThemeChanged: "officeThemeChanged",
    DocumentThemeChanged: "documentThemeChanged"
};
Microsoft.Office.WebExtension.AsyncResultStatus = {
    Succeeded: "succeeded",
    Failed: "failed"
};
Microsoft.Office.WebExtension.Parameters = {
    BindingType: "bindingType",
    CoercionType: "coercionType",
    ValueFormat: "valueFormat",
    FilterType: "filterType",
    Id: "id",
    GoToType: "goToType",
    SelectionMode: "selectionMode",
    PromptText: "promptText",
    ItemName: "itemName",
    FailOnCollision: "failOnCollision",
    StartRow: "startRow",
    StartColumn: "startColumn",
    RowCount: "rowCount",
    ColumnCount: "columnCount",
    Callback: "callback",
    AsyncContext: "asyncContext",
    Data: "data",
    Rows: "rows",
    OverwriteIfStale: "overwriteIfStale",
    FileType: "fileType",
    EventType: "eventType",
    Handler: "handler",
    SliceSize: "sliceSize",
    SliceIndex: "sliceIndex",
    ActiveView: "activeView",
    Xml: "xml",
    Namespace: "namespace",
    Prefix: "prefix",
    XPath: "xPath",
    TaskId: "taskId",
    FieldId: "fieldId",
    FieldValue: "fieldValue",
    ServerUrl: "serverUrl",
    ListName: "listName",
    ResourceId: "resourceId",
    ViewType: "viewType",
    ViewName: "viewName",
    GetRawValue: "getRawValue",
    CellFormat: "cellFormat",
    TableOptions: "tableOptions"
};
Microsoft.Office.Internal.Parameters = {
    DocumentTheme: "documentTheme",
    OfficeTheme: "officeTheme"
};
Microsoft.Office.WebExtension.DefaultParameterValues = {};
OSF.OUtil.setNamespace("DDA", OSF);
OSF.DDA.DocumentMode = {
    ReadOnly: 1,
    ReadWrite: 0
};
OSF.OUtil.setNamespace("DispIdHost", OSF.DDA);
OSF.DDA.DispIdHost.Methods = {
    InvokeMethod: "invokeMethod",
    AddEventHandler: "addEventHandler",
    RemoveEventHandler: "removeEventHandler"
};
OSF.DDA.DispIdHost.Delegates = {
    ExecuteAsync: "executeAsync",
    RegisterEventAsync: "registerEventAsync",
    UnregisterEventAsync: "unregisterEventAsync",
    ParameterMap: "parameterMap"
};
OSF.OUtil.setNamespace("AsyncResultEnum", OSF.DDA);
OSF.DDA.AsyncResultEnum.Properties = {
    Context: "Context",
    Value: "Value",
    Status: "Status",
    Error: "Error"
};
OSF.DDA.AsyncResultEnum.ErrorProperties = {
    Name: "Name",
    Message: "Message",
    Code: "Code"
};
OSF.DDA.PropertyDescriptors = {
    AsyncResultStatus: "AsyncResultStatus",
    FileProperties: "FileProperties",
    FilePropertiesDescriptor: "FilePropertiesDescriptor",
    FileSliceProperties: "FileSliceProperties",
    Subset: "subset",
    BindingProperties: "BindingProperties",
    TableDataProperties: "TableDataProperties",
    DataPartProperties: "DataPartProperties",
    DataNodeProperties: "DataNodeProperties"
};
OSF.DDA.EventDescriptors = {
    BindingSelectionChangedEvent: "BindingSelectionChangedEvent",
    DataNodeInsertedEvent: "DataNodeInsertedEvent",
    DataNodeReplacedEvent: "DataNodeReplacedEvent",
    DataNodeDeletedEvent: "DataNodeDeletedEvent",
    OfficeThemeChangedEvent: "OfficeThemeChangedEvent",
    DocumentThemeChangedEvent: "DocumentThemeChangedEvent",
    ActiveViewChangedEvent: "ActiveViewChangedEvent"
};
OSF.DDA.ListDescriptors = {
    BindingList: "BindingList",
    DataPartList: "DataPartList",
    DataNodeList: "DataNodeList"
};
OSF.DDA.FileProperties = {
    Handle: "FileHandle",
    FileSize: "FileSize",
    SliceSize: Microsoft.Office.WebExtension.Parameters.SliceSize
};
OSF.DDA.FilePropertiesDescriptor = {
    Url: "Url"
};
OSF.DDA.BindingProperties = {
    Id: "BindingId",
    Type: Microsoft.Office.WebExtension.Parameters.BindingType,
    RowCount: "BindingRowCount",
    ColumnCount: "BindingColumnCount",
    HasHeaders: "HasHeaders"
};
OSF.DDA.TableDataProperties = {
    TableRows: "TableRows",
    TableHeaders: "TableHeaders"
};
OSF.DDA.DataPartProperties = {
    Id: Microsoft.Office.WebExtension.Parameters.Id,
    BuiltIn: "DataPartBuiltIn"
};
OSF.DDA.DataNodeProperties = {
    Handle: "DataNodeHandle",
    BaseName: "DataNodeBaseName",
    NamespaceUri: "DataNodeNamespaceUri",
    NodeType: "DataNodeType"
};
OSF.DDA.DataNodeEventProperties = {
    OldNode: "OldNode",
    NewNode: "NewNode",
    NextSiblingNode: "NextSiblingNode",
    InUndoRedo: "InUndoRedo"
};
OSF.DDA.AsyncResultEnum.ErrorCode = {
    Success: 0,
    Failed: 1
};
OSF.DDA.getXdmEventName = function (b, a) {
    if (a == Microsoft.Office.WebExtension.EventType.BindingSelectionChanged || a == Microsoft.Office.WebExtension.EventType.BindingDataChanged) return b + "_" + a;
    else return a
};
OSF.DDA.ErrorCodeManager = function () {
    var a = {};
    return {
        getErrorArgs: function (b) {
            return a[b] || a[this.errorCodes.ooeInternalError]
        },
        addErrorMessage: function (c, b) {
            a[c] = b
        },
        errorCodes: {
            ooeSuccess: 0,
            ooeCoercionTypeNotSupported: 1e3,
            ooeGetSelectionNotMatchDataType: 1001,
            ooeCoercionTypeNotMatchBinding: 1002,
            ooeInvalidGetRowColumnCounts: 1003,
            ooeSelectionNotSupportCoercionType: 1004,
            ooeInvalidGetStartRowColumn: 1005,
            ooeNonUniformPartialGetNotSupported: 1006,
            ooeGetDataIsTooLarge: 1008,
            ooeFileTypeNotSupported: 1009,
            ooeUnsupportedDataObject: 2e3,
            ooeCannotWriteToSelection: 2001,
            ooeDataNotMatchSelection: 2002,
            ooeOverwriteWorksheetData: 2003,
            ooeDataNotMatchBindingSize: 2004,
            ooeInvalidSetStartRowColumn: 2005,
            ooeInvalidDataFormat: 2006,
            ooeDataNotMatchCoercionType: 2007,
            ooeDataNotMatchBindingType: 2008,
            ooeSetDataIsTooLarge: 2009,
            ooeNonUniformPartialSetNotSupported: 2010,
            ooeSelectionCannotBound: 3e3,
            ooeBindingNotExist: 3002,
            ooeBindingToMultipleSelection: 3003,
            ooeInvalidSelectionForBindingType: 3004,
            ooeOperationNotSupportedOnThisBindingType: 3005,
            ooeNamedItemNotFound: 3006,
            ooeMultipleNamedItemFound: 3007,
            ooeInvalidNamedItemForBindingType: 3008,
            ooeUnknownBindingType: 3009,
            ooeOperationNotSupportedOnMatrixData: 3010,
            ooeSettingNameNotExist: 4e3,
            ooeSettingsCannotSave: 4001,
            ooeSettingsAreStale: 4002,
            ooeOperationNotSupported: 5e3,
            ooeInternalError: 5001,
            ooeDocumentReadOnly: 5002,
            ooeEventHandlerNotExist: 5003,
            ooeInvalidApiCallInContext: 5004,
            ooeShuttingDown: 5005,
            ooeUnsupportedEnumeration: 5007,
            ooeIndexOutOfRange: 5008,
            ooeCustomXmlNodeNotFound: 6e3,
            ooeCustomXmlError: 6100,
            ooeNoCapability: 7e3,
            ooeCannotNavTo: 7001,
            ooeSpecifiedIdNotExist: 7002,
            ooeNavOutOfBound: 7004,
            ooeElementMissing: 8e3,
            ooeProtectedError: 8001,
            ooeInvalidCellsValue: 8010,
            ooeInvalidTableOptionValue: 8011,
            ooeInvalidFormatValue: 8012,
            ooeRowIndexOutOfRange: 8020,
            ooeColIndexOutOfRange: 8021,
            ooeFormatValueOutOfRange: 8022
        },
        initializeErrorMessages: function (b) {
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported] = {
                name: b.L_InvalidCoercion,
                message: b.L_CoercionTypeNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType] = {
                name: b.L_DataReadError,
                message: b.L_GetSelectionNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding] = {
                name: b.L_InvalidCoercion,
                message: b.L_CoercionTypeNotMatchBinding
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts] = {
                name: b.L_DataReadError,
                message: b.L_InvalidGetRowColumnCounts
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType] = {
                name: b.L_DataReadError,
                message: b.L_SelectionNotSupportCoercionType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn] = {
                name: b.L_DataReadError,
                message: b.L_InvalidGetStartRowColumn
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported] = {
                name: b.L_DataReadError,
                message: b.L_NonUniformPartialGetNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge] = {
                name: b.L_DataReadError,
                message: b.L_GetDataIsTooLarge
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported] = {
                name: b.L_DataReadError,
                message: b.L_FileTypeNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject] = {
                name: b.L_DataWriteError,
                message: b.L_UnsupportedDataObject
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection] = {
                name: b.L_DataWriteError,
                message: b.L_CannotWriteToSelection
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection] = {
                name: b.L_DataWriteError,
                message: b.L_DataNotMatchSelection
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData] = {
                name: b.L_DataWriteError,
                message: b.L_OverwriteWorksheetData
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize] = {
                name: b.L_DataWriteError,
                message: b.L_DataNotMatchBindingSize
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn] = {
                name: b.L_DataWriteError,
                message: b.L_InvalidSetStartRowColumn
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat] = {
                name: b.L_InvalidFormat,
                message: b.L_InvalidDataFormat
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType] = {
                name: b.L_InvalidDataObject,
                message: b.L_DataNotMatchCoercionType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType] = {
                name: b.L_InvalidDataObject,
                message: b.L_DataNotMatchBindingType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge] = {
                name: b.L_DataWriteError,
                message: b.L_SetDataIsTooLarge
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported] = {
                name: b.L_DataWriteError,
                message: b.L_NonUniformPartialSetNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound] = {
                name: b.L_BindingCreationError,
                message: b.L_SelectionCannotBound
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist] = {
                name: b.L_InvalidBindingError,
                message: b.L_BindingNotExist
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection] = {
                name: b.L_BindingCreationError,
                message: b.L_BindingToMultipleSelection
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType] = {
                name: b.L_BindingCreationError,
                message: b.L_InvalidSelectionForBindingType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType] = {
                name: b.L_InvalidBindingOperation,
                message: b.L_OperationNotSupportedOnThisBindingType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound] = {
                name: b.L_BindingCreationError,
                message: b.L_NamedItemNotFound
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound] = {
                name: b.L_BindingCreationError,
                message: b.L_MultipleNamedItemFound
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType] = {
                name: b.L_BindingCreationError,
                message: b.L_InvalidNamedItemForBindingType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType] = {
                name: b.L_InvalidBinding,
                message: b.L_UnknownBindingType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData] = {
                name: b.L_InvalidBindingOperation,
                message: b.L_OperationNotSupportedOnMatrixData
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist] = {
                name: b.L_ReadSettingsError,
                message: b.L_SettingNameNotExist
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave] = {
                name: b.L_SaveSettingsError,
                message: b.L_SettingsCannotSave
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale] = {
                name: b.L_SettingsStaleError,
                message: b.L_SettingsAreStale
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported] = {
                name: b.L_HostError,
                message: b.L_OperationNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError] = {
                name: b.L_InternalError,
                message: b.L_InternalErrorDescription
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly] = {
                name: b.L_PermissionDenied,
                message: b.L_DocumentReadOnly
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist] = {
                name: b.L_EventRegistrationError,
                message: b.L_EventHandlerNotExist
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext] = {
                name: b.L_InvalidAPICall,
                message: b.L_InvalidApiCallInContext
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown] = {
                name: b.L_ShuttingDown,
                message: b.L_ShuttingDown
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration] = {
                name: b.L_UnsupportedEnumeration,
                message: b.L_UnsupportedEnumerationMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange] = {
                name: b.L_IndexOutOfRange,
                message: b.L_IndexOutOfRange
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound] = {
                name: b.L_InvalidNode,
                message: b.L_CustomXmlNodeNotFound
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError] = {
                name: b.L_CustomXmlError,
                message: b.L_CustomXmlError
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability] = {
                name: b.L_PermissionDenied,
                message: b.L_NoCapability
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo] = {
                name: b.L_CannotNavigateTo,
                message: b.L_CannotNavigateTo
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist] = {
                name: b.L_SpecifiedIdNotExist,
                message: b.L_SpecifiedIdNotExist
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound] = {
                name: b.L_NavOutOfBound,
                message: b.L_NavOutOfBound
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing] = {
                name: b.L_MissingParameter,
                message: b.L_ElementMissing
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError] = {
                name: b.L_PermissionDenied,
                message: b.L_NoCapability
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue] = {
                name: b.L_InvalidValue,
                message: b.L_InvalidCellsValue
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue] = {
                name: b.L_InvalidValue,
                message: b.L_InvalidTableOptionValue
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue] = {
                name: b.L_InvalidValue,
                message: b.L_InvalidFormatValue
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange] = {
                name: b.L_OutOfRange,
                message: b.L_RowIndexOutOfRange
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange] = {
                name: b.L_OutOfRange,
                message: b.L_ColIndexOutOfRange
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange] = {
                name: b.L_OutOfRange,
                message: b.L_FormatValueOutOfRange
            }
        }
    }
}();
OSF.DDA.MethodDispId = {
    dispidMethodMin: 64,
    dispidGetSelectedDataMethod: 64,
    dispidSetSelectedDataMethod: 65,
    dispidAddBindingFromSelectionMethod: 66,
    dispidAddBindingFromPromptMethod: 67,
    dispidGetBindingMethod: 68,
    dispidReleaseBindingMethod: 69,
    dispidGetBindingDataMethod: 70,
    dispidSetBindingDataMethod: 71,
    dispidAddRowsMethod: 72,
    dispidClearAllRowsMethod: 73,
    dispidGetAllBindingsMethod: 74,
    dispidLoadSettingsMethod: 75,
    dispidSaveSettingsMethod: 76,
    dispidGetDocumentCopyMethod: 77,
    dispidAddBindingFromNamedItemMethod: 78,
    dispidAddColumnsMethod: 79,
    dispidGetDocumentCopyChunkMethod: 80,
    dispidReleaseDocumentCopyMethod: 81,
    dispidNavigateToMethod: 82,
    dispidGetActiveViewMethod: 83,
    dispidGetDocumentThemeMethod: 84,
    dispidGetOfficeThemeMethod: 85,
    dispidGetFilePropertiesMethod: 86,
    dispidClearFormatsMethod: 87,
    dispidSetTableOptionsMethod: 88,
    dispidSetFormatsMethod: 89,
    dispidAddDataPartMethod: 128,
    dispidGetDataPartByIdMethod: 129,
    dispidGetDataPartsByNamespaceMethod: 130,
    dispidGetDataPartXmlMethod: 131,
    dispidGetDataPartNodesMethod: 132,
    dispidDeleteDataPartMethod: 133,
    dispidGetDataNodeValueMethod: 134,
    dispidGetDataNodeXmlMethod: 135,
    dispidGetDataNodesMethod: 136,
    dispidSetDataNodeValueMethod: 137,
    dispidSetDataNodeXmlMethod: 138,
    dispidAddDataNamespaceMethod: 139,
    dispidGetDataUriByPrefixMethod: 140,
    dispidGetDataPrefixByUriMethod: 141,
    dispidMethodMax: 141,
    dispidGetSelectedTaskMethod: 110,
    dispidGetSelectedResourceMethod: 111,
    dispidGetTaskMethod: 112,
    dispidGetResourceFieldMethod: 113,
    dispidGetWSSUrlMethod: 114,
    dispidGetTaskFieldMethod: 115,
    dispidGetProjectFieldMethod: 116,
    dispidGetSelectedViewMethod: 117
};
OSF.DDA.EventDispId = {
    dispidEventMin: 0,
    dispidInitializeEvent: 0,
    dispidSettingsChangedEvent: 1,
    dispidDocumentSelectionChangedEvent: 2,
    dispidBindingSelectionChangedEvent: 3,
    dispidBindingDataChangedEvent: 4,
    dispidDocumentOpenEvent: 5,
    dispidDocumentCloseEvent: 6,
    dispidActiveViewChangedEvent: 7,
    dispidDocumentThemeChangedEvent: 8,
    dispidOfficeThemeChangedEvent: 9,
    dispidDataNodeAddedEvent: 60,
    dispidDataNodeReplacedEvent: 61,
    dispidDataNodeDeletedEvent: 62,
    dispidEventMax: 63,
    dispidTaskSelectionChangedEvent: 56,
    dispidResourceSelectionChangedEvent: 57,
    dispidViewSelectionChangedEvent: 58
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Common", Microsoft.Office);
Microsoft.Office.Common.InvokeType = {
    async: 0,
    sync: 1,
    asyncRegisterEvent: 2,
    asyncUnregisterEvent: 3,
    syncRegisterEvent: 4,
    syncUnregisterEvent: 5
};
Microsoft.Office.Common.InvokeResultCode = {
    noError: 0,
    errorInRequest: -1,
    errorHandlingRequest: -2,
    errorInResponse: -3,
    errorHandlingResponse: -4,
    errorHandlingRequestAccessDenied: -5,
    errorHandlingMethodCallTimedout: -6
};
Microsoft.Office.Common.MessageType = {
    request: 0,
    response: 1
};
Microsoft.Office.Common.ActionType = {
    invoke: 0,
    registerEvent: 1,
    unregisterEvent: 2
};
Microsoft.Office.Common.ResponseType = {
    forCalling: 0,
    forEventing: 1
};
Microsoft.Office.Common.MethodObject = function (c, b, a) {
    this._method = c;
    this._invokeType = b;
    this._blockingOthers = a
};
Microsoft.Office.Common.MethodObject.prototype = {
    getMethod: function () {
        return this._method
    },
    getInvokeType: function () {
        return this._invokeType
    },
    getBlockingFlag: function () {
        return this._blockingOthers
    }
};
Microsoft.Office.Common.EventMethodObject = function (b, a) {
    this._registerMethodObject = b;
    this._unregisterMethodObject = a
};
Microsoft.Office.Common.EventMethodObject.prototype = {
    getRegisterMethodObject: function () {
        return this._registerMethodObject
    },
    getUnregisterMethodObject: function () {
        return this._unregisterMethodObject
    }
};
Microsoft.Office.Common.ServiceEndPoint = function (c) {
    var a = this,
        b = Function._validateParams(arguments, [{
            name: "serviceEndPointId",
            type: String,
            mayBeNull: false
        }]);
    if (b) throw b;
    a._methodObjectList = {};
    a._eventHandlerProxyList = {};
    a._Id = c;
    a._conversations = {};
    a._policyManager = null
};
Microsoft.Office.Common.ServiceEndPoint.prototype = {
    registerMethod: function (g, h, b, e) {
        var c = "invokeType",
            a = false,
            d = Function._validateParams(arguments, [{
                name: "methodName",
                type: String,
                mayBeNull: a
            }, {
                name: "method",
                type: Function,
                mayBeNull: a
            }, {
                name: c,
                type: Number,
                mayBeNull: a
            }, {
                name: "blockingOthers",
                type: Boolean,
                mayBeNull: a
            }]);
        if (d) throw d;
        if (b !== Microsoft.Office.Common.InvokeType.async && b !== Microsoft.Office.Common.InvokeType.sync) throw Error.argument(c);
        var f = new Microsoft.Office.Common.MethodObject(h, b, e);
        this._methodObjectList[g] = f
    },
    unregisterMethod: function (b) {
        var a = Function._validateParams(arguments, [{
            name: "methodName",
            type: String,
            mayBeNull: false
        }]);
        if (a) throw a;
        delete this._methodObjectList[b]
    },
    registerEvent: function (f, d, c) {
        var a = false,
            b = Function._validateParams(arguments, [{
                name: "eventName",
                type: String,
                mayBeNull: a
            }, {
                name: "registerMethod",
                type: Function,
                mayBeNull: a
            }, {
                name: "unregisterMethod",
                type: Function,
                mayBeNull: a
            }]);
        if (b) throw b;
        var e = new Microsoft.Office.Common.EventMethodObject(new Microsoft.Office.Common.MethodObject(d, Microsoft.Office.Common.InvokeType.syncRegisterEvent, a), new Microsoft.Office.Common.MethodObject(c, Microsoft.Office.Common.InvokeType.syncUnregisterEvent, a));
        this._methodObjectList[f] = e
    },
    registerEventEx: function (h, f, d, e, c) {
        var a = false,
            b = Function._validateParams(arguments, [{
                name: "eventName",
                type: String,
                mayBeNull: a
            }, {
                name: "registerMethod",
                type: Function,
                mayBeNull: a
            }, {
                name: "registerMethodInvokeType",
                type: Number,
                mayBeNull: a
            }, {
                name: "unregisterMethod",
                type: Function,
                mayBeNull: a
            }, {
                name: "unregisterMethodInvokeType",
                type: Number,
                mayBeNull: a
            }]);
        if (b) throw b;
        var g = new Microsoft.Office.Common.EventMethodObject(new Microsoft.Office.Common.MethodObject(f, d, a), new Microsoft.Office.Common.MethodObject(e, c, a));
        this._methodObjectList[h] = g
    },
    unregisterEvent: function (b) {
        var a = Function._validateParams(arguments, [{
            name: "eventName",
            type: String,
            mayBeNull: false
        }]);
        if (a) throw a;
        this.unregisterMethod(b)
    },
    registerConversation: function (b) {
        var a = Function._validateParams(arguments, [{
            name: "conversationId",
            type: String,
            mayBeNull: false
        }]);
        if (a) throw a;
        this._conversations[b] = true
    },
    unregisterConversation: function (b) {
        var a = Function._validateParams(arguments, [{
            name: "conversationId",
            type: String,
            mayBeNull: false
        }]);
        if (a) throw a;
        delete this._conversations[b]
    },
    setPolicyManager: function (a) {
        var b = "policyManager",
            c = Function._validateParams(arguments, [{
                name: b,
                type: Object,
                mayBeNull: false
            }]);
        if (c) throw c;
        if (!a.checkPermission) throw Error.argument(b);
        this._policyManager = a
    },
    getPolicyManager: function () {
        return this._policyManager
    }
};
Microsoft.Office.Common.ClientEndPoint = function (e, b, f) {
    var c = "targetWindow",
        a = this,
        d = Function._validateParams(arguments, [{
            name: "conversationId",
            type: String,
            mayBeNull: false
        }, {
            name: c,
            mayBeNull: false
        }, {
            name: "targetUrl",
            type: String,
            mayBeNull: false
        }]);
    if (d) throw d;
    if (!b.postMessage) throw Error.argument(c);
    a._conversationId = e;
    a._targetWindow = b;
    a._targetUrl = f;
    a._callingIndex = 0;
    a._callbackList = {};
    a._eventHandlerList = {}
};
Microsoft.Office.Common.ClientEndPoint.prototype = {
    invoke: function (h, d, b) {
        var a = this,
            g = Function._validateParams(arguments, [{
                name: "targetMethodName",
                type: String,
                mayBeNull: false
            }, {
                name: "callback",
                type: Function,
                mayBeNull: true
            }, {
                name: "param",
                mayBeNull: true
            }]);
        if (g) throw g;
        var c = a._callingIndex++,
            k = new Date,
            e = {
                callback: d,
                createdOn: k.getTime()
            };
        if (b && typeof b === "object" && typeof b.__timeout__ === "number") {
            e.timeout = b.__timeout__;
            delete b.__timeout__
        }
        a._callbackList[c] = e;
        try {
            var i = new Microsoft.Office.Common.Request(h, Microsoft.Office.Common.ActionType.invoke, a._conversationId, c, b),
                j = Microsoft.Office.Common.MessagePackager.envelope(i);
            a._targetWindow.postMessage(j, a._targetUrl);
            Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer()
        } catch (f) {
            try {
                d !== null && d(Microsoft.Office.Common.InvokeResultCode.errorInRequest, f)
            } finally {
                delete a._callbackList[c]
            }
        }
    },
    registerForEvent: function (d, g, c, i) {
        var a = this,
            f = Function._validateParams(arguments, [{
                name: "targetEventName",
                type: String,
                mayBeNull: false
            }, {
                name: "eventHandler",
                type: Function,
                mayBeNull: false
            }, {
                name: "callback",
                type: Function,
                mayBeNull: true
            }, {
                name: "data",
                mayBeNull: true,
                optional: true
            }]);
        if (f) throw f;
        var b = a._callingIndex++,
            k = new Date;
        a._callbackList[b] = {
            callback: c,
            createdOn: k.getTime()
        };
        try {
            var h = new Microsoft.Office.Common.Request(d, Microsoft.Office.Common.ActionType.registerEvent, a._conversationId, b, i),
                j = Microsoft.Office.Common.MessagePackager.envelope(h);
            a._targetWindow.postMessage(j, a._targetUrl);
            Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();
            a._eventHandlerList[d] = g
        } catch (e) {
            try {
                c !== null && c(Microsoft.Office.Common.InvokeResultCode.errorInRequest, e)
            } finally {
                delete a._callbackList[b]
            }
        }
    },
    unregisterForEvent: function (d, c, h) {
        var a = this,
            f = Function._validateParams(arguments, [{
                name: "targetEventName",
                type: String,
                mayBeNull: false
            }, {
                name: "callback",
                type: Function,
                mayBeNull: true
            }, {
                name: "data",
                mayBeNull: true,
                optional: true
            }]);
        if (f) throw f;
        var b = a._callingIndex++,
            j = new Date;
        a._callbackList[b] = {
            callback: c,
            createdOn: j.getTime()
        };
        try {
            var g = new Microsoft.Office.Common.Request(d, Microsoft.Office.Common.ActionType.unregisterEvent, a._conversationId, b, h),
                i = Microsoft.Office.Common.MessagePackager.envelope(g);
            a._targetWindow.postMessage(i, a._targetUrl);
            Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer()
        } catch (e) {
            try {
                c !== null && c(Microsoft.Office.Common.InvokeResultCode.errorInRequest, e)
            } finally {
                delete a._callbackList[b]
            }
        } finally {
            delete a._eventHandlerList[d]
        }
    }
};
Microsoft.Office.Common.XdmCommunicationManager = function () {
    var i = "channel is not ready.",
        c = "conversationId",
        h = "Unknown conversation Id.",
        b = false,
        a = null,
        k = [],
        e = a,
        v = 10,
        j = b,
        f = a,
        o = 2e3,
        l = 6e4,
        g = {},
        d = {},
        m = b;

    function p(b) {
        for (var a in g)
            if (g[a]._conversations[b]) return g[a];
        Sys.Debug.trace(h);
        throw Error.argument(c)
    }

    function q(b) {
        var a = d[b];
        if (!a) {
            Sys.Debug.trace(h);
            throw Error.argument(c)
        }
        return a
    }

    function t(e, c) {
        var b = e._methodObjectList[c._actionName];
        if (!b) {
            Sys.Debug.trace("The specified method is not registered on service endpoint:" + c._actionName);
            throw Error.argument("messageObject")
        }
        var d = a;
        if (c._actionType === Microsoft.Office.Common.ActionType.invoke) d = b;
        else if (c._actionType === Microsoft.Office.Common.ActionType.registerEvent) d = b.getRegisterMethodObject();
        else d = b.getUnregisterMethodObject();
        return d
    }

    function x(a) {
        k.push(a)
    }

    function w() {
        if (e !== a) {
            if (!j)
                if (k.length > 0) {
                    var b = k.shift();
                    j = b.getInvokeBlockingFlag();
                    b.invoke()
                } else {
                    clearInterval(e);
                    e = a
                }
        } else Sys.Debug.trace(i)
    }

    function s() {
        if (f) {
            var c, e = 0,
                k = new Date,
                h;
            for (var j in d) {
                c = d[j];
                for (var g in c._callbackList) {
                    var b = c._callbackList[g];
                    h = b.timeout ? b.timeout : l;
                    if (Math.abs(k.getTime() - b.createdOn) >= h) try {
                        b.callback && b.callback(Microsoft.Office.Common.InvokeResultCode.errorHandlingMethodCallTimedout, a)
                    } finally {
                        delete c._callbackList[g]
                    } else e++
                }
            }
            if (e === 0) {
                clearInterval(f);
                f = a
            }
        } else Sys.Debug.trace(i)
    }

    function r() {
        j = b
    }

    function u(a) {
        if (Sys.Browser.agent === Sys.Browser.InternetExplorer && window.attachEvent) window.attachEvent("onmessage", a);
        else if (window.addEventListener) window.addEventListener("message", a, b);
        else {
            Sys.Debug.trace("Browser doesn't support the required API.");
            throw Error.argument("Browser")
        }
    }

    function y(c) {
        var d = "Access Denied";
        if (c.data != "") {
            var b;
            try {
                b = Microsoft.Office.Common.MessagePackager.unenvelope(c.data)
            } catch (f) {
                return
            }
            if (typeof b._messageType == "undefined") return;
            if (b._messageType === Microsoft.Office.Common.MessageType.request) {
                var l = c.origin == a || c.origin == "null" ? b._origin : c.origin;
                try {
                    var g = p(b._conversationId),
                        k = g.getPolicyManager();
                    if (k && !k.checkPermission(b._conversationId, b._actionName, b._data)) throw d;
                    var u = t(g, b),
                        n = new Microsoft.Office.Common.InvokeCompleteCallback(c.source, l, b._actionName, b._conversationId, b._correlationId, r),
                        y = new Microsoft.Office.Common.Invoker(u, b._data, n, g._eventHandlerProxyList, b._conversationId, b._actionName);
                    if (e == a) e = setInterval(w, v);
                    x(y)
                } catch (f) {
                    var m = Microsoft.Office.Common.InvokeResultCode.errorHandlingRequest;
                    if (f == d) m = Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied;
                    var s = new Microsoft.Office.Common.Response(b._actionName, b._conversationId, b._correlationId, m, Microsoft.Office.Common.ResponseType.forCalling, f),
                        o = Microsoft.Office.Common.MessagePackager.envelope(s);
                    c.source && c.source.postMessage && c.source.postMessage(o, l)
                }
            } else if (b._messageType === Microsoft.Office.Common.MessageType.response) {
                var h = q(b._conversationId);
                if (b._responseType === Microsoft.Office.Common.ResponseType.forCalling) {
                    var i = h._callbackList[b._correlationId];
                    if (i) try {
                        i.callback && i.callback(b._errorCode, b._data)
                    } finally {
                        delete h._callbackList[b._correlationId]
                    }
                } else {
                    var j = h._eventHandlerList[b._actionName];
                    j !== undefined && j !== a && j(b._data)
                }
            } else return
        }
    }

    function n() {
        if (!m) {
            u(y);
            m = true
        }
    }
    return {
        connect: function (b, c, e) {
            var a = d[b];
            if (!a) {
                n();
                a = new Microsoft.Office.Common.ClientEndPoint(b, c, e);
                d[b] = a
            }
            return a
        },
        getClientEndPoint: function (e) {
            var a = Function._validateParams(arguments, [{
                name: c,
                type: String,
                mayBeNull: b
            }]);
            if (a) throw a;
            return d[e]
        },
        createServiceEndPoint: function (a) {
            n();
            var b = new Microsoft.Office.Common.ServiceEndPoint(a);
            g[a] = b;
            return b
        },
        getServiceEndPoint: function (c) {
            var a = Function._validateParams(arguments, [{
                name: "serviceEndPointId",
                type: String,
                mayBeNull: b
            }]);
            if (a) throw a;
            return g[c]
        },
        deleteClientEndPoint: function (e) {
            var a = Function._validateParams(arguments, [{
                name: c,
                type: String,
                mayBeNull: b
            }]);
            if (a) throw a;
            delete d[e]
        },
        _setMethodTimeout: function (a) {
            var c = Function._validateParams(arguments, [{
                name: "methodTimeout",
                type: Number,
                mayBeNull: b
            }]);
            if (c) throw c;
            l = a <= 0 ? 6e4 : a
        },
        _startMethodTimeoutTimer: function () {
            if (!f) f = setInterval(s, o)
        }
    }
}();
Microsoft.Office.Common.Message = function (g, h, e, f, c) {
    var b = false,
        a = this,
        d = Function._validateParams(arguments, [{
            name: "messageType",
            type: Number,
            mayBeNull: b
        }, {
            name: "actionName",
            type: String,
            mayBeNull: b
        }, {
            name: "conversationId",
            type: String,
            mayBeNull: b
        }, {
            name: "correlationId",
            mayBeNull: b
        }, {
            name: "data",
            mayBeNull: true,
            optional: true
        }]);
    if (d) throw d;
    a._messageType = g;
    a._actionName = h;
    a._conversationId = e;
    a._correlationId = f;
    a._origin = window.location.href;
    if (typeof c == "undefined") a._data = null;
    else a._data = c
};
Microsoft.Office.Common.Message.prototype = {
    getActionName: function () {
        return this._actionName
    },
    getConversationId: function () {
        return this._conversationId
    },
    getCorrelationId: function () {
        return this._correlationId
    },
    getOrigin: function () {
        return this._origin
    },
    getData: function () {
        return this._data
    },
    getMessageType: function () {
        return this._messageType
    }
};
Microsoft.Office.Common.Request = function (c, d, a, b, e) {
    Microsoft.Office.Common.Request.uber.constructor.call(this, Microsoft.Office.Common.MessageType.request, c, a, b, e);
    this._actionType = d
};
OSF.OUtil.extend(Microsoft.Office.Common.Request, Microsoft.Office.Common.Message);
Microsoft.Office.Common.Request.prototype.getActionType = function () {
    return this._actionType
};
Microsoft.Office.Common.Response = function (d, a, b, e, c, f) {
    Microsoft.Office.Common.Response.uber.constructor.call(this, Microsoft.Office.Common.MessageType.response, d, a, b, f);
    this._errorCode = e;
    this._responseType = c
};
OSF.OUtil.extend(Microsoft.Office.Common.Response, Microsoft.Office.Common.Message);
Microsoft.Office.Common.Response.prototype.getErrorCode = function () {
    return this._errorCode
};
Microsoft.Office.Common.Response.prototype.getResponseType = function () {
    return this._responseType
};
Microsoft.Office.Common.MessagePackager = {
    envelope: function (a) {
        return Sys.Serialization.JavaScriptSerializer.serialize(a)
    },
    unenvelope: function (a) {
        return Sys.Serialization.JavaScriptSerializer.deserialize(a, true)
    }
};
Microsoft.Office.Common.ResponseSender = function (e, h, j, f, g, i) {
    var c = false,
        a = this,
        d = Function._validateParams(arguments, [{
            name: "requesterWindow",
            mayBeNull: c
        }, {
            name: "requesterUrl",
            type: String,
            mayBeNull: c
        }, {
            name: "actionName",
            type: String,
            mayBeNull: c
        }, {
            name: "conversationId",
            type: String,
            mayBeNull: c
        }, {
            name: "correlationId",
            mayBeNull: c
        }, {
            name: "responsetype",
            type: Number,
            maybeNull: c
        }]);
    if (d) throw d;
    a._requesterWindow = e;
    a._requesterUrl = h;
    a._actionName = j;
    a._conversationId = f;
    a._correlationId = g;
    a._invokeResultCode = Microsoft.Office.Common.InvokeResultCode.noError;
    a._responseType = i;
    var b = a;
    a._send = function (d) {
        var c = new Microsoft.Office.Common.Response(b._actionName, b._conversationId, b._correlationId, b._invokeResultCode, b._responseType, d),
            a = Microsoft.Office.Common.MessagePackager.envelope(c);
        b._requesterWindow.postMessage(a, b._requesterUrl)
    }
};
Microsoft.Office.Common.ResponseSender.prototype = {
    getRequesterWindow: function () {
        return this._requesterWindow
    },
    getRequesterUrl: function () {
        return this._requesterUrl
    },
    getActionName: function () {
        return this._actionName
    },
    getConversationId: function () {
        return this._conversationId
    },
    getCorrelationId: function () {
        return this._correlationId
    },
    getSend: function () {
        return this._send
    },
    setResultCode: function (a) {
        this._invokeResultCode = a
    }
};
Microsoft.Office.Common.InvokeCompleteCallback = function (d, g, h, e, f, c) {
    var b = this;
    Microsoft.Office.Common.InvokeCompleteCallback.uber.constructor.call(b, d, g, h, e, f, Microsoft.Office.Common.ResponseType.forCalling);
    b._postCallbackHandler = c;
    var a = b;
    b._send = function (d) {
        var c = new Microsoft.Office.Common.Response(a._actionName, a._conversationId, a._correlationId, a._invokeResultCode, a._responseType, d),
            b = Microsoft.Office.Common.MessagePackager.envelope(c);
        a._requesterWindow.postMessage(b, a._requesterUrl);
        a._postCallbackHandler()
    }
};
OSF.OUtil.extend(Microsoft.Office.Common.InvokeCompleteCallback, Microsoft.Office.Common.ResponseSender);
Microsoft.Office.Common.Invoker = function (g, h, d, e, f, i) {
    var b = false,
        a = this,
        c = Function._validateParams(arguments, [{
            name: "methodObject",
            mayBeNull: b
        }, {
            name: "paramValue",
            mayBeNull: true
        }, {
            name: "invokeCompleteCallback",
            mayBeNull: b
        }, {
            name: "eventHandlerProxyList",
            mayBeNull: true
        }, {
            name: "conversationId",
            type: String,
            mayBeNull: b
        }, {
            name: "eventName",
            type: String,
            mayBeNull: b
        }]);
    if (c) throw c;
    a._methodObject = g;
    a._param = h;
    a._invokeCompleteCallback = d;
    a._eventHandlerProxyList = e;
    a._conversationId = f;
    a._eventName = i
};
Microsoft.Office.Common.Invoker.prototype = {
    invoke: function () {
        var a = this;
        try {
            var b;
            switch (a._methodObject.getInvokeType()) {
            case Microsoft.Office.Common.InvokeType.async:
                a._methodObject.getMethod()(a._param, a._invokeCompleteCallback.getSend());
                break;
            case Microsoft.Office.Common.InvokeType.sync:
                b = a._methodObject.getMethod()(a._param);
                a._invokeCompleteCallback.getSend()(b);
                break;
            case Microsoft.Office.Common.InvokeType.syncRegisterEvent:
                var d = a._createEventHandlerProxyObject(a._invokeCompleteCallback);
                b = a._methodObject.getMethod()(d.getSend(), a._param);
                a._eventHandlerProxyList[a._conversationId + a._eventName] = d.getSend();
                a._invokeCompleteCallback.getSend()(b);
                break;
            case Microsoft.Office.Common.InvokeType.syncUnregisterEvent:
                var g = a._eventHandlerProxyList[a._conversationId + a._eventName];
                b = a._methodObject.getMethod()(g, a._param);
                delete a._eventHandlerProxyList[a._conversationId + a._eventName];
                a._invokeCompleteCallback.getSend()(b);
                break;
            case Microsoft.Office.Common.InvokeType.asyncRegisterEvent:
                var c = a._createEventHandlerProxyObject(a._invokeCompleteCallback);
                a._methodObject.getMethod()(c.getSend(), a._invokeCompleteCallback.getSend(), a._param);
                a._eventHandlerProxyList[a._callerId + a._eventName] = c.getSend();
                break;
            case Microsoft.Office.Common.InvokeType.asyncUnregisterEvent:
                var f = a._eventHandlerProxyList[a._callerId + a._eventName];
                a._methodObject.getMethod()(f, a._invokeCompleteCallback.getSend(), a._param);
                delete a._eventHandlerProxyList[a._callerId + a._eventName]
            }
        } catch (e) {
            a._invokeCompleteCallback.setResultCode(Microsoft.Office.Common.InvokeResultCode.errorInResponse);
            a._invokeCompleteCallback.getSend()(e)
        }
    },
    getInvokeBlockingFlag: function () {
        return this._methodObject.getBlockingFlag()
    },
    _createEventHandlerProxyObject: function (a) {
        return new Microsoft.Office.Common.ResponseSender(a.getRequesterWindow(), a.getRequesterUrl(), a.getActionName(), a.getConversationId(), a.getCorrelationId(), Microsoft.Office.Common.ResponseType.forEventing)
    }
};
(function () {
    var a = "undefined",
        c = function () {
            var d = function (a) {
                    a && OSF.OUtil.loadScript(a, function () {
                        Sys.Debug.trace("loaded customized script:" + a)
                    })
                },
                f, h, b, c = null,
                g = OSF.OUtil.parseXdmInfo();
            if (g) {
                b = g.split("|");
                if (b && b.length >= 3) {
                    f = b[0];
                    h = b[2];
                    c = Microsoft.Office.Common.XdmCommunicationManager.connect(f, window.parent, h)
                }
            }
            var e = null;
            if (!c) {
                try {
                    if (typeof window.external.getCustomizedScriptPath !== a) e = window.external.getCustomizedScriptPath()
                } catch (i) {
                    Sys.Debug.trace("no script override through window.external.")
                }
                d(e)
            } else try {
                c.invoke("getCustomizedScriptPathAsync", function (b, a) {
                    d(b === 0 ? a : null)
                }, {
                    __timeout__: 1e3
                })
            } catch (i) {
                Sys.Debug.trace("no script override through cross frame communication.")
            }
        },
        b = function () {
            var b = "function";
            if (typeof Sys !== a && typeof Type !== a && Sys.StringBuilder && typeof Sys.StringBuilder === b && Type.registerNamespace && typeof Type.registerNamespace === b && Type.registerClass && typeof Type.registerClass === b) return true;
            else return false
        };
    if (b()) c();
    else if (typeof Function !== a) {
        var d = (window.location.protocol.toLowerCase() === "https:" ? "https:" : "http:") + "//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";
        OSF.OUtil.loadScript(d, function () {
            if (b()) c();
            else if (typeof Function !== a) throw "Not able to load MicrosoftAjax.js."
        })
    }
})();
OSF.EventDispatch = function (a) {
    this._eventHandlers = {};
    for (var c in a) {
        var b = a[c];
        this._eventHandlers[b] = []
    }
};
OSF.EventDispatch.prototype = {
    getSupportedEvents: function () {
        var a = [];
        for (var b in this._eventHandlers) a.push(b);
        return a
    },
    supportsEvent: function (c) {
        var a = false;
        for (var b in this._eventHandlers)
            if (c == b) {
                a = true;
                break
            }
        return a
    },
    hasEventHandler: function (b, c) {
        var a = this._eventHandlers[b];
        if (a && a.length > 0)
            for (var d in a)
                if (a[d] === c) return true;
        return false
    },
    addEventHandler: function (b, a) {
        if (typeof a != "function") return false;
        var c = this._eventHandlers[b];
        if (c && !this.hasEventHandler(b, a)) {
            c.push(a);
            return true
        } else return false
    },
    removeEventHandler: function (c, d) {
        var a = this._eventHandlers[c];
        if (a && a.length > 0)
            for (var b = 0; b < a.length; b++)
                if (a[b] === d) {
                    a.splice(b, 1);
                    return true
                }
        return false
    },
    clearEventHandlers: function (a) {
        this._eventHandlers[a] = []
    },
    getEventHandlerCount: function (a) {
        return this._eventHandlers[a] != undefined ? this._eventHandlers[a].length : -1
    },
    fireEvent: function (a) {
        if (a.type == undefined) return false;
        var b = a.type;
        if (b && this._eventHandlers[b]) {
            var c = this._eventHandlers[b];
            for (var d in c) c[d](a);
            return true
        } else return false
    }
};
OSF.DDA.DataCoercion = function () {
    var a = null;
    return {
        findArrayDimensionality: function (c) {
            if (OSF.OUtil.isArray(c)) {
                for (var b = 0, a = 0; a < c.length; a++) b = Math.max(b, OSF.DDA.DataCoercion.findArrayDimensionality(c[a]));
                return b + 1
            } else return 0
        },
        getCoercionDefaultForBinding: function (a) {
            switch (a) {
            case Microsoft.Office.WebExtension.BindingType.Matrix:
                return Microsoft.Office.WebExtension.CoercionType.Matrix;
            case Microsoft.Office.WebExtension.BindingType.Table:
                return Microsoft.Office.WebExtension.CoercionType.Table;
            case Microsoft.Office.WebExtension.BindingType.Text:
            default:
                return Microsoft.Office.WebExtension.CoercionType.Text
            }
        },
        getBindingDefaultForCoercion: function (a) {
            switch (a) {
            case Microsoft.Office.WebExtension.CoercionType.Matrix:
                return Microsoft.Office.WebExtension.BindingType.Matrix;
            case Microsoft.Office.WebExtension.CoercionType.Table:
                return Microsoft.Office.WebExtension.BindingType.Table;
            case Microsoft.Office.WebExtension.CoercionType.Text:
            case Microsoft.Office.WebExtension.CoercionType.Html:
            case Microsoft.Office.WebExtension.CoercionType.Ooxml:
            default:
                return Microsoft.Office.WebExtension.BindingType.Text
            }
        },
        determineCoercionType: function (b) {
            if (b == a || b == undefined) return a;
            var c = a,
                d = typeof b;
            if (b.rows !== undefined) c = Microsoft.Office.WebExtension.CoercionType.Table;
            else if (OSF.OUtil.isArray(b)) c = Microsoft.Office.WebExtension.CoercionType.Matrix;
            else if (d == "string" || d == "number" || d == "boolean" || OSF.OUtil.isDate(b)) c = Microsoft.Office.WebExtension.CoercionType.Text;
            else throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject;
            return c
        },
        coerceData: function (b, c, a) {
            a = a || OSF.DDA.DataCoercion.determineCoercionType(b);
            if (a && a != c) {
                OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionBegin);
                b = OSF.DDA.DataCoercion._coerceDataFromTable(c, OSF.DDA.DataCoercion._coerceDataToTable(b, a));
                OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionEnd)
            }
            return b
        },
        _matrixToText: function (a) {
            if (a.length == 1 && a[0].length == 1) return "" + a[0][0];
            for (var b = "", c = 0; c < a.length; c++) b += a[c].join("\t") + "\n";
            return b.substring(0, b.length - 1)
        },
        _textToMatrix: function (c) {
            for (var a = c.split("\n"), b = 0; b < a.length; b++) a[b] = a[b].split("\t");
            return a
        },
        _tableToText: function (c) {
            var b = "";
            if (c.headers != a) b = OSF.DDA.DataCoercion._matrixToText([c.headers]) + "\n";
            var d = OSF.DDA.DataCoercion._matrixToText(c.rows);
            if (d == "") b = b.substring(0, b.length - 1);
            return b + d
        },
        _tableToMatrix: function (b) {
            var c = b.rows;
            b.headers != a && c.unshift(b.headers);
            return c
        },
        _coerceDataFromTable: function (c, b) {
            var a;
            switch (c) {
            case Microsoft.Office.WebExtension.CoercionType.Table:
                a = b;
                break;
            case Microsoft.Office.WebExtension.CoercionType.Matrix:
                a = OSF.DDA.DataCoercion._tableToMatrix(b);
                break;
            case Microsoft.Office.WebExtension.CoercionType.SlideRange:
                try {
                    var d = OSF.DDA.DataCoercion._tableToText(b);
                    a = new OSF.DDA.SlideRange(d)
                } catch (e) {
                    a = OSF.DDA.DataCoercion._tableToText(b)
                }
                break;
            case Microsoft.Office.WebExtension.CoercionType.Text:
            case Microsoft.Office.WebExtension.CoercionType.Html:
            case Microsoft.Office.WebExtension.CoercionType.Ooxml:
            default:
                a = OSF.DDA.DataCoercion._tableToText(b)
            }
            return a
        },
        _coerceDataToTable: function (b, c) {
            if (c == undefined) c = OSF.DDA.DataCoercion.determineCoercionType(b);
            var a;
            switch (c) {
            case Microsoft.Office.WebExtension.CoercionType.Table:
                a = b;
                break;
            case Microsoft.Office.WebExtension.CoercionType.Matrix:
                a = new Microsoft.Office.WebExtension.TableData(b);
                break;
            case Microsoft.Office.WebExtension.CoercionType.Text:
            case Microsoft.Office.WebExtension.CoercionType.Html:
            case Microsoft.Office.WebExtension.CoercionType.Ooxml:
            default:
                a = new Microsoft.Office.WebExtension.TableData(OSF.DDA.DataCoercion._textToMatrix(b))
            }
            return a
        }
    }
}();
OSF.DDA.issueAsyncResult = function (d, f, a) {
    var e = d[Microsoft.Office.WebExtension.Parameters.Callback];
    if (e) {
        var c = {};
        c[OSF.DDA.AsyncResultEnum.Properties.Context] = d[Microsoft.Office.WebExtension.Parameters.AsyncContext];
        var b;
        if (f == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) c[OSF.DDA.AsyncResultEnum.Properties.Value] = a;
        else {
            b = {};
            a = a || OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
            b[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = f || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
            b[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = a.name || a;
            b[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = a.message || a
        }
        e(new OSF.DDA.AsyncResult(c, b))
    }
};
OSF.DDA.generateBindingId = function () {
    return "UnnamedBinding_" + OSF.OUtil.getUniqueId() + "_" + (new Date).getTime()
};
OSF.DDA.SettingsManager = {
    SerializedSettings: "serializedSettings",
    DateJSONPrefix: "Date(",
    DataJSONSuffix: ")",
    serializeSettings: function (b) {
        var d = {};
        for (var c in b) {
            var a = b[c];
            try {
                if (JSON) a = JSON.stringify(a, function (a, b) {
                    return OSF.OUtil.isDate(this[a]) ? OSF.DDA.SettingsManager.DateJSONPrefix + this[a].getTime() + OSF.DDA.SettingsManager.DataJSONSuffix : b
                });
                else a = Sys.Serialization.JavaScriptSerializer.serialize(a);
                d[c] = a
            } catch (e) {}
        }
        return d
    },
    deserializeSettings: function (b) {
        var d = {};
        b = b || {};
        for (var c in b) {
            var a = b[c];
            try {
                if (JSON) a = JSON.parse(a, function (c, a) {
                    var b;
                    if (typeof a === "string" && a && a.length > 6 && a.slice(0, 5) === OSF.DDA.SettingsManager.DateJSONPrefix && a.slice(-1) === OSF.DDA.SettingsManager.DataJSONSuffix) {
                        b = new Date(parseInt(a.slice(5, -1)));
                        if (b) return b
                    }
                    return a
                });
                else a = Sys.Serialization.JavaScriptSerializer.deserialize(a, true);
                d[c] = a
            } catch (e) {}
        }
        return d
    }
};
OSF.DDA.OMFactory = {
    manufactureBinding: function (a, c) {
        var d = a[OSF.DDA.BindingProperties.Id],
            g = a[OSF.DDA.BindingProperties.RowCount],
            f = a[OSF.DDA.BindingProperties.ColumnCount],
            h = a[OSF.DDA.BindingProperties.HasHeaders],
            b;
        switch (a[OSF.DDA.BindingProperties.Type]) {
        case Microsoft.Office.WebExtension.BindingType.Text:
            b = new OSF.DDA.TextBinding(d, c);
            break;
        case Microsoft.Office.WebExtension.BindingType.Matrix:
            b = new OSF.DDA.MatrixBinding(d, c, g, f);
            break;
        case Microsoft.Office.WebExtension.BindingType.Table:
            var i = function () {
                    return OSF.DDA.ExcelDocument && Microsoft.Office.WebExtension.context.document && Microsoft.Office.WebExtension.context.document instanceof OSF.DDA.ExcelDocument
                },
                e;
            if (i() && OSF.DDA.ExcelTableBinding) e = OSF.DDA.ExcelTableBinding;
            else e = OSF.DDA.TableBinding;
            b = new e(d, c, g, f, h);
            break;
        default:
            b = new OSF.DDA.UnknownBinding(d, c)
        }
        return b
    },
    manufactureTableData: function (a) {
        return new Microsoft.Office.WebExtension.TableData(a[OSF.DDA.TableDataProperties.TableRows], a[OSF.DDA.TableDataProperties.TableHeaders])
    },
    manufactureDataNode: function (a) {
        if (a) return new OSF.DDA.CustomXmlNode(a[OSF.DDA.DataNodeProperties.Handle], a[OSF.DDA.DataNodeProperties.NodeType], a[OSF.DDA.DataNodeProperties.NamespaceUri], a[OSF.DDA.DataNodeProperties.BaseName])
    },
    manufactureDataPart: function (a, b) {
        return new OSF.DDA.CustomXmlPart(b, a[OSF.DDA.DataPartProperties.Id], a[OSF.DDA.DataPartProperties.BuiltIn])
    },
    manufactureEventArgs: function (e, c, a) {
        var d = this,
            b;
        switch (e) {
        case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:
            b = new OSF.DDA.DocumentSelectionChangedEventArgs(c);
            break;
        case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:
            b = new OSF.DDA.BindingSelectionChangedEventArgs(d.manufactureBinding(a, c.document), a[OSF.DDA.PropertyDescriptors.Subset]);
            break;
        case Microsoft.Office.WebExtension.EventType.BindingDataChanged:
            b = new OSF.DDA.BindingDataChangedEventArgs(d.manufactureBinding(a, c.document));
            break;
        case Microsoft.Office.WebExtension.EventType.SettingsChanged:
            b = new OSF.DDA.SettingsChangedEventArgs(c);
            break;
        case Microsoft.Office.Internal.EventType.OfficeThemeChanged:
            b = new OSF.DDA.OfficeThemeChangedEventArgs(a);
            break;
        case Microsoft.Office.Internal.EventType.DocumentThemeChanged:
            b = new OSF.DDA.DocumentThemeChangedEventArgs(a);
            break;
        case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:
            b = new OSF.DDA.ActiveViewChangedEventArgs(a);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeInserted:
            b = new OSF.DDA.NodeInsertedEventArgs(d.manufactureDataNode(a[OSF.DDA.DataNodeEventProperties.NewNode]), a[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:
            b = new OSF.DDA.NodeReplacedEventArgs(d.manufactureDataNode(a[OSF.DDA.DataNodeEventProperties.OldNode]), d.manufactureDataNode(a[OSF.DDA.DataNodeEventProperties.NewNode]), a[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:
            b = new OSF.DDA.NodeDeletedEventArgs(d.manufactureDataNode(a[OSF.DDA.DataNodeEventProperties.OldNode]), d.manufactureDataNode(a[OSF.DDA.DataNodeEventProperties.NextSiblingNode]), a[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:
            b = new OSF.DDA.TaskSelectionChangedEventArgs(c);
            break;
        case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:
            b = new OSF.DDA.ResourceSelectionChangedEventArgs(c);
            break;
        case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:
            b = new OSF.DDA.ViewSelectionChangedEventArgs(c);
            break;
        default:
            throw Error.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, e))
        }
        return b
    }
};
OSF.DDA.ListType = function () {
    var a = {};
    a[OSF.DDA.ListDescriptors.BindingList] = OSF.DDA.PropertyDescriptors.BindingProperties;
    a[OSF.DDA.ListDescriptors.DataPartList] = OSF.DDA.PropertyDescriptors.DataPartProperties;
    a[OSF.DDA.ListDescriptors.DataNodeList] = OSF.DDA.PropertyDescriptors.DataNodeProperties;
    return {
        isListType: function (b) {
            return OSF.OUtil.listContainsKey(a, b)
        },
        getDescriptor: function (b) {
            return a[b]
        }
    }
}();
OSF.DDA.AsyncMethodCall = function (d, e, g, i, j, h, m) {
    var b = "function",
        a = d.length,
        c = OSF.OUtil.delayExecutionAndCache(function () {
            return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters, m)
        });

    function f(d, f) {
        for (var e in d) {
            var a = d[e],
                b = f[e];
            if (a["enum"]) switch (typeof b) {
            case "string":
                if (OSF.OUtil.listContainsValue(a["enum"], b)) break;
            case "undefined":
                throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;
                break;
            default:
                throw c()
            }
            if (a["types"])
                if (!OSF.OUtil.listContainsValue(a["types"], typeof b)) throw c()
        }
    }

    function k(h, l, k) {
        if (h.length < a) throw Error.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);
        for (var e = [], b = 0; b < a; b++) e.push(h[b]);
        f(d, e);
        var j = {};
        for (b = 0; b < a; b++) {
            var g = d[b],
                i = e[b];
            if (g.verify) {
                var m = g.verify(i, l, k);
                if (!m) throw c()
            }
            j[g.name] = i
        }
        return j
    }

    function l(k, m, o, n) {
        if (k.length > a + 2) throw Error.parameterCount(Strings.OfficeOM.L_TooManyArguments);
        for (var c, d, l = k.length - 1; l >= a; l--) {
            var j = k[l];
            switch (typeof j) {
            case "object":
                if (c) throw Error.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
                else c = j;
                break;
            case b:
                if (d) throw Error.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);
                else d = j;
                break;
            default:
                throw Error.argument(Strings.OfficeOM.L_InValidOptionalArgument)
            }
        }
        c = c || {};
        for (var i in e)
            if (!OSF.OUtil.listContainsKey(c, i)) {
                var h = undefined,
                    g = e[i];
                if (g.calculate && m) h = g.calculate(m, o, n);
                if (!h && g.defaultValue != undefined) h = g.defaultValue;
                c[i] = h
            }
        if (d)
            if (c[Microsoft.Office.WebExtension.Parameters.Callback]) throw Strings.OfficeOM.L_RedundantCallbackSpecification;
            else c[Microsoft.Office.WebExtension.Parameters.Callback] = d;
        f(e, c);
        return c
    }
    this.verifyAndExtractCall = function (e, c, b) {
        var d = k(e, c, b),
            f = l(e, d, c, b),
            a = {};
        for (var j in d) a[j] = d[j];
        for (var i in f) a[i] = f[i];
        for (var m in g) a[m] = g[m](c, b);
        if (h) a = h(a, c, b);
        return a
    };
    this.processResponse = function (c, b, d, e) {
        var a;
        if (c == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
            if (i) a = i(b, d, e);
            else a = b;
        else if (j) a = j(c, b);
        else a = OSF.DDA.ErrorCodeManager.getErrorArgs(c);
        return a
    };
    this.getCallArgs = function (g) {
        for (var c, d, f = g.length - 1; f >= a; f--) {
            var e = g[f];
            switch (typeof e) {
            case "object":
                c = e;
                break;
            case b:
                d = e
            }
        }
        c = c || {};
        if (d) c[Microsoft.Office.WebExtension.Parameters.Callback] = d;
        return c
    }
};
OSF.DDA.ConvertToDocumentTheme = function (f) {
    var b = false,
        a = true;
    for (var d = [{
        name: "primaryFontColor",
        needToConvertToHex: a
    }, {
        name: "primaryBackgroundColor",
        needToConvertToHex: a
    }, {
        name: "secondaryFontColor",
        needToConvertToHex: a
    }, {
        name: "secondaryBackgroundColor",
        needToConvertToHex: a
    }, {
        name: "accent1",
        needToConvertToHex: a
    }, {
        name: "accent2",
        needToConvertToHex: a
    }, {
        name: "accent3",
        needToConvertToHex: a
    }, {
        name: "accent4",
        needToConvertToHex: a
    }, {
        name: "accent5",
        needToConvertToHex: a
    }, {
        name: "accent6",
        needToConvertToHex: a
    }, {
        name: "hyperlink",
        needToConvertToHex: a
    }, {
        name: "followedHyperlink",
        needToConvertToHex: a
    }, {
        name: "headerLatinFont",
        needToConvertToHex: b
    }, {
        name: "headerEastAsianFont",
        needToConvertToHex: b
    }, {
        name: "headerScriptFont",
        needToConvertToHex: b
    }, {
        name: "headerLocalizedFont",
        needToConvertToHex: b
    }, {
        name: "bodyLatinFont",
        needToConvertToHex: b
    }, {
        name: "bodyEastAsianFont",
        needToConvertToHex: b
    }, {
        name: "bodyScriptFont",
        needToConvertToHex: b
    }, {
        name: "bodyLocalizedFont",
        needToConvertToHex: b
    }], e = {}, c = 0; c < d.length; c++)
        if (d[c].needToConvertToHex) e[d[c].name] = OSF.OUtil.convertIntToHex(f[d[c].name]);
        else e[d[c].name] = f[d[c].name];
    return e
};
OSF.DDA.ConvertToOfficeTheme = function (a) {
    var b = {};
    for (var c in a) b[c] = OSF.OUtil.convertIntToHex(a[c]);
    return b
};
OSF.DDA.AsyncMethodNames = function (b) {
    var c = {};
    for (var a in b) {
        var d = {};
        OSF.OUtil.defineEnumerableProperties(d, {
            id: {
                value: a
            },
            displayName: {
                value: b[a]
            }
        });
        c[a] = d
    }
    return c
}({
    GoToByIdAsync: "goToByIdAsync",
    GetSelectedDataAsync: "getSelectedDataAsync",
    SetSelectedDataAsync: "setSelectedDataAsync",
    GetDocumentCopyAsync: "getFileAsync",
    GetDocumentCopyChunkAsync: "getSliceAsync",
    ReleaseDocumentCopyAsync: "closeAsync",
    GetFilePropertiesAsync: "getFilePropertiesAsync",
    AddFromSelectionAsync: "addFromSelectionAsync",
    AddFromPromptAsync: "addFromPromptAsync",
    AddFromNamedItemAsync: "addFromNamedItemAsync",
    GetAllAsync: "getAllAsync",
    GetByIdAsync: "getByIdAsync",
    ReleaseByIdAsync: "releaseByIdAsync",
    GetDataAsync: "getDataAsync",
    SetDataAsync: "setDataAsync",
    AddRowsAsync: "addRowsAsync",
    AddColumnsAsync: "addColumnsAsync",
    DeleteAllDataValuesAsync: "deleteAllDataValuesAsync",
    ClearFormatsAsync: "clearFormatsAsync",
    SetTableOptionsAsync: "setTableOptionsAsync",
    SetFormatsAsync: "setFormatsAsync",
    RefreshAsync: "refreshAsync",
    SaveAsync: "saveAsync",
    AddHandlerAsync: "addHandlerAsync",
    RemoveHandlerAsync: "removeHandlerAsync",
    GetActiveViewAsync: "getActiveViewAsync",
    AddDataPartAsync: "addAsync",
    GetDataPartByIdAsync: "getByIdAsync",
    GetDataPartsByNameSpaceAsync: "getByNamespaceAsync",
    DeleteDataPartAsync: "deleteAsync",
    GetPartNodesAsync: "getNodesAsync",
    GetPartXmlAsync: "getXmlAsync",
    AddDataPartNamespaceAsync: "addNamespaceAsync",
    GetDataPartNamespaceAsync: "getNamespaceAsync",
    GetDataPartPrefixAsync: "getPrefixAsync",
    GetRelativeNodesAsync: "getNodesAsync",
    GetNodeValueAsync: "getNodeValueAsync",
    GetNodeXmlAsync: "getXmlAsync",
    SetNodeValueAsync: "setNodeValueAsync",
    SetNodeXmlAsync: "setXmlAsync",
    GetOfficeThemeAsync: "getOfficeThemeAsync",
    GetDocumentThemeAsync: "getDocumentThemeAsync",
    GetSelectedTask: "getSelectedTaskAsync",
    GetTask: "getTaskAsync",
    GetWSSUrl: "getWSSUrlAsync",
    GetTaskField: "getTaskFieldAsync",
    GetSelectedResource: "getSelectedResourceAsync",
    GetResourceField: "getResourceFieldAsync",
    GetProjectField: "getProjectFieldAsync",
    GetSelectedView: "getSelectedViewAsync"
});
OSF.DDA.AsyncMethodCallFactory = function () {
    function a(a) {
        var c = null;
        if (a) {
            c = {};
            for (var d = a.length, b = 0; b < d; b++) c[a[b].name] = a[b].value
        }
        return c
    }
    return {
        manufacture: function (b) {
            var d = b.supportedOptions ? a(b.supportedOptions) : [],
                c = b.privateStateCallbacks ? a(b.privateStateCallbacks) : [];
            return new OSF.DDA.AsyncMethodCall(b.requiredArguments || [], d, c, b.onSucceeded, b.onFailed, b.checkCallArgs, b.method.displayName)
        }
    }
}();
OSF.DDA.AsyncMethodCalls = function () {
    var n = "function",
        g = "boolean",
        e = "object",
        c = "number",
        b = "string",
        l = {};

    function a(a) {
        l[a.method.id] = OSF.DDA.AsyncMethodCallFactory.manufacture(a)
    }

    function f(b, d, c) {
        var a = b[Microsoft.Office.WebExtension.Parameters.Data];
        if (a && (a[OSF.DDA.TableDataProperties.TableRows] != undefined || a[OSF.DDA.TableDataProperties.TableHeaders] != undefined)) a = OSF.DDA.OMFactory.manufactureTableData(a);
        a = OSF.DDA.DataCoercion.coerceData(a, c[Microsoft.Office.WebExtension.Parameters.CoercionType]);
        return a == undefined ? null : a
    }

    function h(a) {
        return OSF.DDA.OMFactory.manufactureBinding(a, Microsoft.Office.WebExtension.context.document)
    }

    function j(a) {
        return OSF.DDA.OMFactory.manufactureDataPart(a, Microsoft.Office.WebExtension.context.document.customXmlParts)
    }

    function m(a) {
        return OSF.DDA.OMFactory.manufactureDataNode(a)
    }

    function d(a) {
        return a.id
    }

    function k(b, a) {
        return a
    }

    function i(b, a) {
        return a
    }
    a({
        method: OSF.DDA.AsyncMethodNames.GoToByIdAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            types: [b, c]
        }, {
            name: Microsoft.Office.WebExtension.Parameters.GoToType,
            "enum": Microsoft.Office.WebExtension.GoToType
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.SelectionMode,
            value: {
                "enum": Microsoft.Office.WebExtension.SelectionMode,
                defaultValue: Microsoft.Office.WebExtension.SelectionMode.Default
            }
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetSelectedDataAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.CoercionType,
            "enum": Microsoft.Office.WebExtension.CoercionType
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.ValueFormat,
            value: {
                "enum": Microsoft.Office.WebExtension.ValueFormat,
                defaultValue: Microsoft.Office.WebExtension.ValueFormat.Unformatted
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.FilterType,
            value: {
                "enum": Microsoft.Office.WebExtension.FilterType,
                defaultValue: Microsoft.Office.WebExtension.FilterType.All
            }
        }],
        privateStateCallbacks: [],
        onSucceeded: f
    });
    a({
        method: OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Data,
            types: [b, e, c, g]
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.CoercionType,
            value: {
                "enum": Microsoft.Office.WebExtension.CoercionType,
                calculate: function (a) {
                    return OSF.DDA.DataCoercion.determineCoercionType(a[Microsoft.Office.WebExtension.Parameters.Data])
                }
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.CellFormat,
            value: {
                types: [e],
                defaultValue: []
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.TableOptions,
            value: {
                types: [e],
                defaultValue: []
            }
        }],
        privateStateCallbacks: []
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync,
        onSucceeded: function (a) {
            return new Microsoft.Office.WebExtension.FileProperties(a)
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.FileType,
            "enum": Microsoft.Office.WebExtension.FileType
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.SliceSize,
            value: {
                types: [c],
                defaultValue: 4 * 1024 * 1024
            }
        }],
        onSucceeded: function (a, c, b) {
            return new OSF.DDA.File(a[OSF.DDA.FileProperties.Handle], a[OSF.DDA.FileProperties.FileSize], b[Microsoft.Office.WebExtension.Parameters.SliceSize])
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetDocumentCopyChunkAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.SliceIndex,
            types: [c]
        }],
        privateStateCallbacks: [{
            name: OSF.DDA.FileProperties.Handle,
            value: function (b, a) {
                return a[OSF.DDA.FileProperties.Handle]
            }
        }, {
            name: OSF.DDA.FileProperties.SliceSize,
            value: function (b, a) {
                return a[OSF.DDA.FileProperties.SliceSize]
            }
        }],
        checkCallArgs: function (a, d, c) {
            var b = a[Microsoft.Office.WebExtension.Parameters.SliceIndex];
            if (b < 0 || b >= d.sliceCount) throw OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange;
            a[OSF.DDA.FileSliceOffset] = parseInt(b * c[OSF.DDA.FileProperties.SliceSize]);
            return a
        },
        onSucceeded: function (a, d, c) {
            var b = {};
            OSF.OUtil.defineEnumerableProperties(b, {
                data: {
                    value: a[Microsoft.Office.WebExtension.Parameters.Data]
                },
                index: {
                    value: c[Microsoft.Office.WebExtension.Parameters.SliceIndex]
                },
                size: {
                    value: a[OSF.DDA.FileProperties.SliceSize]
                }
            });
            return b
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.ReleaseDocumentCopyAsync,
        privateStateCallbacks: [{
            name: OSF.DDA.FileProperties.Handle,
            value: function (b, a) {
                return a[OSF.DDA.FileProperties.Handle]
            }
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.AddFromSelectionAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.BindingType,
            "enum": Microsoft.Office.WebExtension.BindingType
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            value: {
                types: [b],
                calculate: OSF.DDA.generateBindingId
            }
        }],
        privateStateCallbacks: [],
        onSucceeded: h
    });
    a({
        method: OSF.DDA.AsyncMethodNames.AddFromPromptAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.BindingType,
            "enum": Microsoft.Office.WebExtension.BindingType
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            value: {
                types: [b],
                calculate: OSF.DDA.generateBindingId
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.PromptText,
            value: {
                types: [b],
                calculate: function () {
                    return Strings.OfficeOM.L_AddBindingFromPromptDefaultText
                }
            }
        }],
        privateStateCallbacks: [],
        onSucceeded: h
    });
    a({
        method: OSF.DDA.AsyncMethodNames.AddFromNamedItemAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.ItemName,
            types: [b]
        }, {
            name: Microsoft.Office.WebExtension.Parameters.BindingType,
            "enum": Microsoft.Office.WebExtension.BindingType
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            value: {
                types: [b],
                calculate: OSF.DDA.generateBindingId
            }
        }],
        privateStateCallbacks: [{
            name: Microsoft.Office.WebExtension.Parameters.FailOnCollision,
            value: function () {
                return true
            }
        }],
        onSucceeded: h
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetAllAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: function (a) {
            return OSF.OUtil.mapList(a[OSF.DDA.ListDescriptors.BindingList], h)
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetByIdAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: h
    });
    a({
        method: OSF.DDA.AsyncMethodNames.ReleaseByIdAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: function (d, a, b) {
            var c = b[Microsoft.Office.WebExtension.Parameters.Id];
            delete a._eventDispatches[c]
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetDataAsync,
        requiredArguments: [],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.CoercionType,
            value: {
                "enum": Microsoft.Office.WebExtension.CoercionType,
                calculate: function (b, a) {
                    return OSF.DDA.DataCoercion.getCoercionDefaultForBinding(a.type)
                }
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.ValueFormat,
            value: {
                "enum": Microsoft.Office.WebExtension.ValueFormat,
                defaultValue: Microsoft.Office.WebExtension.ValueFormat.Unformatted
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.FilterType,
            value: {
                "enum": Microsoft.Office.WebExtension.FilterType,
                defaultValue: Microsoft.Office.WebExtension.FilterType.All
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.StartRow,
            value: {
                types: [c],
                defaultValue: 0
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.StartColumn,
            value: {
                types: [c],
                defaultValue: 0
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.RowCount,
            value: {
                types: [c],
                defaultValue: 0
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.ColumnCount,
            value: {
                types: [c],
                defaultValue: 0
            }
        }],
        checkCallArgs: function (a, b) {
            if (a[Microsoft.Office.WebExtension.Parameters.StartRow] == 0 && a[Microsoft.Office.WebExtension.Parameters.StartColumn] == 0 && a[Microsoft.Office.WebExtension.Parameters.RowCount] == 0 && a[Microsoft.Office.WebExtension.Parameters.ColumnCount] == 0) {
                delete a[Microsoft.Office.WebExtension.Parameters.StartRow];
                delete a[Microsoft.Office.WebExtension.Parameters.StartColumn];
                delete a[Microsoft.Office.WebExtension.Parameters.RowCount];
                delete a[Microsoft.Office.WebExtension.Parameters.ColumnCount]
            }
            if (a[Microsoft.Office.WebExtension.Parameters.CoercionType] != OSF.DDA.DataCoercion.getCoercionDefaultForBinding(b.type) && (a[Microsoft.Office.WebExtension.Parameters.StartRow] || a[Microsoft.Office.WebExtension.Parameters.StartColumn] || a[Microsoft.Office.WebExtension.Parameters.RowCount] || a[Microsoft.Office.WebExtension.Parameters.ColumnCount])) throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
            return a
        },
        privateStateCallbacks: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            value: d
        }],
        onSucceeded: f
    });
    a({
        method: OSF.DDA.AsyncMethodNames.SetDataAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Data,
            types: [b, e, c, g]
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.CoercionType,
            value: {
                "enum": Microsoft.Office.WebExtension.CoercionType,
                calculate: function (a) {
                    return OSF.DDA.DataCoercion.determineCoercionType(a[Microsoft.Office.WebExtension.Parameters.Data])
                }
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.StartRow,
            value: {
                types: [c],
                defaultValue: 0
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.StartColumn,
            value: {
                types: [c],
                defaultValue: 0
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.CellFormat,
            value: {
                types: [e],
                defaultValue: []
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.TableOptions,
            value: {
                types: [e],
                defaultValue: []
            }
        }],
        checkCallArgs: function (a, c) {
            var b = Microsoft.Office.WebExtension.Parameters;
            if (a[b.StartRow] == 0 && a[b.StartColumn] == 0 && OSF.OUtil.isArray(a[b.CellFormat]) && a[b.CellFormat].length === 0 && OSF.OUtil.isArray(a[b.TableOptions]) && a[b.TableOptions].length === 0) {
                delete a[b.StartRow];
                delete a[b.StartColumn];
                delete a[b.CellFormat];
                delete a[b.TableOptions]
            }
            if (a[b.CoercionType] != OSF.DDA.DataCoercion.getCoercionDefaultForBinding(c.type) && (a[b.StartRow] && a[b.StartRow] != 0 || a[b.StartColumn] && a[b.StartColumn] != 0 || a[b.CellFormat] || a[b.TableOptions])) throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
            return a
        },
        privateStateCallbacks: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            value: d
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.AddRowsAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Data,
            types: [e]
        }],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            value: d
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.AddColumnsAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Data,
            types: [e]
        }],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            value: d
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.DeleteAllDataValuesAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            value: d
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.ClearFormatsAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            value: d
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.SetTableOptionsAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.TableOptions,
            defaultValue: []
        }],
        privateStateCallbacks: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            value: d
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.SetFormatsAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.CellFormat,
            defaultValue: []
        }],
        privateStateCallbacks: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            value: d
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.RefreshAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: function (a) {
            var b = a[OSF.DDA.SettingsManager.SerializedSettings],
                c = OSF.DDA.SettingsManager.deserializeSettings(b);
            return c
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.SaveAsync,
        requiredArguments: [],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale,
            value: {
                types: [g],
                defaultValue: true
            }
        }],
        privateStateCallbacks: [{
            name: OSF.DDA.SettingsManager.SerializedSettings,
            value: function (b, a) {
                return OSF.DDA.SettingsManager.serializeSettings(a)
            }
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.AddHandlerAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            verify: function (b, c, a) {
                return a.supportsEvent(b)
            }
        }, {
            name: Microsoft.Office.WebExtension.Parameters.Handler,
            types: [n]
        }],
        supportedOptions: [],
        privateStateCallbacks: []
    });
    a({
        method: OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            verify: function (b, c, a) {
                return a.supportsEvent(b)
            }
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.Handler,
            value: {
                types: [n],
                defaultValue: null
            }
        }],
        privateStateCallbacks: []
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetDocumentThemeAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: OSF.DDA.ConvertToDocumentTheme
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetOfficeThemeAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: OSF.DDA.ConvertToOfficeTheme
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetActiveViewAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: function (b) {
            var a = b[Microsoft.Office.WebExtension.Parameters.ActiveView];
            return a == undefined ? null : a
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.AddDataPartAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Xml,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: j
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetDataPartByIdAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Id,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: j
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetDataPartsByNameSpaceAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Namespace,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: function (a) {
            return OSF.OUtil.mapList(a[OSF.DDA.ListDescriptors.DataPartList], j)
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.DeleteDataPartAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: OSF.DDA.DataPartProperties.Id,
            value: d
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetPartNodesAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.XPath,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: OSF.DDA.DataPartProperties.Id,
            value: d
        }],
        onSucceeded: function (a) {
            return OSF.OUtil.mapList(a[OSF.DDA.ListDescriptors.DataNodeList], m)
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetPartXmlAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: OSF.DDA.DataPartProperties.Id,
            value: d
        }],
        onSucceeded: f
    });
    a({
        method: OSF.DDA.AsyncMethodNames.AddDataPartNamespaceAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Prefix,
            types: [b]
        }, {
            name: Microsoft.Office.WebExtension.Parameters.Namespace,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: OSF.DDA.DataPartProperties.Id,
            value: k
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetDataPartNamespaceAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Prefix,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: OSF.DDA.DataPartProperties.Id,
            value: k
        }],
        onSucceeded: f
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetDataPartPrefixAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Namespace,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: OSF.DDA.DataPartProperties.Id,
            value: k
        }],
        onSucceeded: f
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetRelativeNodesAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.XPath,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: OSF.DDA.DataNodeProperties.Handle,
            value: i
        }],
        onSucceeded: function (a) {
            return OSF.OUtil.mapList(a[OSF.DDA.ListDescriptors.DataNodeList], m)
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetNodeValueAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: OSF.DDA.DataNodeProperties.Handle,
            value: i
        }],
        onSucceeded: f
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetNodeXmlAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: OSF.DDA.DataNodeProperties.Handle,
            value: i
        }],
        onSucceeded: f
    });
    a({
        method: OSF.DDA.AsyncMethodNames.SetNodeValueAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Data,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: OSF.DDA.DataNodeProperties.Handle,
            value: i
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.SetNodeXmlAsync,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Xml,
            types: [b]
        }],
        supportedOptions: [],
        privateStateCallbacks: [{
            name: OSF.DDA.DataNodeProperties.Handle,
            value: i
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetSelectedTask,
        onSucceeded: function (a) {
            return a[Microsoft.Office.WebExtension.Parameters.TaskId]
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetTask,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.TaskId,
            types: [b]
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetTaskField,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.TaskId,
            types: [b]
        }, {
            name: Microsoft.Office.WebExtension.Parameters.FieldId,
            types: [c]
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.GetRawValue,
            value: {
                types: [g],
                defaultValue: false
            }
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetResourceField,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.ResourceId,
            types: [b]
        }, {
            name: Microsoft.Office.WebExtension.Parameters.FieldId,
            types: [c]
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.GetRawValue,
            value: {
                types: [g],
                defaultValue: false
            }
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetProjectField,
        requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.FieldId,
            types: [c]
        }],
        supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.GetRawValue,
            value: {
                types: [g],
                defaultValue: false
            }
        }]
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetSelectedResource,
        onSucceeded: function (a) {
            return a[Microsoft.Office.WebExtension.Parameters.ResourceId]
        }
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetWSSUrl
    });
    a({
        method: OSF.DDA.AsyncMethodNames.GetSelectedView
    });
    return l
}();
OSF.DDA.HostParameterMap = function (a, b) {
    var i = "fromHost",
        c = this,
        j = "toHost",
        h = i,
        e = "self",
        g = {};
    g[Microsoft.Office.WebExtension.Parameters.Data] = {
        toHost: function (a) {
            if (a.rows !== undefined) {
                var b = {};
                b[OSF.DDA.TableDataProperties.TableRows] = a.rows;
                b[OSF.DDA.TableDataProperties.TableHeaders] = a.headers;
                a = b
            }
            return a
        },
        fromHost: function (a) {
            return a
        }
    };

    function d(j, i) {
        var m = j ? {} : undefined;
        for (var f in j) {
            var e = j[f],
                c;
            if (OSF.DDA.ListType.isListType(f)) {
                c = [];
                for (var n in e) c.push(d(e[n], i))
            } else if (OSF.OUtil.listContainsKey(g, f)) c = g[f][i](e);
            else if (i == h && a.preserveNesting(f)) c = d(e, i);
            else {
                var k = b[f];
                if (k) {
                    var l = k[i];
                    if (l) {
                        c = l[e];
                        if (c === undefined) c = e
                    }
                } else c = e
            }
            m[f] = c
        }
        return m
    }

    function k(i, h) {
        var f;
        for (var c in h) {
            var d;
            if (a.isComplexType(c)) d = k(i, b[c][j]);
            else d = i[c]; if (d != undefined) {
                if (!f) f = {};
                var g = h[c];
                if (g == e) g = c;
                f[g] = a.pack(c, d)
            }
        }
        return f
    }

    function f(k, j, g) {
        if (!g) g = {};
        for (var d in j) {
            var l = j[d],
                c;
            if (l == e) c = k;
            else c = k[l]; if (c === null || c === undefined) g[d] = undefined;
            else {
                c = a.unpack(d, c);
                var i;
                if (a.isComplexType(d)) {
                    i = b[d][h];
                    if (a.preserveNesting(d)) g[d] = f(c, i);
                    else f(c, i, g)
                } else {
                    if (OSF.DDA.ListType.isListType(d)) {
                        i = {};
                        var n = OSF.DDA.ListType.getDescriptor(d);
                        i[n] = e;
                        for (var m in c) c[m] = f(c[m], i)
                    }
                    g[d] = c
                }
            }
        }
        return g
    }

    function l(l, g, a) {
        var e = b[l][a],
            c;
        if (a == "toHost") {
            var j = d(g, a);
            c = k(j, e)
        } else if (a == i) {
            var h = f(g, e);
            c = d(h, a)
        }
        return c
    }
    if (!b) b = {};
    c.setMapping = function (k, c) {
        var a, d;
        if (c.map) {
            a = c.map;
            d = {};
            for (var f in a) {
                var g = a[f];
                if (g == e) g = f;
                d[g] = f
            }
        } else {
            a = c.toHost;
            d = c.fromHost
        }
        var i = b[k] = {};
        i[j] = a;
        i[h] = d
    };
    c.toHost = function (b, a) {
        return l(b, a, j)
    };
    c.fromHost = function (a, b) {
        return l(a, b, h)
    };
    c.self = e;
    c.dynamicTypes = g;
    c.mapValues = d;
    c.specialProcessorDynamicTypes = a.dynamicTypes
};
OSF.DDA.SpecialProcessor = function (c, b) {
    var a = this;
    a.isComplexType = function (a) {
        return OSF.OUtil.listContainsValue(c, a)
    };
    a.isDynamicType = function (a) {
        return OSF.OUtil.listContainsKey(b, a)
    };
    a.preserveNesting = function (b) {
        var a = [OSF.DDA.PropertyDescriptors.Subset, OSF.DDA.DataNodeEventProperties.OldNode, OSF.DDA.DataNodeEventProperties.NewNode, OSF.DDA.DataNodeEventProperties.NextSiblingNode];
        return OSF.OUtil.listContainsValue(a, b)
    };
    a.pack = function (c, d) {
        var a;
        if (this.isDynamicType(c)) a = b[c].toHost(d);
        else a = d;
        return a
    };
    a.unpack = function (c, d) {
        var a;
        if (this.isDynamicType(c)) a = b[c].fromHost(d);
        else a = d;
        return a
    }
};
OSF.DDA.DispIdHost.Facade = function (e, d) {
    var a = {},
        b = OSF.DDA.AsyncMethodNames,
        c = OSF.DDA.MethodDispId;
    a[b.GoToByIdAsync.id] = c.dispidNavigateToMethod;
    a[b.GetSelectedDataAsync.id] = c.dispidGetSelectedDataMethod;
    a[b.SetSelectedDataAsync.id] = c.dispidSetSelectedDataMethod;
    a[b.GetDocumentCopyChunkAsync.id] = c.dispidGetDocumentCopyChunkMethod;
    a[b.ReleaseDocumentCopyAsync.id] = c.dispidReleaseDocumentCopyMethod;
    a[b.GetDocumentCopyAsync.id] = c.dispidGetDocumentCopyMethod;
    a[b.AddFromSelectionAsync.id] = c.dispidAddBindingFromSelectionMethod;
    a[b.AddFromPromptAsync.id] = c.dispidAddBindingFromPromptMethod;
    a[b.AddFromNamedItemAsync.id] = c.dispidAddBindingFromNamedItemMethod;
    a[b.GetAllAsync.id] = c.dispidGetAllBindingsMethod;
    a[b.GetByIdAsync.id] = c.dispidGetBindingMethod;
    a[b.ReleaseByIdAsync.id] = c.dispidReleaseBindingMethod;
    a[b.GetDataAsync.id] = c.dispidGetBindingDataMethod;
    a[b.SetDataAsync.id] = c.dispidSetBindingDataMethod;
    a[b.GetFilePropertiesAsync.id] = c.dispidGetFilePropertiesMethod;
    a[b.AddRowsAsync.id] = c.dispidAddRowsMethod;
    a[b.AddColumnsAsync.id] = c.dispidAddColumnsMethod;
    a[b.DeleteAllDataValuesAsync.id] = c.dispidClearAllRowsMethod;
    a[b.ClearFormatsAsync.id] = c.dispidClearFormatsMethod;
    a[b.RefreshAsync.id] = c.dispidLoadSettingsMethod;
    a[b.SaveAsync.id] = c.dispidSaveSettingsMethod;
    a[b.SetTableOptionsAsync.id] = c.dispidSetTableOptionsMethod;
    a[b.SetFormatsAsync.id] = c.dispidSetFormatsMethod;
    a[b.GetActiveViewAsync.id] = c.dispidGetActiveViewMethod;
    a[b.AddDataPartAsync.id] = c.dispidAddDataPartMethod;
    a[b.GetDataPartByIdAsync.id] = c.dispidGetDataPartByIdMethod;
    a[b.GetDataPartsByNameSpaceAsync.id] = c.dispidGetDataPartsByNamespaceMethod;
    a[b.GetPartXmlAsync.id] = c.dispidGetDataPartXmlMethod;
    a[b.GetPartNodesAsync.id] = c.dispidGetDataPartNodesMethod;
    a[b.DeleteDataPartAsync.id] = c.dispidDeleteDataPartMethod;
    a[b.GetNodeValueAsync.id] = c.dispidGetDataNodeValueMethod;
    a[b.GetNodeXmlAsync.id] = c.dispidGetDataNodeXmlMethod;
    a[b.GetRelativeNodesAsync.id] = c.dispidGetDataNodesMethod;
    a[b.SetNodeValueAsync.id] = c.dispidSetDataNodeValueMethod;
    a[b.SetNodeXmlAsync.id] = c.dispidSetDataNodeXmlMethod;
    a[b.AddDataPartNamespaceAsync.id] = c.dispidAddDataNamespaceMethod;
    a[b.GetDataPartNamespaceAsync.id] = c.dispidGetDataUriByPrefixMethod;
    a[b.GetDataPartPrefixAsync.id] = c.dispidGetDataPrefixByUriMethod;
    a[b.GetDocumentThemeAsync.id] = c.dispidGetDocumentThemeMethod;
    a[b.GetOfficeThemeAsync.id] = c.dispidGetOfficeThemeMethod;
    a[b.GetSelectedTask.id] = c.dispidGetSelectedTaskMethod;
    a[b.GetTask.id] = c.dispidGetTaskMethod;
    a[b.GetWSSUrl.id] = c.dispidGetWSSUrlMethod;
    a[b.GetTaskField.id] = c.dispidGetTaskFieldMethod;
    a[b.GetSelectedResource.id] = c.dispidGetSelectedResourceMethod;
    a[b.GetResourceField.id] = c.dispidGetResourceFieldMethod;
    a[b.GetProjectField.id] = c.dispidGetProjectFieldMethod;
    a[b.GetSelectedView.id] = c.dispidGetSelectedViewMethod;
    b = Microsoft.Office.WebExtension.EventType;
    c = OSF.DDA.EventDispId;
    a[b.SettingsChanged] = c.dispidSettingsChangedEvent;
    a[b.DocumentSelectionChanged] = c.dispidDocumentSelectionChangedEvent;
    a[b.BindingSelectionChanged] = c.dispidBindingSelectionChangedEvent;
    a[b.BindingDataChanged] = c.dispidBindingDataChangedEvent;
    a[b.ActiveViewChanged] = c.dispidActiveViewChangedEvent;
    a[b.DocumentThemeChanged] = c.dispidDocumentThemeChangedEvent;
    a[b.OfficeThemeChanged] = c.dispidOfficeThemeChangedEvent;
    a[b.TaskSelectionChanged] = c.dispidTaskSelectionChangedEvent;
    a[b.ResourceSelectionChanged] = c.dispidResourceSelectionChangedEvent;
    a[b.ViewSelectionChanged] = c.dispidViewSelectionChangedEvent;
    a[b.DataNodeInserted] = c.dispidDataNodeAddedEvent;
    a[b.DataNodeReplaced] = c.dispidDataNodeReplacedEvent;
    a[b.DataNodeDeleted] = c.dispidDataNodeDeletedEvent;

    function f(a, c, d, b) {
        if (typeof a == "number") {
            if (!b) b = c.getCallArgs(d);
            OSF.DDA.issueAsyncResult(b, a, OSF.DDA.ErrorCodeManager.getErrorArgs(a))
        } else throw a
    }
    this[OSF.DDA.DispIdHost.Methods.InvokeMethod] = function (o, j, k, m) {
        var b;
        try {
            var h = o.id,
                c = OSF.DDA.AsyncMethodCalls[h];
            b = c.verifyAndExtractCall(j, k, m);
            var i = a[h],
                n = e(h),
                g;
            if (d.toHost) g = d.toHost(i, b);
            else g = b;
            n[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({
                dispId: i,
                hostCallArgs: g,
                onCalling: function () {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                },
                onReceiving: function () {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                },
                onComplete: function (f, e) {
                    var a;
                    if (f == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                        if (d.fromHost) a = d.fromHost(i, e);
                        else a = e;
                    else a = e;
                    var g = c.processResponse(f, a, k, b);
                    OSF.DDA.issueAsyncResult(b, f, g)
                }
            })
        } catch (l) {
            f(l, c, j, b)
        }
    };
    this[OSF.DDA.DispIdHost.Methods.AddEventHandler] = function (j, g, h) {
        var c, b, l;

        function i(a) {
            if (a == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                var e = g.addEventHandler(b, l);
                if (!e) a = OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed
            }
            var d;
            if (a != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) d = OSF.DDA.ErrorCodeManager.getErrorArgs(a);
            OSF.DDA.issueAsyncResult(c, a, d)
        }
        try {
            var k = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];
            c = k.verifyAndExtractCall(j, h, g);
            b = c[Microsoft.Office.WebExtension.Parameters.EventType];
            l = c[Microsoft.Office.WebExtension.Parameters.Handler];
            if (g.getEventHandlerCount(b) == 0) {
                var m = a[b],
                    o = e(b)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
                o({
                    eventType: b,
                    dispId: m,
                    targetId: h.id || "",
                    onCalling: function () {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                    },
                    onReceiving: function () {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                    },
                    onComplete: i,
                    onEvent: function (a) {
                        var c = d.fromHost(m, a);
                        g.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(b, h, c))
                    }
                })
            } else i(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
        } catch (n) {
            f(n, k, j, c)
        }
    };
    this[OSF.DDA.DispIdHost.Methods.RemoveEventHandler] = function (j, c, l) {
        var d, b, g;

        function i(a) {
            var b;
            if (a != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) b = OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist);
            OSF.DDA.issueAsyncResult(d, a, b)
        }
        try {
            var k = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];
            d = k.verifyAndExtractCall(j, l, c);
            b = d[Microsoft.Office.WebExtension.Parameters.EventType];
            g = d[Microsoft.Office.WebExtension.Parameters.Handler];
            var h;
            if (g == null) {
                c.clearEventHandlers(b);
                h = true
            } else if (!c.hasEventHandler(b, g)) h = false;
            else h = c.removeEventHandler(b, g); if (c.getEventHandlerCount(b) == 0) {
                var o = a[b],
                    n = e(b)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
                n({
                    eventType: b,
                    dispId: o,
                    targetId: l.id || "",
                    onCalling: function () {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                    },
                    onReceiving: function () {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                    },
                    onComplete: i
                })
            } else i(h ? OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess : Strings.OfficeOM.L_EventRegistrationError)
        } catch (m) {
            f(m, k, j, d)
        }
    }
};
OSF.DDA.DispIdHost.addAsyncMethods = function (a, b, e) {
    for (var f in b) {
        var c = b[f],
            d = c.displayName;
        !a[d] && OSF.OUtil.defineEnumerableProperty(a, d, {
            value: function (b) {
                return function () {
                    var c = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];
                    c(b, arguments, a, e)
                }
            }(c)
        })
    }
};
OSF.DDA.DispIdHost.addEventSupport = function (a, b) {
    var d = OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName,
        c = OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;
    !a[d] && OSF.OUtil.defineEnumerableProperty(a, d, {
        value: function () {
            var c = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];
            c(arguments, b, a)
        }
    });
    !a[c] && OSF.OUtil.defineEnumerableProperty(a, c, {
        value: function () {
            var c = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];
            c(arguments, b, a)
        }
    })
};
OSF.DDA.Context = function (c, d, e, b) {
    var a = this;
    OSF.OUtil.defineEnumerableProperties(a, {
        contentLanguage: {
            value: c.get_dataLocale()
        },
        displayLanguage: {
            value: c.get_appUILocale()
        }
    });
    d && OSF.OUtil.defineEnumerableProperty(a, "document", {
        value: d
    });
    e && OSF.OUtil.defineEnumerableProperty(a, "license", {
        value: e
    });
    if (b) {
        var f = b.displayName || "appOM";
        delete b.displayName;
        OSF.OUtil.defineEnumerableProperty(a, f, {
            value: b
        })
    }
};
OSF.DDA.OutlookContext = function (b, a, c, d) {
    OSF.DDA.OutlookContext.uber.constructor.call(this, b, null, c, d);
    a && OSF.OUtil.defineEnumerableProperty(this, "roamingSettings", {
        value: a
    })
};
OSF.OUtil.extend(OSF.DDA.OutlookContext, OSF.DDA.Context);
OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension, "context", {
    "get": function () {
        var a;
        if (OSF && OSF._OfficeAppFactory) a = OSF._OfficeAppFactory.getContext();
        return a
    }
});
Microsoft.Office.WebExtension.useShortNamespace = function (a) {
    if (a) OSF.NamespaceManager.enableShortcut();
    else OSF.NamespaceManager.disableShortcut()
};
Microsoft.Office.WebExtension.select = function (a, b) {
    var c;
    if (a && typeof a == "string") {
        var d = a.indexOf("#");
        if (d != -1) {
            var h = a.substring(0, d),
                g = a.substring(d + 1);
            switch (h) {
            case "binding":
            case "bindings":
                if (g) c = new OSF.DDA.BindingPromise(g)
            }
        }
    }
    if (!c) {
        if (b) {
            var e = typeof b;
            if (e == "function") {
                var f = {};
                f[Microsoft.Office.WebExtension.Parameters.Callback] = b;
                OSF.DDA.issueAsyncResult(f, OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext, OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext))
            } else throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, e)
        }
    } else {
        c.onFail = b;
        return c
    }
};
OSF.DDA.BindingPromise = function (b, a) {
    this._id = b;
    OSF.OUtil.defineEnumerableProperty(this, "onFail", {
        "get": function () {
            return a
        },
        "set": function (c) {
            var b = typeof c;
            if (b != "undefined" && b != "function") throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, b);
            a = c
        }
    })
};
OSF.DDA.BindingPromise.prototype = {
    _fetch: function (b) {
        var a = this;
        if (a.binding) b && b(a.binding);
        else if (!a._binding) {
            var c = a;
            Microsoft.Office.WebExtension.context.document.bindings.getByIdAsync(a._id, function (a) {
                if (a.status == Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded) {
                    OSF.OUtil.defineEnumerableProperty(c, "binding", {
                        value: a.value
                    });
                    b && b(c.binding)
                } else c.onFail && c.onFail(a)
            })
        }
        return a
    },
    getDataAsync: function () {
        var a = arguments;
        this._fetch(function (b) {
            b.getDataAsync.apply(b, a)
        });
        return this
    },
    setDataAsync: function () {
        var a = arguments;
        this._fetch(function (b) {
            b.setDataAsync.apply(b, a)
        });
        return this
    },
    addHandlerAsync: function () {
        var a = arguments;
        this._fetch(function (b) {
            b.addHandlerAsync.apply(b, a)
        });
        return this
    },
    removeHandlerAsync: function () {
        var a = arguments;
        this._fetch(function (b) {
            b.removeHandlerAsync.apply(b, a)
        });
        return this
    },
    setTableOptionsAsync: function () {
        var a = arguments;
        this._fetch(function (b) {
            b.setTableOptionsAsync.apply(b, a)
        });
        return this
    },
    setFormatsAsync: function () {
        var a = arguments;
        this._fetch(function (b) {
            b.setFormatsAsync.apply(b, a)
        });
        return this
    },
    clearFormatsAsync: function () {
        var a = arguments;
        this._fetch(function (b) {
            b.clearFormatsAsync.apply(b, a)
        });
        return this
    }
};
OSF.DDA.License = function (a) {
    OSF.OUtil.defineEnumerableProperty(this, "value", {
        value: a
    })
};
OSF.DDA.Settings = function (b) {
    var a = "name";
    b = b || {};
    OSF.OUtil.defineEnumerableProperties(this, {
        "get": {
            value: function (e) {
                var d = Function._validateParams(arguments, [{
                    name: a,
                    type: String,
                    mayBeNull: false
                }]);
                if (d) throw d;
                var c = b[e];
                return typeof c === "undefined" ? null : c
            }
        },
        "set": {
            value: function (e, d) {
                var c = Function._validateParams(arguments, [{
                    name: a,
                    type: String,
                    mayBeNull: false
                }, {
                    name: "value",
                    mayBeNull: true
                }]);
                if (c) throw c;
                b[e] = d
            }
        },
        remove: {
            value: function (d) {
                var c = Function._validateParams(arguments, [{
                    name: a,
                    type: String,
                    mayBeNull: false
                }]);
                if (c) throw c;
                delete b[d]
            }
        }
    });
    OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.SaveAsync], b)
};
OSF.DDA.RefreshableSettings = function (a) {
    OSF.DDA.RefreshableSettings.uber.constructor.call(this, a);
    OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.RefreshAsync], a);
    OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.SettingsChanged]))
};
OSF.OUtil.extend(OSF.DDA.RefreshableSettings, OSF.DDA.Settings);
OSF.DDA.OutlookAppOm = function () {};
OSF.DDA.Document = function (b, c) {
    var a;
    switch (b.get_clientMode()) {
    case OSF.ClientMode.ReadOnly:
        a = Microsoft.Office.WebExtension.DocumentMode.ReadOnly;
        break;
    case OSF.ClientMode.ReadWrite:
        a = Microsoft.Office.WebExtension.DocumentMode.ReadWrite
    }
    c && OSF.OUtil.defineEnumerableProperty(this, "settings", {
        value: c
    });
    OSF.OUtil.defineMutableProperties(this, {
        mode: {
            value: a
        },
        url: {
            value: b.get_docUrl()
        }
    })
};
OSF.DDA.JsomDocument = function (c, d, e) {
    var a = this;
    OSF.DDA.JsomDocument.uber.constructor.call(a, c, e);
    OSF.OUtil.defineEnumerableProperty(a, "bindings", {
        "get": function () {
            return d
        }
    });
    var b = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(a, [b.GetSelectedDataAsync, b.SetSelectedDataAsync]);
    OSF.DDA.DispIdHost.addEventSupport(a, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]))
};
OSF.OUtil.extend(OSF.DDA.JsomDocument, OSF.DDA.Document);
OSF.DDA.BindingFacade = function (b) {
    this._eventDispatches = [];
    OSF.OUtil.defineEnumerableProperty(this, "document", {
        value: b
    });
    var a = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(this, [a.AddFromSelectionAsync, a.AddFromNamedItemAsync, a.GetAllAsync, a.GetByIdAsync, a.ReleaseByIdAsync])
};
OSF.DDA.UnknownBinding = function (b, a) {
    OSF.OUtil.defineEnumerableProperties(this, {
        document: {
            value: a
        },
        id: {
            value: b
        }
    })
};
OSF.DDA.Binding = function (a, c) {
    OSF.OUtil.defineEnumerableProperties(this, {
        document: {
            value: c
        },
        id: {
            value: a
        }
    });
    var d = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(this, [d.GetDataAsync, d.SetDataAsync]);
    var e = Microsoft.Office.WebExtension.EventType,
        b = c.bindings._eventDispatches;
    if (!b[a]) b[a] = new OSF.EventDispatch([e.BindingSelectionChanged, e.BindingDataChanged]);
    var f = b[a];
    OSF.DDA.DispIdHost.addEventSupport(this, f)
};
OSF.DDA.TextBinding = function (b, a) {
    OSF.DDA.TextBinding.uber.constructor.call(this, b, a);
    OSF.OUtil.defineEnumerableProperty(this, "type", {
        value: Microsoft.Office.WebExtension.BindingType.Text
    })
};
OSF.OUtil.extend(OSF.DDA.TextBinding, OSF.DDA.Binding);
OSF.DDA.MatrixBinding = function (d, c, b, a) {
    OSF.DDA.MatrixBinding.uber.constructor.call(this, d, c);
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.BindingType.Matrix
        },
        rowCount: {
            value: b ? b : 0
        },
        columnCount: {
            value: a ? a : 0
        }
    })
};
OSF.OUtil.extend(OSF.DDA.MatrixBinding, OSF.DDA.Binding);
OSF.DDA.TableBinding = function (f, e, d, c, b) {
    OSF.DDA.TableBinding.uber.constructor.call(this, f, e);
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.BindingType.Table
        },
        rowCount: {
            value: d ? d : 0
        },
        columnCount: {
            value: c ? c : 0
        },
        hasHeaders: {
            value: b ? b : false
        }
    });
    var a = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(this, [a.AddRowsAsync, a.AddColumnsAsync, a.DeleteAllDataValuesAsync])
};
OSF.OUtil.extend(OSF.DDA.TableBinding, OSF.DDA.Binding);
Microsoft.Office.WebExtension.TableData = function (b, a) {
    function c(a) {
        if (a == null || a == undefined) return null;
        try {
            for (var b = OSF.DDA.DataCoercion.findArrayDimensionality(a, 2); b < 2; b++) a = [a];
            return a
        } catch (c) {}
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        headers: {
            "get": function () {
                return a
            },
            "set": function (b) {
                a = c(b)
            }
        },
        rows: {
            "get": function () {
                return b
            },
            "set": function (a) {
                b = a == null || OSF.OUtil.isArray(a) && a.length == 0 ? [] : c(a)
            }
        }
    });
    this.headers = a;
    this.rows = b
};
Microsoft.Office.WebExtension.FileProperties = function (a) {
    OSF.OUtil.defineEnumerableProperties(this, {
        url: {
            value: a[OSF.DDA.FilePropertiesDescriptor.Url]
        }
    })
};
OSF.DDA.Error = function (c, a, b) {
    OSF.OUtil.defineEnumerableProperties(this, {
        name: {
            value: c
        },
        message: {
            value: a
        },
        code: {
            value: b
        }
    })
};
OSF.DDA.AsyncResult = function (b, a) {
    OSF.OUtil.defineEnumerableProperties(this, {
        value: {
            value: b[OSF.DDA.AsyncResultEnum.Properties.Value]
        },
        status: {
            value: a ? Microsoft.Office.WebExtension.AsyncResultStatus.Failed : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded
        }
    });
    b[OSF.DDA.AsyncResultEnum.Properties.Context] && OSF.OUtil.defineEnumerableProperty(this, "asyncContext", {
        value: b[OSF.DDA.AsyncResultEnum.Properties.Context]
    });
    a && OSF.OUtil.defineEnumerableProperty(this, "error", {
        value: new OSF.DDA.Error(a[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], a[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], a[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])
    })
};
OSF.DDA.DocumentSelectionChangedEventArgs = function (a) {
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged
        },
        document: {
            value: a
        }
    })
};
OSF.DDA.BindingSelectionChangedEventArgs = function (c, a) {
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.BindingSelectionChanged
        },
        binding: {
            value: c
        }
    });
    for (var b in a) OSF.OUtil.defineEnumerableProperty(this, b, {
        value: a[b]
    })
};
OSF.DDA.BindingDataChangedEventArgs = function (a) {
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.BindingDataChanged
        },
        binding: {
            value: a
        }
    })
};
OSF.DDA.SettingsChangedEventArgs = function (a) {
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.SettingsChanged
        },
        settings: {
            value: a
        }
    })
};
OSF.DDA.OfficeThemeChangedEventArgs = function (a) {
    var b = OSF.DDA.ConvertToOfficeTheme(a);
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.Internal.EventType.OfficeThemeChanged
        },
        officeTheme: {
            value: b
        }
    })
};
OSF.DDA.DocumentThemeChangedEventArgs = function (a) {
    var b = OSF.DDA.ConvertToDocumentTheme(a);
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.Internal.EventType.DocumentThemeChanged
        },
        documentTheme: {
            value: b
        }
    })
};
OSF.DDA.ActiveViewChangedEventArgs = function (a) {
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.ActiveViewChanged
        },
        activeView: {
            value: a.activeView
        }
    })
};
OSF.O15HostSpecificFileVersion = {
    Fallback: "15.01",
    GenerateVersion: function (d, c) {
        var a = 2;
        return b(d, a) + "." + b(c, a);

        function b(b, c) {
            b = b || 0;
            c = c || 0;
            for (var a = b.toString(), e = c - a.length, d = 0; d < e; d++) a = "0" + a;
            return a
        }
    }
};
var __extends = this.__extends || function (b, c) {
        function a() {
            this.constructor = b
        }
        a.prototype = c.prototype;
        b.prototype = new a
    },
    OSFLog;
(function (e) {
    var b = "SessionId",
        c = "AssetId",
        a = true,
        d = function () {
            function b(a) {
                this._table = a;
                this._fields = {}
            }
            Object.defineProperty(b.prototype, "Fields", {
                "get": function () {
                    return this._fields
                },
                enumerable: a,
                configurable: a
            });
            Object.defineProperty(b.prototype, "Table", {
                "get": function () {
                    return this._table
                },
                enumerable: a,
                configurable: a
            });
            b.prototype.SerializeFields = function () {};
            b.prototype.SetSerializedField = function (b, a) {
                if (typeof a !== "undefined" && a !== null) this._serializedFields[b] = a.toString()
            };
            b.prototype.SerializeRow = function () {
                var a = this;
                a._serializedFields = {};
                a.SetSerializedField("Table", a._table);
                a.SerializeFields();
                return JSON.stringify(a._serializedFields)
            };
            return b
        }();
    e.BaseUsageData = d;
    var g = function (m) {
        var l = "ErrorResult",
            k = "Stage7Time",
            j = "Stage6Time",
            i = "Stage5Time",
            h = "Stage4Time",
            g = "Stage3Time",
            f = "Stage2Time",
            e = "Stage1Time",
            d = "AppInfo";
        __extends(b, m);

        function b() {
            m.call(this, "AppLoadTime")
        }
        Object.defineProperty(b.prototype, d, {
            "get": function () {
                return this.Fields[d]
            },
            "set": function (a) {
                this.Fields[d] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, c, {
            "get": function () {
                return this.Fields[c]
            },
            "set": function (a) {
                this.Fields[c] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, e, {
            "get": function () {
                return this.Fields[e]
            },
            "set": function (a) {
                this.Fields[e] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, f, {
            "get": function () {
                return this.Fields[f]
            },
            "set": function (a) {
                this.Fields[f] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, g, {
            "get": function () {
                return this.Fields[g]
            },
            "set": function (a) {
                this.Fields[g] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, h, {
            "get": function () {
                return this.Fields[h]
            },
            "set": function (a) {
                this.Fields[h] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, i, {
            "get": function () {
                return this.Fields[i]
            },
            "set": function (a) {
                this.Fields[i] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, j, {
            "get": function () {
                return this.Fields[j]
            },
            "set": function (a) {
                this.Fields[j] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, k, {
            "get": function () {
                return this.Fields[k]
            },
            "set": function (a) {
                this.Fields[k] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, l, {
            "get": function () {
                return this.Fields[l]
            },
            "set": function (a) {
                this.Fields[l] = a
            },
            enumerable: a,
            configurable: a
        });
        b.prototype.SerializeFields = function () {
            var a = this;
            a.SetSerializedField(d, a.AppInfo);
            a.SetSerializedField(c, a.AssetId);
            a.SetSerializedField(e, a.Stage1Time);
            a.SetSerializedField(f, a.Stage2Time);
            a.SetSerializedField(g, a.Stage3Time);
            a.SetSerializedField(h, a.Stage4Time);
            a.SetSerializedField(i, a.Stage5Time);
            a.SetSerializedField(j, a.Stage6Time);
            a.SetSerializedField(k, a.Stage7Time);
            a.SetSerializedField(l, a.ErrorResult)
        };
        return b
    }(d);
    e.AppLoadTimeUsageData = g;
    var f = function (o) {
        var n = "AppSizeHeight",
            m = "AppSizeWidth",
            l = "ClientId",
            k = "HostVersion",
            j = "Host",
            i = "UserId",
            h = "Browser",
            g = "AppURL",
            f = "AppId",
            e = "CorrelationId";
        __extends(d, o);

        function d() {
            o.call(this, "AppActivated")
        }
        Object.defineProperty(d.prototype, e, {
            "get": function () {
                return this.Fields[e]
            },
            "set": function (a) {
                this.Fields[e] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(d.prototype, b, {
            "get": function () {
                return this.Fields[b]
            },
            "set": function (a) {
                this.Fields[b] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(d.prototype, f, {
            "get": function () {
                return this.Fields[f]
            },
            "set": function (a) {
                this.Fields[f] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(d.prototype, g, {
            "get": function () {
                return this.Fields[g]
            },
            "set": function (a) {
                this.Fields[g] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(d.prototype, c, {
            "get": function () {
                return this.Fields[c]
            },
            "set": function (a) {
                this.Fields[c] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(d.prototype, h, {
            "get": function () {
                return this.Fields[h]
            },
            "set": function (a) {
                this.Fields[h] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(d.prototype, i, {
            "get": function () {
                return this.Fields[i]
            },
            "set": function (a) {
                this.Fields[i] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(d.prototype, j, {
            "get": function () {
                return this.Fields[j]
            },
            "set": function (a) {
                this.Fields[j] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(d.prototype, k, {
            "get": function () {
                return this.Fields[k]
            },
            "set": function (a) {
                this.Fields[k] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(d.prototype, l, {
            "get": function () {
                return this.Fields[l]
            },
            "set": function (a) {
                this.Fields[l] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(d.prototype, m, {
            "get": function () {
                return this.Fields[m]
            },
            "set": function (a) {
                this.Fields[m] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(d.prototype, n, {
            "get": function () {
                return this.Fields[n]
            },
            "set": function (a) {
                this.Fields[n] = a
            },
            enumerable: a,
            configurable: a
        });
        d.prototype.SerializeFields = function () {
            var a = this;
            a.SetSerializedField(e, a.CorrelationId);
            a.SetSerializedField(b, a.SessionId);
            a.SetSerializedField(f, a.AppId);
            a.SetSerializedField(g, a.AppURL);
            a.SetSerializedField(c, a.AssetId);
            a.SetSerializedField(h, a.Browser);
            a.SetSerializedField(i, a.UserId);
            a.SetSerializedField(j, a.Host);
            a.SetSerializedField(k, a.HostVersion);
            a.SetSerializedField(l, a.ClientId);
            a.SetSerializedField(m, a.AppSizeWidth);
            a.SetSerializedField(n, a.AppSizeHeight)
        };
        return d
    }(d);
    e.AppActivatedUsageData = f;
    var h = function (i) {
        var g = "CloseMethod",
            f = "OpenTime",
            e = "AppSizeFinalHeight",
            d = "AppSizeFinalWidth",
            c = "FocusTime";
        __extends(h, i);

        function h() {
            i.call(this, "AppClosed")
        }
        Object.defineProperty(h.prototype, b, {
            "get": function () {
                return this.Fields[b]
            },
            "set": function (a) {
                this.Fields[b] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, c, {
            "get": function () {
                return this.Fields[c]
            },
            "set": function (a) {
                this.Fields[c] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, d, {
            "get": function () {
                return this.Fields[d]
            },
            "set": function (a) {
                this.Fields[d] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, e, {
            "get": function () {
                return this.Fields[e]
            },
            "set": function (a) {
                this.Fields[e] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, f, {
            "get": function () {
                return this.Fields[f]
            },
            "set": function (a) {
                this.Fields[f] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, g, {
            "get": function () {
                return this.Fields[g]
            },
            "set": function (a) {
                this.Fields[g] = a
            },
            enumerable: a,
            configurable: a
        });
        h.prototype.SerializeFields = function () {
            var a = this;
            a.SetSerializedField(b, a.SessionId);
            a.SetSerializedField(c, a.FocusTime);
            a.SetSerializedField(d, a.AppSizeFinalWidth);
            a.SetSerializedField(e, a.AppSizeFinalHeight);
            a.SetSerializedField(f, a.OpenTime);
            a.SetSerializedField(g, a.CloseMethod)
        };
        return h
    }(d);
    e.AppClosedUsageData = h;
    var i = function (i) {
        var g = "ErrorType",
            f = "ResponseTime",
            e = "Parameters",
            d = "APIID",
            c = "APIType";
        __extends(h, i);

        function h() {
            i.call(this, "APIUsage")
        }
        Object.defineProperty(h.prototype, b, {
            "get": function () {
                return this.Fields[b]
            },
            "set": function (a) {
                this.Fields[b] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, c, {
            "get": function () {
                return this.Fields[c]
            },
            "set": function (a) {
                this.Fields[c] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, d, {
            "get": function () {
                return this.Fields[d]
            },
            "set": function (a) {
                this.Fields[d] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, e, {
            "get": function () {
                return this.Fields[e]
            },
            "set": function (a) {
                this.Fields[e] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, f, {
            "get": function () {
                return this.Fields[f]
            },
            "set": function (a) {
                this.Fields[f] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, g, {
            "get": function () {
                return this.Fields[g]
            },
            "set": function (a) {
                this.Fields[g] = a
            },
            enumerable: a,
            configurable: a
        });
        h.prototype.SerializeFields = function () {
            var a = this;
            a.SetSerializedField(b, a.SessionId);
            a.SetSerializedField(c, a.APIType);
            a.SetSerializedField(d, a.APIID);
            a.SetSerializedField(e, a.Parameters);
            a.SetSerializedField(f, a.ResponseTime);
            a.SetSerializedField(g, a.ErrorType)
        };
        return h
    }(d);
    e.APIUsageUsageData = i
})(OSFLog || (OSFLog = {}));
var Logger;
(function (a) {
    "use strict";
    (function (a) {
        a._map = [];
        a._map[0] = "info";
        a.info = 0;
        a._map[1] = "warning";
        a.warning = 1;
        a._map[2] = "error";
        a.error = 2
    })(a.TraceLevel || (a.TraceLevel = {}));
    var e = a.TraceLevel;
    (function (a) {
        a._map = [];
        a._map[0] = "none";
        a.none = 0;
        a._map[1] = "flush";
        a.flush = 1
    })(a.SendFlag || (a.SendFlag = {}));
    var f = a.SendFlag;

    function d(a, c, d) {
        if (OSF.Logger && OSF.Logger.ulsEndpoint) {
            var b = {
                    traceLevel: a,
                    message: c,
                    flag: d,
                    internalLog: true
                },
                e = JSON.stringify(b);
            OSF.Logger.ulsEndpoint.writeLog(e)
        }
    }
    a.sendLog = d;

    function b() {
        try {
            return new c
        } catch (a) {
            return null
        }
    }
    var c = function () {
        function b() {
            var a = this,
                b = a;
            a.telemetryEndPoint = "https://telemetryservice.firstpartyapps.oaspapps.com/telemetryservice/telemetryproxy.html";
            a.buffer = [];
            a.proxyFrameReady = false;
            OSF.OUtil.addEventListener(window, "message", function (a) {
                return b.tellProxyFrameReady(a)
            });
            a.loadProxyFrame()
        }
        b.prototype.writeLog = function (b) {
            var a = this;
            if (a.proxyFrameReady === true) a.proxyFrame.contentWindow.postMessage(b, "*");
            else a.buffer.length < 128 && a.buffer.push(b)
        };
        b.prototype.tellProxyFrameReady = function (d) {
            var b = this,
                g = b;
            if (d.data === "ProxyFrameReadyToLog") {
                b.proxyFrameReady = true;
                for (var c = 0; c < b.buffer.length; c++) b.writeLog(b.buffer[c]);
                b.buffer.length = 0;
                OSF.OUtil.removeEventListener(window, "message", function (a) {
                    return g.tellProxyFrameReady(a)
                })
            } else if (d.data === "ProxyFrameReadyToInit") {
                var e = {
                        appName: "Office APPs",
                        sessionId: a.Guid.generateNew()
                    },
                    f = JSON.stringify(e);
                b.proxyFrame.contentWindow.postMessage(f, "*")
            }
        };
        b.prototype.loadProxyFrame = function () {
            var a = this;
            a.proxyFrame = document.createElement("iframe");
            a.proxyFrame.setAttribute("style", "display:none");
            a.proxyFrame.setAttribute("src", a.telemetryEndPoint);
            document.head.appendChild(a.proxyFrame)
        };
        return b
    }();
    (function (c) {
        var a = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];

        function b() {
            for (var c = "", d = (new Date).getTime(), b = 0; b < 32 && d > 0; b++) {
                if (b == 8 || b == 12 || b == 16 || b == 20) c += "-";
                c += a[d % 16];
                d = Math.floor(d / 16)
            }
            for (; b < 32; b++) {
                if (b == 8 || b == 12 || b == 16 || b == 20) c += "-";
                c += a[Math.floor(Math.random() * 16)]
            }
            return c
        }
        c.generateNew = b
    })(a.Guid || (a.Guid = {}));
    var g = a.Guid;
    if (!OSF.Logger) OSF.Logger = a;
    a.ulsEndpoint = b()
})(Logger || (Logger = {}));
var OSFAppTelemetry;
(function (c) {
    var b = null;
    "use strict";
    var a, n = function () {
            function a() {}
            return a
        }(),
        d = function () {
            function a(b, a) {
                this.name = b;
                this.handler = a
            }
            return a
        }(),
        e = function () {
            function a() {
                this.clientIDKey = "Office API client";
                this.logIdSetKey = "Office App Log Id Set"
            }
            a.prototype.getClientId = function () {
                var b = this,
                    a = b.getValue(b.clientIDKey);
                if (!a || a.length <= 0 || a.length > 40) {
                    a = OSF.Logger.Guid.generateNew();
                    b.setValue(b.clientIDKey, a)
                }
                return a
            };
            a.prototype.saveLog = function (c, d) {
                var b = this,
                    a = b.getValue(b.logIdSetKey);
                a = (a && a.length > 0 ? a + ";" : "") + c;
                b.setValue(b.logIdSetKey, a);
                b.setValue(c, d)
            };
            a.prototype.enumerateLog = function (c, e) {
                var a = this,
                    d = a.getValue(a.logIdSetKey);
                if (d) {
                    var f = d.split(";");
                    for (var h in f) {
                        var b = f[h],
                            g = a.getValue(b);
                        if (g) {
                            c && c(b, g);
                            e && a.remove(b)
                        }
                    }
                    e && a.remove(a.logIdSetKey)
                }
            };
            a.prototype.getValue = function (c) {
                var a = OSF.OUtil.getLocalStorage(),
                    b = "";
                if (a) b = a.getItem(c);
                return b
            };
            a.prototype.setValue = function (c, b) {
                var a = OSF.OUtil.getLocalStorage();
                a && a.setItem(c, b)
            };
            a.prototype.remove = function (b) {
                var a = OSF.OUtil.getLocalStorage();
                if (a) try {
                    a.removeItem(b)
                } catch (c) {}
            };
            return a
        }(),
        f = function () {
            function a() {}
            a.prototype.LogData = function (a) {
                if (!OSF.Logger) return;
                OSF.Logger.sendLog(OSF.Logger.TraceLevel.info, a.SerializeRow(), OSF.Logger.SendFlag.none)
            };
            a.prototype.LogRawData = function (a) {
                if (!OSF.Logger) return;
                OSF.Logger.sendLog(OSF.Logger.TraceLevel.info, a, OSF.Logger.SendFlag.none)
            };
            return a
        }();

    function l(f) {
        if (!OSF.Logger) return;
        if (a) return;
        a = new n;
        a.sessionId = OSF.Logger.Guid.generateNew();
        a.hostVersion = f.get_appVersion();
        a.appId = f.get_id();
        a.host = f.get_appName();
        a.browser = window.navigator.userAgent;
        a.correlationId = f.get_correlationId();
        a.clientId = (new e).getClientId();
        var g = location.href.indexOf("?");
        a.appURL = g == -1 ? location.href : location.href.substring(0, g);
        (function (f, a) {
            var d, e, c;
            a.assetId = "";
            a.userId = "";
            try {
                d = decodeURIComponent(f);
                e = new DOMParser;
                c = e.parseFromString(d, "text/xml");
                a.userId = c.getElementsByTagName("t")[0].attributes.getNamedItem("cid").nodeValue;
                a.assetId = c.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue
            } catch (g) {} finally {
                d = b;
                c = b;
                e = b
            }
        })(f.get_eToken(), a);
        (function () {
            var j = new Date,
                e = new Date,
                h = 0,
                i = false,
                a = [];
            a.push(new d("focus", function () {
                e = new Date
            }));
            a.push(new d("blur", function () {
                if (e) {
                    h += Math.abs((new Date).getTime() - e.getTime());
                    e = b
                }
            }));
            var g = function () {
                for (var b = 0; b < a.length; b++) OSF.OUtil.removeEventListener(window, a[b].name, a[b].handler);
                a.length = 0;
                if (!i) {
                    c.onAppClosed(Math.abs((new Date).getTime() - j.getTime()), h);
                    i = true
                }
            };
            a.push(new d("beforeunload", g));
            a.push(new d("unload", g));
            for (var f = 0; f < a.length; f++) OSF.OUtil.addEventListener(window, a[f].name, a[f].handler)
        })();
        c.onAppActivated()
    }
    c.initialize = l;

    function g() {
        if (!a) return;
        (new e).enumerateLog(function (b, a) {
            return (new f).LogRawData(a)
        }, true);
        var b = new OSFLog.AppActivatedUsageData;
        b.SessionId = a.sessionId;
        b.AppId = a.appId;
        b.AssetId = a.assetId;
        b.AppURL = a.appURL;
        b.UserId = a.userId;
        b.ClientId = a.clientId;
        b.Browser = a.browser;
        b.Host = a.host;
        b.HostVersion = a.hostVersion;
        b.CorrelationId = a.correlationId;
        b.AppSizeWidth = window.innerWidth;
        b.AppSizeHeight = window.innerHeight;
        (new f).LogData(b)
    }
    c.onAppActivated = g;

    function m(g, h, d, c, e) {
        if (!a) return;
        var b = new OSFLog.APIUsageUsageData;
        b.SessionId = a.sessionId;
        b.APIType = g;
        b.APIID = h;
        b.Parameters = d;
        b.ResponseTime = c;
        b.ErrorType = e;
        (new f).LogData(b)
    }
    c.onCallDone = m;

    function i(g, c, e, f) {
        var a = b;
        if (c)
            if (typeof c == "number") a = String(c);
            else if (typeof c === "object")
            for (var d in c) {
                if (a !== b) a += ",";
                else a = ""; if (typeof c[d] == "number") a += String(c[d])
            } else a = "";
        OSF.AppTelemetry.onCallDone("method", g, a, e, f)
    }
    c.onMethodDone = i;

    function k(c, a) {
        OSF.AppTelemetry.onCallDone("event", c, b, 0, a)
    }
    c.onEventDone = k;

    function h(d, e, a, c) {
        OSF.AppTelemetry.onCallDone(d ? "registerevent" : "unregisterevent", e, b, a, c)
    }
    c.onRegisterDone = h;

    function j(d, c) {
        if (!a) return;
        var b = new OSFLog.AppClosedUsageData;
        b.SessionId = a.sessionId;
        b.FocusTime = c;
        b.OpenTime = d;
        b.AppSizeFinalWidth = window.innerWidth;
        b.AppSizeFinalHeight = window.innerHeight;
        (new e).saveLog(a.sessionId, b.SerializeRow())
    }
    c.onAppClosed = j;
    OSF.AppTelemetry = c
})(OSFAppTelemetry || (OSFAppTelemetry = {}));
OSF.InitializationHelper = function (d, b, f, e, c) {
    var a = this;
    a._hostInfo = d;
    a._webAppState = b;
    a._context = f;
    a._settings = e;
    a._hostFacade = c
};
OSF.InitializationHelper.prototype.getAppContext = function (v, h) {
    var i = "undefined";
    if (this._hostInfo.isRichClient) {
        var c, a = window.external.GetContext(),
            d = a.GetAppType(),
            e = false;
        for (var n in OSF.AppName)
            if (OSF.AppName[n] == d) {
                e = true;
                break
            }
        if (!e) throw "Unsupported client type " + d;
        var u = a.GetSolutionRef(),
            f;
        if (typeof a.GetApiSetVersion !== i) f = a.GetApiSetVersion();
        var q = OSF.O15HostSpecificFileVersion.GenerateVersion(a.GetAppVersionMajor(), f),
            p = a.GetAppUILocale(),
            m = a.GetAppDataLocale(),
            r = a.GetDocUrl(),
            l = a.GetAppCapabilities(),
            s = a.GetActivationMode(),
            k = a.GetControlIntegrationLevel(),
            o = [],
            b;
        try {
            b = a.GetSolutionToken()
        } catch (t) {}
        var g;
        if (typeof a.GetCorrelationId !== i) g = a.GetCorrelationId();
        b = b ? b.toString() : "";
        c = new OSF.OfficeAppContext(u, d, q, p, m, r, l, o, s, k, b, g);
        h(c);
        OSF.AppTelemetry && OSF.AppTelemetry.initialize(c)
    } else {
        var j = function (e, a) {
            var b;
            if (a._appName === OSF.AppName.ExcelWebApp) {
                var c = a._settings;
                b = {};
                for (var g in c) {
                    var f = c[g];
                    b[f[0]] = f[1]
                }
            } else b = a._settings; if (e === 0 && a._id != undefined && a._appName != undefined && a._appVersion != undefined && a._appUILocale != undefined && a._dataLocale != undefined && a._docUrl != undefined && a._clientMode != undefined && a._settings != undefined && a._reason != undefined) {
                var d = new OSF.OfficeAppContext(a._id, a._appName, a._appVersion, a._appUILocale, a._dataLocale, a._docUrl, a._clientMode, b, a._reason, a._osfControlType, a._eToken, a._correlationId);
                h(d);
                OSF.AppTelemetry && OSF.AppTelemetry.initialize(d)
            } else throw "Function ContextActivationManager_getAppContextAsync call failed. ErrorCode is " + e
        };
        this._webAppState.clientEndPoint.invoke("ContextActivationManager_getAppContextAsync", j, this._webAppState.id)
    }
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication = function () {
    var d = "ContextActivationManager_notifyHost",
        b = false,
        a = this,
        e = OSF.OUtil.parseXdmInfo();
    if (e) {
        a._hostInfo.isRichClient = b;
        var c = e.split("|");
        if (c != undefined && c.length === 3) {
            a._webAppState.conversationID = c[0];
            a._webAppState.id = c[1];
            a._webAppState.webAppUrl = c[2]
        }
    } else a._hostInfo.isRichClient = true; if (!a._hostInfo.isRichClient) {
        a._webAppState.clientEndPoint = Microsoft.Office.Common.XdmCommunicationManager.connect(a._webAppState.conversationID, a._webAppState.wnd, a._webAppState.webAppUrl);
        a._webAppState.serviceEndPoint = Microsoft.Office.Common.XdmCommunicationManager.createServiceEndPoint(a._webAppState.id);
        var f = a._webAppState.conversationID + OSF.SharedConstants.NotificationConversationIdSuffix;
        a._webAppState.serviceEndPoint.registerConversation(f);
        var g = function (c) {
            switch (c) {
            case OSF.AgaveHostAction.Select:
                a._webAppState.focused = true;
                window.focus();
                break;
            case OSF.AgaveHostAction.UnSelect:
                a._webAppState.focused = b;
                break;
            default:
                Sys.Debug.trace("actionId " + c + " notifyAgave is wrong.")
            }
        };
        a._webAppState.serviceEndPoint.registerMethod("Office_notifyAgave", g, Microsoft.Office.Common.InvokeType.async, b);
        window.onfocus = function () {
            if (!a._webAppState.focused) {
                a._webAppState.focused = true;
                a._webAppState.clientEndPoint.invoke(d, null, [a._webAppState.id, OSF.AgaveHostAction.Select])
            }
        };
        window.onblur = function () {
            if (a._webAppState.focused) {
                a._webAppState.focused = b;
                a._webAppState.clientEndPoint.invoke(d, null, [a._webAppState.id, OSF.AgaveHostAction.UnSelect])
            }
        }
    }
};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize = function (a) {
    var e = new OSF.DDA.License(a.get_eToken()),
        f = window.open;
    window.open = function (d, c, b) {
        var a = null;
        try {
            a = f(d, c, b)
        } catch (g) {}
        if (!a) {
            var e = {
                strUrl: d,
                strWindowName: c,
                strWindowFeatures: b
            };
            OSF._OfficeAppFactory.getClientEndPoint().invoke("ContextActivationManager_openWindowInHost", null, e)
        }
        return a
    };
    if (a.get_appName() == OSF.AppName.OutlookWebApp || a.get_appName() == OSF.AppName.Outlook) {
        OSF._OfficeAppFactory.setContext(new OSF.DDA.OutlookContext(a, this._settings, e, a.appOM));
        Microsoft.Office.WebExtension.initialize()
    } else if (a.get_osfControlType() === OSF.OsfControlType.DocumentLevel || a.get_osfControlType() === OSF.OsfControlType.ContainerLevel) {
        OSF._OfficeAppFactory.setContext(new OSF.DDA.Context(a, a.doc, e));
        var b, c, d = a.get_reason();
        if (this._hostInfo.isRichClient) {
            b = OSF.DDA.DispIdHost.getRichClientDelegateMethods;
            c = OSF.DDA.SafeArray.Delegate.ParameterMap;
            d = OSF.DDA.RichInitializationReason[d]
        } else {
            b = OSF.DDA.DispIdHost.getXLSDelegateMethods;
            c = OSF.DDA.XLS.Delegate.ParameterMap
        }
        OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(b, c));
        Microsoft.Office.WebExtension.initialize(d)
    } else throw OSF.OUtil.formatString(Strings.OfficeOM.L_OsfControlTypeNotSupported)
};
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM = function (a, e, p) {
    var c = false,
        b;
    b = ".js";
    var m = {
        "1-15.00": "excel-15" + b,
        "1-15.01": "excel-15.01" + b,
        "2-15.00": "word-15" + b,
        "2-15.01": "word-15.01" + b,
        "4-15.00": "powerpoint-15" + b,
        "4-15.01": "powerpoint-15.01" + b,
        "8-15.00": "outlook-15" + b,
        "8-15.01": "outlook-15.01" + b,
        "16-15": "excelwebapp-15" + b,
        "16-15.01": "excelwebapp-15.01" + b,
        "64-15": "outlookwebapp-15" + b,
        "64-15.01": "outlookwebapp-15.01" + b,
        "128-15.00": "project-15" + b,
        "128-15.01": "project-15.01" + b
    };
    OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);
    var h = this;

    function f(e) {
        var c, b;
        if (h._hostInfo.isRichClient) b = OSF.DDA.RichClientSettingsManager.read();
        else b = a.get_settings();
        var d = OSF.DDA.SettingsManager.deserializeSettings(b);
        if (e) c = new OSF.DDA.RefreshableSettings(d);
        else c = new OSF.DDA.Settings(d);
        return c
    }
    var g = a.get_appVersion();
    if (g > OSF.O15HostSpecificFileVersion.Fallback) g = OSF.O15HostSpecificFileVersion.Fallback;
    var d = p + m[a.get_appName() + "-" + g];
    if (a.get_appName() == OSF.AppName.Excel) {
        var l = function () {
            a.doc = new OSF.DDA.ExcelDocument(a, f(c));
            e()
        };
        OSF.OUtil.loadScript(d, l)
    } else if (a.get_appName() == OSF.AppName.ExcelWebApp) {
        var i = function () {
            a.doc = new OSF.DDA.ExcelWebAppDocument(a, f(true));
            e()
        };
        OSF.OUtil.loadScript(d, i)
    } else if (a.get_appName() == OSF.AppName.Word) {
        var o = function () {
            a.doc = new OSF.DDA.WordDocument(a, f(c));
            e()
        };
        OSF.OUtil.loadScript(d, o)
    } else if (a.get_appName() == OSF.AppName.PowerPoint) {
        var j = function () {
            a.doc = new OSF.DDA.PowerPointDocument(a, f(c));
            e()
        };
        OSF.OUtil.loadScript(d, j)
    } else if (a.get_appName() == OSF.AppName.OutlookWebApp || a.get_appName() == OSF.AppName.Outlook) {
        var k = function () {
            h._settings = f(c);
            a.appOM = new OSF.DDA.OutlookAppOm(a, h._webAppState.wnd, e)
        };
        OSF.OUtil.loadScript(d, k)
    } else if (a.get_appName() == OSF.AppName.Project) {
        var n = function () {
            a.doc = new OSF.DDA.ProjectDocument(a);
            e()
        };
        OSF.OUtil.loadScript(d, n)
    } else throw OSF.OUtil.formatString(stringNS.L_AppNotExistInitializeNotCalled, a.get_appName())
};