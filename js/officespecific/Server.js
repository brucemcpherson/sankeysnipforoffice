
/**
* used to expose memebers of a namespace
* @param {string||object} namespace name
* @param {method} method name
* this is not really needed for Office version
* but it preserves the same structure between apps script & office
* also helps turn every operation into a promise, used with Provoke.
*/
function exposeRun(namespace, method, argArray) {

    if (!namespace || typeof namespace === "string") {
        var func = (namespace ? this[namespace][method] : this[method]);
    }
    else {
        var func = namespace[method];
    }
    if (argArray && argArray.length) {
        return func.apply(this, argArray);
    }
    else {
        return func();
    }

}

/**
* wrapper for office api functions
* @namespace Server
*/
var Server = (function (ns) {

    ns.control = {

        current: {
            sheetName: "",
            data: {},
            name: 'sankeySnipCurrent',
            binding: null
        },
        selected: {
            data: {}
        },
        watching: {
            watcher: null
        }
    };

    ns.isOffice = function () {
        return !Utils.isUndefined(Office);
    };
    
    /**
     * check for feature support of host app
     */
    ns.featureBehavior = function () {
      // dont support close button at all
      DomUtils.hide(Process.control.buttons.close ,true);
      
      // inserting images may not be supported
      if (!Office.context.requirements.isSetSupported('ImageCorecion','1.1')) {
          DomUtils.hide(Process.control.buttons.insert,true);
          App.showNotification("Feature unavailable in this version of Office", 
          "Chart will be shown as a preview only as its not possible to embed images in a sheet in your version of Excel");
      }
    };
    /**
     * in office, we dont need to pollFrequency
     * so we'll modify the watcher to be passive
     * it'll just need a poke each time ofice detects a change
     */
    ns.pollingBehavior = function () {
        // we' going to tweak the watcher for office
        Server.control.watching.watcher = Process.control.watching.watcher;
        Server.control.watching.watcher.getWatching().pollFrequency = 0;
    };
    /**
     * initislize all the binding and watching
     */
    ns.initialize = function () {

        // keep an eye on the selection having changed
        ns.watchForSelectionChanges();

        // set up the current worksheet as data source
        ns.getCurrentWorksheet();


    };


    /** 
     * if the active selection changes, then we need to check that the worksheet is still the same one
     */
    ns.watchForSelectionChanges = function () {
        Office.context.document.addHandlerAsync("documentSelectionChanged", function (e) {
            // get current worksheet and set up watching
            ns.getCurrentWorksheet();
        });
    };

    ns.storeData = function (eData, scope) {

        scope.data = {
            values: eData.values,
            checksum: Utils.keyDigest(eData.values),
            range: eData
        };
        return scope;
    };

    /**
     * returns a promise to values and other info for the data range
     * of the active sheet
     * Onlys selected object properties will be loaded including values
     * @return {Promise} to the result
     */
    ns.getDataRange = function (type) {
        return new Promise(function (resolve, reject) {
            Excel.run(function (ctx) {
                var usedRange = ctx.workbook.worksheets
                    .getActiveWorksheet()
                    .getUsedRange(true)
                    .load("values,rowIndex,rowCount,columnIndex,columnCount,address");
                return ctx.sync().then(function () {
                    resolve(usedRange);
                });
            })
        });
    };


    /**
     * returns a promise to values and other info for the selected range
     * of the active sheet
     * Onlys selected object properties will be loaded including values
     * @return {Promise} to the result
     */
    ns.getActiveRange = function (type) {
        return new Promise(function (resolve, reject) {
            Excel.run(function (ctx) {
                var activeRange = ctx.workbook
                    .getSelectedRange(true)
                    .load("values,rowIndex,rowCount,columnIndex,columnCount,address");
                return ctx.sync().then(function () {
                    resolve(activeRange);
                });
            })
        });
    };

    /**
     * store the used range for the current sheet
     * @return {Promise} used range promise
     */
    ns.storeUsedRange = function () {
        var scope = ns.control.current;
        return ns.getDataRange().then(function (rangeData) {
            ns.storeData(rangeData, scope);
        });

    };
    /**
     * store the active range for the current sheet
     * @return {Promise} used range promise
     */
    ns.storeActiveRange = function () {
        var scope = ns.control.selected;
        return ns.getActiveRange().then(function (rangeData) {
            ns.storeData(rangeData, scope);
            // always do a poke because if we're getting the active range, then something may have changed
            if (ns.control.watching.watcher) {
                ns.control.watching.watcher.poke();
            }
        });

    };
    ns.generateTestData = function (data) {

        return ns.addSheet()
        .then(function (sheet) {
            return Excel.run(function (ctx) {
                var address = 'A1:' + Utils.columnLabelMaker(data[0].length) + data.length;
                var range = ctx.workbook.worksheets.getItem(sheet.name).getRange(address);
                range.values = data;
                return ctx.sync()
                .then (function (e) {
                    return Excel.run(function (ctx) {
                        ctx.workbook.worksheets.getItem(sheet.name).activate();
                        return ctx.sync();
                    });
                });
            });
        });

    };

    ns.addSheet = function (sheetName) {
        return new Promise(function (resolve, reject) {
            Excel.run(function (ctx) {
                var sheet = ctx.workbook.worksheets.add(sheetName).load("name");
                return ctx.sync().then(function () {
                    resolve(sheet);
                });
            })

        });
    };

    /**
     * remove binding handler, if there is one
     * @param {object} scope where to find the binding info
     * @return {Promise} when done
     */
    ns.removeDataChangeHandler = function (scope) {

        return new Promise(function (resolve, reject) {
            if (scope.binding) {
                scope.binding.removeHandlerAsync(Office.EventType.BindingDataChanged, function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result);
                    }
                    else {
                        reject(result);
                    }
                });
            }
            else {
                // there isntt one.. but that's just fine too.
                resolve();
            }
        });
    };


    /**
    * add data change handler to binding 
    * @param {Binding} binding the binding to add to
    * @param {function} handler the handler
    * @return {Promise} to when its done
    */
    ns.addHandler = function (binding, handler) {
        return new Promise(function (resolve, reject) {
            binding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, {}, function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result);
                }
                else {
                    reject(result);
                }
            });
        });
    };

    /**
     * add binding 
     * @param {string} theRange the range to bind to
     * @param {string} name the binding name
     * @return {Promise} to when its done
     */
    ns.addBinding = function (theRange, name) {
        return new Promise(function (resolve, reject) {
            Office.context.document.bindings.addFromNamedItemAsync(theRange, Office.BindingType.Matrix, {
                id: name
            }, function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result);
                }
                else {
                    reject(result);
                }
            });
        });
    };
    /** 
     * bind to a given range
     * @param {string} theRange an a1 style range
     * @return {Promise} when its done
     */
    ns.bindToRange = function (theRange, scope) {

        // first need to cancel any outstanding handlers
        return ns.removeDataChangeHandler(scope)
            .then(function () {
                return ns.addBinding(theRange, scope.name);
            })
            .then(function (result) {
                scope.binding = result.value;
                return ns.addHandler(scope.binding, function (e) {
                    ns.storeUsedRange();
                    ns.storeActiveRange();
                });
            });

    };
    /**
    * Get the active worksheet
    * @return {Promise} to the sheet
    */
    ns.getActiveWorksheet = function () {
        return new Promise(function (resolve, reject) {
            Excel.run(function (ctx) {
                // ask for load the name property of the current sheet
                var sheet = ctx.workbook.worksheets.getActiveWorksheet().load("name");

                // sync to get the data from the other side
                return ctx.sync()
                    .then(function () {
                        resolve(sheet);
                    });
            });
        });
    };

    /**
     * get the current worksheet
     * if its not the same as before, then get the used range, the active range, and watch out for changes
     */
    ns.getCurrentWorksheet = function () {
        var scope = ns.control.current;

        // first get the active worksheet
        ns.getActiveWorksheet()
            .then(function (sheet) {
                if (!scope.sheetName || scope.sheetName !== sheet.name) {
                    scope.sheetName = sheet.name;
                    // get the current data for this sheet, and set up a new binding
                    return Promise.all([ns.storeUsedRange(), ns.storeActiveRange(), ns.bindToRange(sheet.name + "!a:z", scope)])
                }
                else {
                    return ns.storeActiveRange();
                };
            });
    };


    return ns;
})(Server || {});


