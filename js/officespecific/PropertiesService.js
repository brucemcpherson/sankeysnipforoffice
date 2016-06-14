/**  polyfill for Apps Script Properties service
* only does user and document properties
* @namespace PropertiesService
*/
var PropertiesService = (function (ps) {

    // uses for document properties
    ps.getDocumentProperties = function () {
        return {
            setProperty: function (key, value) {
                Office.context.document.settings.set(key, value);
                return ps.flushDocumentProperties();
            },
            getProperty: function (key) {
                return Office.context.document.settings.get(key);
            },
            deleteProperty: function (key) {
                Office.context.document.settings.remove(key);
                return ps.flushDocumentProperties();
            }
        }
    };

    // the settings only write to in memory copy
    // need to write async to doc for permamnence
    ps.flushDocumentProperties = function () {
        return new Promise(function (resolve, reject) {
            Office.context.document.settings.saveAsync(function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    reject(asyncResult);
                }
                else {
                    resolve(asyncResult); 
                }
            });
        });
    }

    // uses for user properties
    ps.getUserProperties = function () {
        return {
            setProperty: function (key, value) {
                return localStorage.setItem(key, value);
            },
            getProperty: function (key) {
                return localStorage.getItem(key);
            },
            deleteProperty: function (key) {
                return localStorage.removeItem(key);
            }
        }
    };
    return ps;

})(PropertiesService || {});
