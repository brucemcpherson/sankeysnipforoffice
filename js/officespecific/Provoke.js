/**
 * @namespace Provoke
 * this is a polyfill to simulate what happens in apps script
 * although many of the operations could be synch, im treating them all as async
 * promise management for async calls
 */

var Provoke = (function (ns) {

  /**
  * run something asynchronously
  * @param {string} namespace the namespace (null for global)
  * @param {string} method the method or function to call
  * @param {[...]} the args
  * @return {Promise} a promise
  */
  ns.run = function (namespace, method) {

    // the args to the server function
    var runArgs = Array.prototype.slice.call(arguments).slice(2);

    if (arguments.length < 2) {
      throw new Error('need at least a namespace and method');
    }

    // this will return a promise
    return new Promise(function (resolve, reject) {

      try {
        var result = exposeRun(namespace, method, runArgs);
        
        // the result might be a promise
        if (result instanceof Promise) {
          result.then(function (p) {
            resolve(p);
          })
          .catch(function(err) {
            reject(err);
          });
        }
        else {
          // it wasnt a promise, so just resolve with the returned result
          resolve(result);
        }
      }

      catch (err) {
        reject(err);
      };


    });

  };

  return ns;

})(Provoke || {});



