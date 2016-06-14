/**
* utilities for apps script 
* these typically dont do the same thing
* just some alternative
* @namespace Utilities
*/
var Utilities = (function (ns) {

    /**
     * encode a string to base 64
     * @param {string} str the string
     * @return {string} the encoded string
     */
    ns.base64Encode = function (str) {
        return btoa(str);
    };

    // just to avoid reference errors in apps script converts
    ns.DigestAlgorithm = {};
    /**
     * this just is going to compute a sha1 hash
     * might add to it later if proper digest is needed
     * @param {string} algo the digest algorithm . just ignored for now
     * @param {string} str the string to be digested
     * @return {string} the digest
     */
    ns.computeDigest = function (algo, str) {
        var hash = CryptoJS.SHA1(str);
        return hash.toString();
    }
  
    return ns;
})(Utilities || {});