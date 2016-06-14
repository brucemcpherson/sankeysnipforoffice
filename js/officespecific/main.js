/**
* Sankey snip app
*/

(function () {
    
    // each initialize fights with each other so do them separately and check theyve completed
    var goo = new Promise(function (resolve,reject){
        google.load('visualization', '1.1', { 'packages': ['sankey'] });
        google.setOnLoadCallback(function () {
            resolve();
        });
    });
    
    Office.initialize = function (reason) {
        // also need the google stuff to have loaded
        goo.then(function() {
            Server.initialize();
            App.initialize();
            
            Process.initialize().then(function () {
                // set any listeners
                Home.initialize();
                
                // disable any features not supported by the host version of Office
                Server.featureBehavior();
                
                // chaneg the polling behavior for Office
                Server.pollingBehavior();
                
                // watch for changes
                Client.start();
                

            });
        })
        .catch(function(err) {
            throw err;
        });
    };

})();
