
/**
 * simulate binding with apps script
 * various changes server side can be watched for server side
 * and resolved client side
 * This is the office version so of course it does support binding
 * however I want to limit the number of the graph replots
 * so i'll use the same polling approach as in apps script
 * I'll set up binding here to insert new data
 * and this watcher will return it if there is any
 * @constructor SeverBinder
 */
var ServerWatcher = (function (ns) {

/**
 * checks to see if there's been any binding callbacks since the last time
 * @param {object} watch what we're watching
 */
 ns.poll = function (watch) {
    
    // this is only  a minimal implementatino of serverwatcher
    // it only watches for data in the active or sheet level for now
    // its not fully clientwatcher compliant, as the Google Sheets one is
    // since this app only needs a couple of things
    
    // start building the result
    var pack = {
      checksum:watch.checksum,
      changed:{}
    };

    // get data if requested
    if (watch.watch.data) {
      var scope = watch.domain.scope === "Active" ? Server.control.selected : Server.control.current;
      var values = scope.data.values;
      // need to play with values in case its a single cell
      if (values && values.length && !Array.isArray(values[0])){
        values = [values];
      }
      var cs = scope.data.checksum;
      pack.changed.data = cs !== pack.checksum.data;
      if (pack.changed.data) {
        pack.data = values;
        pack.checksum.data = cs;
      }
    }

    return pack;
    
    
 };
 
  return ns;
})(ServerWatcher || {});


/**
   * polled every now and again to report back on changes
   * @param {object} watch instructions on what to check
   * @retun {object} updated status
 
  ns.poll = function (watch) {
    
    // get the active stuff
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getActiveSheet();
    var aRange = ss.getActiveRange();

    // first select the sheet .. given or active
    var s = watch.domain.sheet ? ss.getSheetByName(watch.domain.sheet) : sh;
    
    // if the scope is "sheet", then it will always be the datarange used
    if (watch.domain.scope === "Sheet") {
      var r = s.getDataRange();
    }
    
    // the scope is range - if there's a given range use it - otherwise use the datarange on the selected sheet
    else if (watch.domain.scope === "Range") {
      var r = (watch.domain.range ? sh.getRange(watch.domain.range) : sh).getDataRange();
    }
    
    // regardless of any other settings always use the active range
    else if (watch.domain.scope === "Active") {
      var r = aRange;
    }
    
    // otherwise its a mess up
    else {
      throw 'scope ' + watch.domain.scope + ' is not valid scope - should be Sheet, Range or Active';
    }
    
    // start building the result
    var pack = {
      checksum:watch.checksum,
      changed:{}
    };

    // get data if requested
    if (watch.watch.data) {
      var values = r['get'+watch.domain.property]();
      var cs = Utils.keyDigest(values);
      pack.changed.data = cs !== pack.checksum.data;
      if (pack.changed.data) {
        pack.data = values;
        pack.checksum.data = cs;
      }
    }
    
    // provide sheets if requested
    if (watch.watch.sheets) {
      var sheets = ss.getSheets().map(function(d) { return d.getName(); });
      var cs = Utils.keyDigest(sheets);
      pack.changed.sheets = cs !== pack.checksum.sheets;
      if (pack.changed.sheets) {
        pack.sheets = sheets;
        pack.checksum.sheets = cs;
      }
    }
    
    // provide active if requested
    if (watch.watch.active) {
      var a = {
        id:ss.getId(),
        sheet:sh.getName(),
        range:aRange.getA1Notation(),
        dataRange:sh.getDataRange().getA1Notation(),
        dimensions: {
          numRows : aRange.getNumRows(),
          numColumns : aRange.getNumColumns(),
          rowOffset : aRange.getRowIndex(),
          colOffset : aRange.getColumn()
        }
      }
      var cs = Utils.keyDigest (a);
      pack.changed.active = cs !== pack.checksum.active;
      if (pack.changed.active) {
        pack.active = a;
        pack.checksum.active = cs;
      }
      
    }
    return pack;
    */