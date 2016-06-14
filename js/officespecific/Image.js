// responsible for placing a png image on a sheet
var Image = (function(image) {
  'use strict';
  
  // insert an image at row / column in the sheet
  image.insert = function ( range,  png , offx, offy) {
    // range is ignored for Office version - it will insert at active cell

    return new Promise (function (resolve, reject){
      Office.context.document.setSelectedDataAsync (png, {
        coercionType:Office.CoercionType.Image,
        imageLeft:offx || 0,
        imageRight:offy || 0
      },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject (result.error.message);
        }
        else {
          resolve (result);
        }
      })
    });
    
  };
  
  // place an image
  image.place = function (png) {
    return image.insert (null, png , 10,6);
  }
  
  return image;
}) (Image || {});


