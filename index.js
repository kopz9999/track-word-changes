#! /usr/bin/env node

// Required classes
var WordDocumentHandler = require('./lib/word-document-handler');
var minimist = require('minimist');

// Logic Here:
var userArgs = minimist(process.argv.slice(2));
var fileTrackerPtr = null;

// Swap variables
var overwriteOpt, overwrite, revisionTrackOpt, revisionTrack;

// Helper functions
var printWordDocumentHandlerSuccess = function(wordDocumentHandler) {
  var str = "Successfully Processed: " + wordDocumentHandler.filePath;
  console.log(str);
};

var printWordDocumentHandlerError = function(wordDocumentHandler) {
  var str = "Failed to Process: " + wordDocumentHandler.filePath;
  console.log(str);
  wordDocumentHandler.errors.forEach(function( errObj ){
    console.log(errObj.message);
  });
};

userArgs['_'].forEach(function(filePath) {
  overwriteOpt = userArgs['overwrite'];
  overwrite = overwriteOpt === undefined ? true : 
    ( overwriteOpt == 'true' );
  revisionTrackOpt = userArgs['revision_track']
  revisionTrack = revisionTrackOpt === undefined ? true : 
    ( revisionTrackOpt == 'true' );
  fileTrackerPtr = new WordDocumentHandler({
    filePath: filePath,
    overWrite: overwrite,
    onSuccessCallback: printWordDocumentHandlerSuccess,
    onErrorCallback: printWordDocumentHandlerError
  });

  fileTrackerPtr.trackChanges( revisionTrack );
  
});
