#! /usr/bin/env node

// Required classes
var WordDocumentHandler = require('./lib/word_document_handler');
var minimist = require('minimist');

// Logic Here:
var userArgs = minimist(process.argv.slice(2));
var handlers = [];
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
  handlers.push(
    new WordDocumentHandler({
      filePath: filePath,
      overWrite: overwrite
    })
  );
});

// Process files
handlers.forEach(function( tracker ) {
  revisionTrackOpt = userArgs['revision_track']
  revisionTrack = revisionTrackOpt === undefined ? true : 
    ( revisionTrackOpt == 'true' );
  tracker.trackChanges( revisionTrack );
  if (tracker.errors.length > 0 ) {
    printWordDocumentHandlerError(tracker);
  } else {
    printWordDocumentHandlerSuccess(tracker);
  }
});
