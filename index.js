#! /usr/bin/env node

// Required classes
var WordDocumentHandler = require('./lib/word_document_handler');
var minimist = require('minimist');

// Logic Here:
var userArgs = minimist(process.argv.slice(2));
var handlers = [];
var fileTrackerPtr = null;

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
    tracker.errors.forEach(function( errObj ){
      console.log(errObj.message);
    });
  }
});
