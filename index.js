#! /usr/bin/env node

// Required classes
var WordDocumentHandler = require('./lib/word_document_handler');

// Logic Here:
var userArgs = process.argv.slice(2);
var handlers = [];
var fileTrackerPtr = null;

userArgs.forEach(function(filePath) {
  handlers.push( new WordDocumentHandler(filePath) );
});

// Process files
handlers.forEach(function( tracker ) {
  tracker.trackChanges(true);
  if (tracker.errors.length > 0 ) {
    tracker.errors.forEach(function( errObj ){
      console.log(errObj.message);
    });
  }
});
