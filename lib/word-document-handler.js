var fs = require('fs-extra');
var express = require('express');
var i18n = require("i18n");
var path = require('path');
var xpath = require('xpath');
var dom = require('xmldom').DOMParser;
var AdmZip = require('adm-zip');
var exec = require('child_process').exec;
require('js-array-extensions');

i18n.configure({
  locales:['en'],
  directory: __dirname + '/locales'
});

/**
* This class contains the logic to enable tracking on documents
*
* @class WordDocumentHandler
* @constructor
*/

var WordDocumentHandler = function(opts) {
  this.filePath = opts.filePath;
  this.overWrite = opts.overWrite === undefined ? false : opts.overWrite;
  this.onSuccessCallback = opts.onSuccessCallback;
  this.onErrorCallback = opts.onErrorCallback;
  this.fileExtension = path.extname(this.filePath);
  this.normalPath = this.filePath.substr(0, this.filePath.lastIndexOf('.'));
  this.workPath = null;
  this.workDir = null;
  this.settingsPath = null;
};

/**
* This set of functions is related to callbacks
*
* @!group Callbacks
*/

WordDocumentHandler.prototype.onError = function() {
  if (typeof( this.onErrorCallback ) === "function") {
    this.onErrorCallback( this );    
  }
};

WordDocumentHandler.prototype.onSuccess = function() {
  if (typeof( this.onSuccessCallback ) === "function") {
    this.onSuccessCallback( this );    
  }
};

/**
* @!endgroup 
*/


/**
* This set of functions is related to validations
*
* @!group Validations
*/

WordDocumentHandler.prototype._runValidations = function() {
  this._validateFilePresence();
  this._validateExtension();
};

WordDocumentHandler.prototype._validateFilePresence = function() {
  if (!fs.existsSync( this.filePath )) {
    this.errors.push( new WordDocumentHandler.Error( 
        i18n.__('errors.file_not_found') 
      )
    );
  }
};

WordDocumentHandler.prototype._validateExtension = function() {
  var _self = this;
  var result = WordDocumentHandler.config
                .validDocumentExtensions
                .any(function(el) { return el == _self.fileExtension; });
  var allowedExtensions = null;
  if (!result) {
    allowedExtensions = WordDocumentHandler.config
                          .validDocumentExtensions.join( ',' );
    this.errors.push( new WordDocumentHandler.Error( 
        i18n.__('errors.invalid_extension', allowedExtensions)
      )
    );
  }
  
};

WordDocumentHandler.prototype.isValid = function() {
  this.errors = [];
  this._runValidations();
  return this.errors.length == 0;
};

/**
* @!endgroup 
*/

WordDocumentHandler.prototype.trackChanges = function(state) {
  var _self = this;
  if (this.isValid()) {
    this.__assertDocumentType({ 
      success: function() {
        _self.__dotrackChanges(state);
      }
    });
  } else this.onError();
};

/**
* This set of functions is related to convert a document from .doc to .docx
*
* @!group DocumentConvert
*/

WordDocumentHandler.prototype.__assertDocumentType = function(opts) {
  var _self = this, fileName, dirName, convertName, convertCommand;
  var result = WordDocumentHandler.config
                .convertableDocumentExtensions
                .any(function(el) { return el == _self.fileExtension; });

  if (result) {
    _self = this;
    fileName = path.basename( this.filePath );
    dirName = path.dirname( this.filePath );
    convertName = path.basename(this.normalPath) + '.docx';
    
    convertCommand = 'cd '+ dirName +' && (unoconv --stdout -f docx '+ 
      fileName+' > '+convertName+' )';
    exec(convertCommand, function (error, stdout, stderr) {
      _self.__onDocumentConvert(opts);
    });
  } else {
    this.__createCopy( opts );
  }
  
};

WordDocumentHandler.prototype.__onDocumentConvert = function(opts) {
  var expectedPath = this.normalPath + '.docx';
  if (!fs.existsSync( expectedPath )) {
    this.errors.push( new WordDocumentHandler.Error( 
        i18n.__('errors.conversion_failed', 'docx')
      )
    );
    this.onError();
  } else {
    this.filePath = expectedPath;
    this.__createCopy( opts );
  }
};

/**
* @!endgroup 
*/

/**
* This set of functions is related to turn on track changes
*
* @!group TrackChanges
*/

WordDocumentHandler.prototype.__dotrackChanges = function(state) {
  var file = null, doc = null, nodeQuery = null, revisionNodes = null, 
    rootNode = null, hasTrackRevisions;
  this.__unzip();
  this.settingsPath = path.join(this.workDir, 'word', 'settings.xml');

  file = fs.readFileSync(this.settingsPath, "utf8");
  doc = new dom().parseFromString(file);
  nodeQuery = xpath.useNamespaces({
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  });
  revisionNodes = nodeQuery('//w:settings/w:trackRevisions', doc);
  rootNode = nodeQuery('//w:settings', doc);
  hasTrackRevisions = revisionNodes.length > 0;

  if (state) {
    if (!hasTrackRevisions) this.__turnOnTrackChanges(rootNode);
  }
  else {
    if (hasTrackRevisions) this.__turnOffTrackChanges(rootNode, revisionNodes);
  }
  this._saveChanges(true);
};

WordDocumentHandler.prototype.__turnOnTrackChanges = function(rootNode) {
  var doc = new dom().parseFromString('<w:trackRevisions/>',
    'text/xml');
  var _self = this;
  rootNode[0].appendChild(doc);
  fs.writeFileSync(this.settingsPath, rootNode[0].parentNode.toString());
};

WordDocumentHandler.prototype.__turnOffTrackChanges = function(rootNode, revisionNodes) {
  rootNode[0].removeChild( revisionNodes[0] );
  fs.writeFileSync(this.settingsPath, rootNode[0].parentNode.toString());
};

/**
* @!endgroup 
*/

/**
* This set of functions is related to workspace management for file
*
* @!group WorkspaceManagement
*/

WordDocumentHandler.prototype._saveChanges = function(removeWorkPath) {
  var destinationPath, files, directoryEntries, zipCommand, _self = this;

  destinationPath = this.overWrite ? this.filePath : 
    (path.join('..', path.basename( this.workDir + '.docx' )));

  destinationPath = path.join('..', 
    path.basename( this.overWrite ? this.filePath : (this.workDir + '.docx') ) );

  files = [ 
    destinationPath,
    "./_rels",
    "./[Content_Types].xml",
    "./docProps",
    "./word"
  ];
  directoryEntries = files.join(' ');

  zipCommand = 'cd '+ this.workDir +' && zip -r '+ directoryEntries
    + ' && cd ' + process.cwd();

  exec(zipCommand, function (error, stdout, stderr) {
    _self._onCompressSuccess( removeWorkPath );
  });
};

WordDocumentHandler.prototype._onCompressSuccess = function(removeWorkPath) {
  if (removeWorkPath) {
    this._removeWorkPath();
  }
  this.onSuccess();
};

WordDocumentHandler.prototype._removeWorkPath = function() {
  fs.unlinkSync( this.workPath );
  fs.removeSync( this.workDir );
};

WordDocumentHandler.prototype.__unzip = function() {
  var zip = new AdmZip( this.workPath );
  zip.extractAllTo(this.workDir, true);
};

WordDocumentHandler.prototype.__createCopy = function(opts) {
  var dt = new Date();
  var pipeRequest = null;

  this.workDir = (this.normalPath + '_' + dt.toString())
    .replace(/(\:|\s|\(|\))/g, '_');

  this.workPath = this.workDir + '.zip';
  pipeRequest = fs.createWriteStream(this.workPath)
  fs.createReadStream(this.filePath)
    .pipe(pipeRequest);

  pipeRequest.on('finish', function () {
    opts.success();
  });

};

/**
* @!endgroup 
*/

// Error Class

WordDocumentHandler.Error = function(message) {
  this.message = message;
};

// Configuration

WordDocumentHandler.config = {
  validDocumentExtensions: ['.doc', '.docx'],
  convertableDocumentExtensions: ['.doc']
};


module.exports = WordDocumentHandler;
