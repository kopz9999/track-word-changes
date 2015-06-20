var fs = require('fs-extra');
var express = require('express');
var i18n = require("i18n");
var path = require('path');
var xpath = require('xpath');
var dom = require('xmldom').DOMParser;
var AdmZip = require('adm-zip');
var exec = require('child_process').exec;

i18n.configure({
  locales:['en'],
  directory: __dirname + '/locales'
});

WordDocumentHandler = function(opts) {
  this.filePath = opts.filePath;
  this.overWrite = opts.overWrite === undefined ? false : opts.overWrite;
  this.workPath = null;
  this.workDir = null;
  this.settingsPath = null;
};

WordDocumentHandler.prototype._runValidations = function() {
  this.errors = [];
  if (!fs.existsSync( this.filePath )) {
    this.errors.push( new WordDocumentHandler.Error( 
        i18n.__('errors.file_not_found', this.filePath) 
      )
    );
  }
};

WordDocumentHandler.prototype.isValid = function() {
  this._runValidations();
  return this.errors.length == 0;
};

WordDocumentHandler.prototype.trackChanges = function(state) {
  if (this.isValid()) {
    var _self = this;
    this.__createCopy({ 
      success: function() {
        _self.__dotrackChanges(state);
      }
    });
  }
};

WordDocumentHandler.prototype.__dotrackChanges = function(state) {
  var file = null, doc = null, nodeQuery = null, revisionNodes = null, 
    rootNode = null;
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
    if (removeWorkPath) _self._removeWorkPath();
  });
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
  var normalPath = this.filePath.substr(0, this.filePath.lastIndexOf('.'));
  var pipeRequest = null;

  this.workDir = (normalPath + '_' + dt.toString())
    .replace(/(\:|\s|\(|\))/g, '_');

  this.workPath = this.workDir + '.zip';
  pipeRequest = fs.createWriteStream(this.workPath)
  fs.createReadStream(this.filePath)
    .pipe(pipeRequest);

  pipeRequest.on('finish', function () {
    opts.success();
  });

};

WordDocumentHandler.Error = function(message) {
  this.message = message;
};

module.exports = WordDocumentHandler;
