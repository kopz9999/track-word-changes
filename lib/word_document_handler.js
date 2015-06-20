var fs = require('fs');
var express = require('express');
var i18n = require("i18n");

i18n.configure({
  locales:['en'],
  directory: __dirname + '/locales'
});

WordDocumentHandler = function(filePath) {
  this.filePath = filePath;
  this.workPath = null;
  this.workDir = null;
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
        _self.__unzip();
      }
    });
  }
};

WordDocumentHandler.prototype.__unzip = function(callback) {
  var AdmZip = require('adm-zip');
  var zip = new AdmZip( this.workPath );
  zip.extractAllTo(this.workDir, true);
};

WordDocumentHandler.prototype.__createCopy = function(opts) {
  var dt = new Date();
  var normalPath = this.filePath.substr(0, this.filePath.lastIndexOf('.'));
  var pipeRequest = null;

  this.workDir = normalPath + '_x'; //+ dt.toString();
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
