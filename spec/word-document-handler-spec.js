var fs = require('fs-extra');
var path = require('path');
var WordDocumentHandler = require('../lib/word-document-handler');
require('js-array-extensions');

var fileSpecPath = '../spec/fixtures/files/';
var sandboxPath = './sandbox/';

var resolveFixtureFile = function(relativePath) {
  return path.resolve(__dirname, fileSpecPath + relativePath );
};

var resolveSandboxFile = function(relativePath) {
  return path.resolve(__dirname, sandboxPath + relativePath );
};

var absoluteSandboxPath = function() {
  return path.resolve(__dirname, sandboxPath );
};

describe("WordDocumentHandler", function () {
  beforeEach( function() {
    fs.ensureDirSync( absoluteSandboxPath() );
  });

  afterEach(function() {
    fs.removeSync( absoluteSandboxPath() );
  });
    
  describe("validations", function () {
    describe('validate file presence', function(){
      it("should be valid when file exists", function () {
        var destinationPath = resolveSandboxFile('testfile.docx');
        var fileTrackerPtr = new WordDocumentHandler({filePath: destinationPath });
        fs.copySync(resolveFixtureFile('testfile.docx'), destinationPath );
        expect( fileTrackerPtr.isValid() ).toBeTruthy();
      });
      it("should be valid when file does not exists", function () {
        var destinationPath = resolveSandboxFile('testfile.docx');
        var fileTrackerPtr = new WordDocumentHandler({filePath: destinationPath });
        expect( fileTrackerPtr.isValid() ).toBeFalsy();
        expect( fileTrackerPtr.errors.length == 1 ).toBeTruthy();
        expect( fileTrackerPtr.errors[0].message )
          .toEqual('Word document not found');
      });
    });

    describe('validate file extension', function(){
      it("should be valid when file is .docx", function () {
        var destinationPath = resolveSandboxFile('testfile.docx');
        var fileTrackerPtr = new WordDocumentHandler({filePath: destinationPath });
        fs.copySync(resolveFixtureFile('testfile.docx'), destinationPath );
        expect( fileTrackerPtr.isValid() ).toBeTruthy();
      });
      it("should be valid when file is .doc", function () {
        var destinationPath = resolveSandboxFile('testfileold.doc');
        var fileTrackerPtr = new WordDocumentHandler({filePath: destinationPath });
        fs.copySync(resolveFixtureFile('testfileold.doc'), destinationPath );
        expect( fileTrackerPtr.isValid() ).toBeTruthy();
      });
      it("should be valid when file does not exists", function () {
        var destinationPath = resolveSandboxFile('testfile.txt');
        var fileTrackerPtr = new WordDocumentHandler({filePath: destinationPath });
        fs.closeSync(fs.openSync(destinationPath, 'w'));
        
        expect( fileTrackerPtr.isValid() ).toBeFalsy();
        expect( fileTrackerPtr.errors.length == 1 ).toBeTruthy();
        expect( fileTrackerPtr.errors[0].message )
          .toEqual('Document has not valid a extension. '+
            'The only extensions available are: .doc,.docx');
      });
    });
  });

  describe("assert document type", function(){
    describe("with a valid document", function(){
      it("should be valid when file is .doc", function () {
        var flag = false;
        var destinationPath = resolveSandboxFile('testfileold.doc');
        var fileTrackerPtr = new WordDocumentHandler({filePath: destinationPath });
        fs.copySync(resolveFixtureFile('testfileold.doc'), destinationPath );

        runs(function() {
          fileTrackerPtr.__assertDocumentType({ 
            success: function() {
              flag = true;
            }
          });
        });

        waitsFor(function() {
          return flag;
        }, "File operations should have finished", 500);
        
        runs(function() {
          var entries = fs.readdirSync( absoluteSandboxPath() );
          var exists = entries.any(function(el) { return el == 'testfileold.docx'; });
          expect(exists).toBeTruthy();
        });

      });
      

    });

  });


  describe("creates a copy", function(){
    it("should be valid when file is .doc", function () {
      var flag = false;
      var destinationPath = resolveSandboxFile('testfileold.doc');
      var fileTrackerPtr = new WordDocumentHandler({filePath: destinationPath });
      fs.copySync(resolveFixtureFile('testfileold.doc'), destinationPath );

      runs(function() {
        fileTrackerPtr.__assertDocumentType({ 
          success: function() {
            flag = true;
          }
        });
      });

      waitsFor(function() {
        return flag;
      }, "File operations should have finished", 500);
      
      runs(function() {
        var entries = fs.readdirSync( absoluteSandboxPath() );
        var exists = entries.any(function(el) { return el == 'testfileold.docx'; });
        expect(exists).toBeTruthy();
      });

    });
    

  });


}); 