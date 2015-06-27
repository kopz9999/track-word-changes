var fs = require('fs-extra');
var path = require('path');
var WordDocumentHandler = require('../lib/word-document-handler');
var AdmZip = require('adm-zip');
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

var waitForFileOperations = 2000;

describe("WordDocumentHandler", function () {
  beforeEach( function() {
    fs.removeSync( absoluteSandboxPath() );
    fs.ensureDirSync( absoluteSandboxPath() );
  });

  afterEach(function() {
    fs.removeSync( absoluteSandboxPath() );
  });
  
  // Validations
  describe("validations", function () {
    describe('validate file presence', function(){
      it("should be valid when file exists", function () {
        var destinationPath = resolveSandboxFile('testfile.docx');
        var fileTrackerPtr = new WordDocumentHandler({
          filePath: destinationPath 
        });
        fs.copySync(resolveFixtureFile('testfile.docx'), destinationPath );
        expect( fileTrackerPtr.isValid() ).toBeTruthy();
      });
      it("should be valid when file does not exists", function () {
        var destinationPath = resolveSandboxFile('testfile.docx');
        var fileTrackerPtr = new WordDocumentHandler({
          filePath: destinationPath 
        });
        expect( fileTrackerPtr.isValid() ).toBeFalsy();
        expect( fileTrackerPtr.errors.length ).toEqual(1);
        expect( fileTrackerPtr.errors[0].message )
          .toEqual('Word document not found');
      });
    });

    describe('validate file extension', function(){
      it("should be valid when file is .docx", function () {
        var destinationPath = resolveSandboxFile('testfile.docx');
        var fileTrackerPtr = new WordDocumentHandler({
          filePath: destinationPath 
        });
        fs.copySync(resolveFixtureFile('testfile.docx'), destinationPath );
        expect( fileTrackerPtr.isValid() ).toBeTruthy();
      });
      it("should be valid when file is .doc", function () {
        var destinationPath = resolveSandboxFile('testfileold.doc');
        var fileTrackerPtr = new WordDocumentHandler({
          filePath: destinationPath 
        });
        fs.copySync(resolveFixtureFile('testfileold.doc'), destinationPath );
        expect( fileTrackerPtr.isValid() ).toBeTruthy();
      });
      it("should be valid when file does not exists", function () {
        var destinationPath = resolveSandboxFile('testfile.txt');
        var fileTrackerPtr = new WordDocumentHandler({
          filePath: destinationPath 
        });
        fs.closeSync(fs.openSync(destinationPath, 'w'));
        
        expect( fileTrackerPtr.isValid() ).toBeFalsy();
        expect( fileTrackerPtr.errors.length ).toEqual(1);
        expect( fileTrackerPtr.errors[0].message )
          .toEqual('Document has not valid a extension. '+
            'The only extensions available are: .doc,.docx');
      });
    });
  });

  // Check that old word document 1997-2003 can be converted
  describe("assert document conversion", function(){
    describe("with a valid document", function(){
      it("should be valid when file is .doc", function () {
        var flag = false;
        var destinationPath = resolveSandboxFile('testfileold.doc');
        var fileTrackerPtr = new WordDocumentHandler({
          filePath: destinationPath 
        });
        fs.copySync(resolveFixtureFile('testfileold.doc'), destinationPath );
        fileTrackerPtr.onSuccessCallback = function() {
          flag = true;
        };
        runs(function() {
          fileTrackerPtr.trackChanges(true);
        });
        waitsFor(function() {
          return flag;
        }, "File operations should have finished", waitForFileOperations);
        runs(function() {
          var entries = fs.readdirSync( absoluteSandboxPath() );
          var exists = entries.any(function(el) { 
            return el == 'testfileold.docx'; 
            
          });
          expect(exists).toBeTruthy();
          expect(entries.length).toEqual(2);
        });
      });
      // Run this test when unoconv is down
      xit("should have errors when unoconv is down", function(){
        var flag = false;
        var destinationPath = resolveSandboxFile('testfileold.doc');
        var fileTrackerPtr = new WordDocumentHandler({
          filePath: destinationPath 
        });
        fs.copySync(resolveFixtureFile('testfileold.doc'), destinationPath );
        fileTrackerPtr.onErrorCallback = function() {
          flag = true;
        };
        runs(function() {
          fileTrackerPtr.trackChanges(true);
        });
        waitsFor(function() {
          return flag;
        }, "File operations should have finished", waitForFileOperations);
        runs(function() {
          var entries = fs.readdirSync( absoluteSandboxPath() );
          expect(entries.length).toEqual(1);
          expect( fileTrackerPtr.errors.length ).toEqual(1);
          expect( fileTrackerPtr.errors[0].message )
            .toEqual("File couldn't be converted to docx");
        });
      });
    });

  });
  
  // Using overwrite and turn on-off options
  describe("using different options", function(){
    it("should create a copy when deactivating overwrite", function () {
      var flag = false;
      var destinationPath = resolveSandboxFile('testfile.docx');
      var fileTrackerPtr = new WordDocumentHandler({
        filePath: destinationPath,
        overWrite: false
      });
      fs.copySync(resolveFixtureFile('testfile.docx'), destinationPath );
      fileTrackerPtr.onSuccessCallback = function() {
        flag = true;
      };
      runs(function() {
        fileTrackerPtr.trackChanges(true);
      });
      waitsFor(function() {
        return flag;
      }, "File operations should have finished", waitForFileOperations);
      runs(function() {
        var entries = fs.readdirSync( absoluteSandboxPath() );
        expect(entries.length).toEqual(2);
      });
    });
    describe('turn on track changes', function() {
      describe("document with track changes off", function() {
        it("should turn on track changes", function () {
          var flag = false;
          var destinationPath = resolveSandboxFile('testfile.docx');
          var fileTrackerPtr = new WordDocumentHandler({
            filePath: destinationPath 
          });
          fs.copySync(resolveFixtureFile('testfile.docx'), destinationPath );
          fileTrackerPtr.onSuccessCallback = function() {
            flag = true;
          };
          runs(function() {
            fileTrackerPtr.trackChanges(true);
          });
          waitsFor(function() {
            return flag;
          }, "File operations should have finished", waitForFileOperations);
          runs(function() {
            var zip = new AdmZip( destinationPath );
            var docFolder = resolveSandboxFile( 'testfileFolder' );
            var settingsPath = path.join(docFolder, 'word', 'settings.xml');
            var fileString = null;
            
            zip.extractAllTo(docFolder, true);
            fileString = fs.readFileSync(settingsPath, "utf8");
            expect( fileString.indexOf("<w:trackRevisions/>") > -1 ).toBeTruthy();
          });
        });
      })
      describe("document with track changes on", function() {
        it("should not modify document when it has track changes on", function () {
          var flag = false;
          var destinationPath = resolveSandboxFile('testfile2.docx');
          var fileTrackerPtr = new WordDocumentHandler({
            filePath: destinationPath 
          });
          fs.copySync(resolveFixtureFile('testfile2.docx'), destinationPath );
          fileTrackerPtr.onSuccessCallback = function() {
            flag = true;
          };
          runs(function() {
            spyOn(fileTrackerPtr, '__turnOnTrackChanges').andCallThrough();        
            fileTrackerPtr.trackChanges(true);
          });
          waitsFor(function() {
            return flag;
          }, "File operations should have finished", waitForFileOperations);
          runs(function() {
            expect(fileTrackerPtr.__turnOnTrackChanges).not.toHaveBeenCalled();
          });
        });
      })
    });
    describe('turn off track changes', function() {
      describe("document with track changes off", function() {
        it("should not modify document when it has track changes off", function () {
          var flag = false;
          var destinationPath = resolveSandboxFile('testfile.docx');
          var fileTrackerPtr = new WordDocumentHandler({
            filePath: destinationPath 
          });
          fs.copySync(resolveFixtureFile('testfile.docx'), destinationPath );
          fileTrackerPtr.onSuccessCallback = function() {
            flag = true;
          };
          runs(function() {
            spyOn(fileTrackerPtr, '__turnOffTrackChanges').andCallThrough();        
            fileTrackerPtr.trackChanges(false);
          });
          waitsFor(function() {
            return flag;
          }, "File operations should have finished", waitForFileOperations);
          runs(function() {
            expect(fileTrackerPtr.__turnOffTrackChanges).not.toHaveBeenCalled();
          });
        });
      })
      describe("document with track changes on", function() {
        it("should turn off track changes", function () {
          var flag = false;
          var destinationPath = resolveSandboxFile('testfile2.docx');
          var fileTrackerPtr = new WordDocumentHandler({
            filePath: destinationPath 
          });
          fs.copySync(resolveFixtureFile('testfile2.docx'), destinationPath );
          fileTrackerPtr.onSuccessCallback = function() {
            flag = true;
          };
          runs(function() {
            fileTrackerPtr.trackChanges(false);
          });
          waitsFor(function() {
            return flag;
          }, "File operations should have finished", waitForFileOperations);
          runs(function() {
            var zip = new AdmZip( destinationPath );
            var docFolder = resolveSandboxFile( 'testfileFolder' );
            var settingsPath = path.join(docFolder, 'word', 'settings.xml');
            var fileString = null;
            
            zip.extractAllTo(docFolder, true);
            fileString = fs.readFileSync(settingsPath, "utf8");
            expect( fileString.indexOf("<w:trackRevisions/>") <= -1 ).toBeTruthy();
          });          
        });
      })
    });

  });

}); 