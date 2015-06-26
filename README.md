# Word Document Tracker

This repository contains a command line tool that turns on track changes for a word document.

# Requirements

In order to provide compatibility with Word 93-2003, you need the *unoconv* utility. If you are working with Ubuntu, you must install the following command line utility: [unoconv](https://apps.ubuntu.com/cat/applications/unoconv)

## Mac and Windows

If you are working on MAC or Windows, you need to install [Libre Office](https://www.libreoffice.org). After installing, clone the [unoconv repository](https://github.com/dagwieers/unoconv). At this point, you must create the following environment variable storing the location of the unoconv python script: 

```
$ export UNOCONV_PATH=/path-of-unoconv-python-script
```

After that, you must install one of the following scripts in the user binaries:
- [Mac Shell](shell/unoconv.sh)
- [Windows Bat](shell/unoconv.bat)

### Unoconv Notes

Please note that you need to **keep unoconv running in background to be able to use it as other user**. In order to open a port, you must be logged as sudo:

```
$ sudo unoconv --listener
```

# Basic Usage

```
$ npm install --global kopz9999/track_word_changes
$ track_word path-to-my-file/Example.docx
```

"Example.docx" will have **turned on track changes**. Everytime you edit the file, a comment will appear.

# Options

You can also provide an option to turn on/off revision tracking by passing true/false in **revision_track**

```
$ track_word --revision_track=false path-to-my-file/Example.docx
```

The above line turns off track changes for Word. This options is true by default, so everytime you run the command you are asking to turn on track changes.

Another available option is **overwrite**, which is true by default since it applies the changes to the specified file. If you pass a false value to this, option you will get a copy of the file with the desired setting:

```
$ track_word --overwrite=false path-to-my-file/Example.docx
```

The above line will create a copy of Example.docx with name **Example_{timestamp}.docx**

You can take a look at the demo video to check functionality: [Demo](https://www.dropbox.com/s/523c3osmtj66zqf/Demo.mov?dl=0)

### Notes

The path *spec/fixtures/files* contains sample files to test:
- testfile.docx
  - Word document without revision enabled
- testfile2.docx
  - Word document with revision enabled
- testfileold.doc
  - Word document with 1997-2003 format