# Word Document Tracker

This repository contains a command line tool that turns on track changes for a word document.

# Basic Usage

```
$ git clone https://github.com/kopz9999/track_word_changes
$ cd track_word_changes
$ npm link
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