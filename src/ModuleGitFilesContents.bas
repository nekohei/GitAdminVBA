Attribute VB_Name = "ModuleGitFilesContents"
Rem .gitattributes
'# Auto detect text files and perform LF normalization
'* text=auto
'
'*.bas       text    eol=crlf
'*.cls       text    eol=crlf
'*.frm       text    eol=crlf
'*.frx       binary  eol=crlf
'*.dcm       text    eol=crlf
'*.vbaproj   text    eol=crlf
'
'*.wsf       text    eol=crlf
'*.bat       text    eol=crlf
'
'*.cls linguist-language=VBA
'*.dcm linguist-language=VBA
'*.vbaproj linguist-language=INI
'
'# file encording
'*.bas working-tree-encoding=UTF-8
'*.cls working-tree-encoding=UTF-8
'*.dcm working-tree-encoding=UTF-8
'*.frm working-tree-encoding=UTF-8
'
'*.bas encoding=UTF-8
'*.cls encoding=UTF-8
'*.dcm encoding=UTF-8
'*.frm encoding=UTF-8
'
'*.bas diff=UTF-8
'*.cls diff=UTF-8
'*.dcm diff=UTF-8
'*.frm diff=UTF-8

Rem .gitignore
'*.tmp
'*.xl*
'‾$*.xl*
'bin/old
'!bin/*.xl*
'!src/*

Rem settings.json
'{
'  "[markdown]": {
'    "editor.wordWrap": "on",
'    "editor.quickSuggestions": {
'      "comments": "off",
'      "strings": "off",
'      "other": "off"
'    },
'    "files.encoding": "utf8",
'  },
'  "files.encoding": "shiftjis",
'  "files.associations": {
'    "*.bas": "vb",
'    "*.cls": "vb",
'    "*.dcm": "vb",
'    "*.frm": "vb"
'  }
'}
