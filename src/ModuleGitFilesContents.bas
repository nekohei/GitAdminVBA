Attribute VB_Name = "ModuleGitFilesContents"
Rem .gitattributes
'# Auto detect text files and perform LF normalization
'* text=auto
'
'*.bas       text    eol=crlf
'*.cls       text    eol=crlf
'*.frm       text    eol=crlf
'*.frx       binary
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
'# file encording (local=CP932, Git internal=UTF-8)
'*.bas working-tree-encoding=CP932
'*.cls working-tree-encoding=CP932
'*.dcm working-tree-encoding=CP932
'*.frm working-tree-encoding=CP932

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
'  "files.encoding": "cp932",
'  "files.associations": {
'    "*.bas": "vb",
'    "*.cls": "vb",
'    "*.dcm": "vb",
'    "*.frm": "vb"
'  }
'}
