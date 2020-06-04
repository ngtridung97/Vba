In my experience, these above modules are quite useful to automate stuff. With stand-alone, they are not really related to each other. But when combining, those modules could work well in some contexts.

**Installation:**
Import ```.bas``` and ```.frm``` files into your project (Open VBA Editor, Alt + F11; File > Import File)

**Features:**
+ Office version: [2013, 2016, 2019 and 365](https://www.office.com/)

See how it works below

### Folder Loop
----------

**Context**: Loop through every file in a folder/subfolder recursively to get information inside. It can be used in many contexts such as copy value from the same formatted files, reshape, or simply just open then close them.

**Modules used**: [LoopFolder.bas](https://github.com/ngtridung97/Vba/blob/master/LoopFolder.bas), [VisibleCell.bas](https://github.com/ngtridung97/Vba/blob/master/VisibleCell.bas)

### Outlook and SAP Control
----------

**Context**: Send an email containing [FB03 SAP Tcode](http://www.saptransactions.com/codes/FB03/) (Company Code, Fiscal Year, Document Number) to another account (usually from PC can run SAP). Then open SAP and auto-input those strings in order to download hard copies. Finally, zip and reply to the receipt email.

**Modules used**: [OutlookControl.bas](https://github.com/ngtridung97/Vba/blob/master/OutlookControl.bas), [SAPControl.bas](https://github.com/ngtridung97/Vba/blob/master/SAPControl.bas), [ZipFiles.bas](https://github.com/ngtridung97/Vba/blob/master/ZipFiles.bas)

### ADOdb Query
----------
**Context**: Communication between Excel and Database via ```ADOdb``` connection.

**Modules used**: [EventMenu.bas](https://github.com/ngtridung97/Vba/blob/master/EventMenu.bas), [ReduceSize.bas](https://github.com/ngtridung97/Vba/blob/master/ReduceSize.bas)

### Drag and Drop .eml file
----------
**Context**: Move/Copy email to another directory by using ```Treeview Nodes``` in Userform

**Modules used**:

### Feedback & Suggestions
----------
Please feel free to fork, comment or give feedback to ng.tridung97@gmail.com
