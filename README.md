In my experience, these above modules are quite useful to automate stuff. With stand-alone, they are not really related to each other. But when combining, those modules could work well in some contexts.

**Installation:**
Import ```.bas``` and ```.frm``` files into your project (Open VBA Editor, Alt + F11; File > Import File)

**Features:**
+ Office version: [2013, 2016, 2019 and 365](https://www.office.com/)

See how it works below

### Folder Loop
----------

Loop through each file in a folder/subfolder recursively to get information inside. Can be used in many cases such as copy value from the same formatted files, reshape, or simply just open then close them.

**Modules**: [LoopFolder.bas](https://github.com/ngtridung97/Vba/blob/master/LoopFolder.bas), [VisibleCell.bas](https://github.com/ngtridung97/Vba/blob/master/VisibleCell.bas)

### Outlook and SAP Control
----------

Send an email containing [FB03 SAP Tcode](http://www.saptransactions.com/codes/FB03/) (Company Code, Fiscal Year, Document Number) to another account (usually to a PC can run SAP). Then open SAP and auto-input those strings in order to download hard copies. Finally, zip and reply to the receipt email.

**Modules**: [OutlookControl.bas](https://github.com/ngtridung97/Vba/blob/master/OutlookControl.bas), [SAPControl.bas](https://github.com/ngtridung97/Vba/blob/master/SAPControl.bas), [ZipFiles.bas](https://github.com/ngtridung97/Vba/blob/master/ZipFiles.bas)

### ADOdb Query
----------
Communication between Excel and Database via ```ADOdb``` connection.

**Modules**: [EventMenu.bas](https://github.com/ngtridung97/Vba/blob/master/EventMenu.bas), [ReduceSize.bas](https://github.com/ngtridung97/Vba/blob/master/ReduceSize.bas)

### Drag and Drop .eml file
----------
List email information and Move/Copy them to another directory by ```Treeview Nodes``` in Userform

**Modules**:

### Update modified data to server
----------
Loop through each row in selected range (visible cells only), push update to server, pull the newest table back and resotre filters.

**Modules**:

### Fill in the blank cells
----------
Missing data in some cells. Sort reference columns and input IF function into blank cells.

**Modules**:

### Feedback & Suggestions
----------
Please feel free to fork, comment or give feedback to ng.tridung97@gmail.com
