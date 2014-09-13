dwg2csv
=======

Export data from dwg files via AutoCAD ActiveX API

This set of Python scripts and other files was born in 2011, when I was working on a task with goal to migrate some data from AutoCAD to a GIS solution.
For data access interface I had chosen AutoCAD ActiveX API (that was a mistake, after all).
For data processing and scripting the Python was selected because it's my favorite programming language.

Export from DWG (AutoCAD) to CSV.

COM (ActiveX) AutoCAD API used for access to drawing entities and for extracting geometry and attributes from them.

Python modules `comtypes.client`, `comtypes.gen` used for interaction with AutoCAD.
Module `csv.writer` used for data export.

For make it work you will need MS Windows, Autodesk AutoCAD, Python >= 2.5.

Contents:

* dwg.dump.py -- exporting program, work with current AutoCAD drawing unless you're run this script with a parameter: filename.dwg.
* snippets.py -- AutoCAD ActiveX objects wrapper.
* trig.py -- functions for coordinates transformation and other math stuff.
* dwg.list -- list of input dwg files example.
* rip.cmd -- runner cmd script example.
* test.py -- tests for recovery DWG entities from exported data.
* ora/csv.lob2ora.py -- CSV to Oracle loader, load exported from DWG data to Oracle DB using cx_Oracle. For coords data CLOB field was used because of data size.
* ora/csv2ora.cmd -- cmd script for csv2ora loader.

extra/autocad.bulge/FacetBulge_rev1.zip -- some formulas (partly correct) from
http://www.cadtutor.net/forum/showthread.php?51511-Points-along-a-lwpoly-arc

http://vasnake.blogspot.com/2011/06/export-from-dwg-to-csv.html
http://vasnake.blogspot.com/search/label/Autodesk
