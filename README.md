# BGGLoggedPlayImporter
A Google Sheets script to import logged play data from BoardGameGeek

Installation Instructions:
1) Create a new Google Sheets Document
2) Select 'Script Editor' from the 'Tools' menu
3) In the 'Code.gs' file which has now opened, copy the contents of 'BGGImporter.gs' and save
4) Add a new HTML file from the 'File -> New' menu
5) Rename the file to 'BGGImporterDialog.html' and copy the contents of 'BGGImportDialog.html' into the file and save
6) Close the script editor and reload the spreadsheet

Walkthrough Instructions:
1) A new menu option will appear titled 'BGG Logged Play Import'
2) Select 'Run Import' from the new menu
3a) The first time you run this you will need to walk through some authorization to allow this to run
3b) Any time after the first time you open this option your previous options will be remembered
4) A dialog will appear, select the relevant options and click 'Import'
5) The dialog will close and the logged play data will be imported 
