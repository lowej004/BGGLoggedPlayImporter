# BGGLoggedPlayImporter
A Google Sheets script to import logged play data from BoardGameGeek

Installation Instructions:
1) Create a new Google Sheets Document
2) Select 'Script Editor' from the 'Tools' menu
3) In the 'Code.gs' file which has now opened, copy the contents of 'BGGImporter.gs' and save
4) Repeat this step for each of the files in the solution
5) Register an application with BoardGameGeek to get an authorisation token (details can be found here as to how to do this: https://boardgamegeek.com/using_the_xml_api)
6) Update BGGCollectionImporter.gs line 14 and BGGPlayImporter.gs line 36 to insert the bearer token that you have been given from BoardGameGeek
7) Close the script editor and reload the spreadsheet

Walkthrough Instructions:
1) A new menu option will appear titled 'BGG Logged Play Import'
2) Select 'Run Play Import' from the new menu
3a) The first time you run this you will need to walk through some Google authorizations to allow this script to run
3b) Any time after the first time you open this option your previous options will be remembered
4) A dialog will appear, select the relevant options and click 'Import'
5) The dialog will close and the logged play data will be imported 
