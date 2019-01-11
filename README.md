# SharePointFileRenamer
A simple commandline application to rename and move files in SharePoint document libraries. 

## Features
- Update Document properties
- Filter docuements by CAML-Query
- Format the filename by properties
- Move the filename to another folder given by properties
- run multiple Tasks in document library (ec. rename and move)

## use cases
Go through all files in a document library and update some fields (Title)
### Update title in document by selected properties and Id
New documents in the library are named from Word-Online as "Document1", "Document2", ....
In the document library is a custom propertie with the type of the Document("Concept", "Manual", "Report")
The filenames should be named as "Concept(1)", 2Manual(2)", "Report(3)"

### Rename all files by selection in a dropdwn field (category) and unique id
### Move all approved files to a subfolder selected by a managed metadata field
### Update a text property with the version of the document


## Update document propertivalue of dropes


