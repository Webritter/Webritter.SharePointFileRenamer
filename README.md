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
### Update title and filename in document by selected properties and Id
- New documents in the library are named by Word-Online as "Document1", "Document2", ....

- In the document library is a custom propertie with the type of the Document("Concept", "Manual", "Report")

- The filenames should be named as "Concept(1)", 2Manual(2)", "Report(3)"

You have to setup a filter for the documents to select only files starting with "Document" and query the ID and the category of the document. With this two query fields you can format the title "{1}({0})" and the filename. 




```xml
   <RunOptionsTask Id="0" Name="Rename Drafts" Enabled="true">
      <LibraryName>Documents</LibraryName>
      <CamlQuery>
        &lt;Where&gt;
          &lt;BeginsWith&gt;
            &lt;FieldRef Name='FileLeafRef' /&gt;
            &lt;Value Type='File'&gt;Document&lt;/Value&gt;
          &lt;/BeginsWith&gt;
        &lt;/Where&gt;
      </CamlQuery>	
      <FileNameFormat>{1}({0})</FileNameFormat>
      <QueryFields>
        <QueryFieldOptions Name="ID" ShouldNotBeNull="true" />
        <QueryFieldOptions Name="Category" ShouldNotBeNull="true" />
      </QueryFields>
      <UpdateFields>
        <UpdateFieldOptions Name="Title" Format="{1}({0})" />
      </UpdateFields>
      <CheckinMessage>File Renamed</CheckinMessage>
      <CheckinType>OverwriteCheckIn</CheckinType>
    </RunOptionsTask>

```

### Rename all files by selection in a dropdwn field (category) and unique id
### Move all approved files to a subfolder selected by a managed metadata field
### Update a text property with the version of the document


## Update document propertivalue of dropes


