# SharePointFileRenamer
A simple commandline application to rename and move files in SharePoint document libraries. 

## Features
- Update Document properties
- Filter docuements by CAML-Query
- Format the filename by properties
- Move the filename to another folder given by properties
- Run multiple Tasks in document library (ec. rename and move)
- Logging with log4net

## use cases
- Go through files in a document library and update some properties in the items and/or rename the files by values of the properties.
- Go through files and move them to different folders.


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
      <CheckinMessage>Draft File Renamed</CheckinMessage>
      <CheckinType>OverwriteCheckIn</CheckinType>
    </RunOptionsTask>

```
### rename all not approved files in the library with DRAFT_ and a unique id


### Move all approved files to a subfolder selected by a managed metadata field




### Update a text property with the version of the document in all not approved files

In your document library you have a text field "DocumentVersion" and have used it in word.exe as a field. This text-Field should be updated with the next published version.

``` xml
   <RunOptionsTask Id="0" Name="VersionUpdater" Enabled="false">
      <LibraryName>Documents</LibraryName>
      <CamlQuery>
        &lt;Where&gt;
          &lt;And&gt;
            &lt;Or&gt;
              &lt;Eq&gt;
                &lt;FieldRef Name='_ModerationStatus' /&gt;
                &lt;Value Type='ModStat'&gt;2&lt;/Value&gt;
              &lt;/Eq&gt;
              &lt;Eq&gt;
                &lt;FieldRef Name='_ModerationStatus' /&gt;
	        &lt;Value Type='ModStat'&gt;3&lt;/Value&gt;
	      &lt;/Eq&gt;
 	    &lt;/Or&gt;
	    &lt;And&gt;
	      &lt;Eq&gt;
	        &lt;FieldRef Name='FSObjType' /&gt;
	        &lt;Value Type='Integer'&gt;0&lt;/Value&gt;
	      &lt;/Eq&gt;
	        &lt;IsNull&gt;
	        &lt;FieldRef Name='CheckoutUser' /&gt;
	      &lt;/IsNull&gt;
	    &lt;/And&gt;
	  &lt;/And&gt;
	&lt;/Where&gt;	
      </CamlQuery>
      <QueryFields>
        <QueryFieldOptions Name="_UIVersionString" ShouldNotBeNull="true" />
      </QueryFields>
      <UpdateFields>
        <UpdateFieldOptions Name="DocumentVersion" Format="V{0}" />
      </UpdateFields>
      <CheckinMessage>Version Updated</CheckinMessage>
      <CheckinType>OverwriteCheckIn</CheckinType>
    </RunOptionsTask>
 
```




## Update document propertivalue of dropes


