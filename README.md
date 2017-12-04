# BISM SSAS Tabular Documentation Tool
 Creates an Excel Document with the metadata of all objects in a SSAS Tabular Model.

 The following objects metadata will be documented in an excel file

 1) Server
 2) Dimensions
 3) DimensionAttributes
 4) Relationships
 5) Hierarchies
 6) Measures
 7) KPIs
 8) Partitions
 9) Perspectives
 10) Roles

#### PreRequisites

1. .Net Framework 4.6 and above
 2. Microsoft® SQL Server® 2016 Analysis Management Objects which can be downloaded from the below link https://www.microsoft.com/en-us/download/confirmation.aspx?id=52676
 3. Office 2010 or Above.

 #### How to Use the tool

1. Run the BISMDocumenter.exe.
2. Enter the ServerName and Click Connect.
3. The list box will be loaded with all the database and cube in the format of Database Name * CubeName.
4. Select the Cube Name to be documented.
5. By Default output path will be a folder "Output" in executable path.If required modify the path.
6. Click Generate Document Button.
7. For each cube selected one excel will be generated in the output path specified.


 Note : Please don't remove \ rename the template folder and the excel file in it. The tool depends on these files.
 
 #### This project has been modified to use for SQL Server 2016 and 2017
