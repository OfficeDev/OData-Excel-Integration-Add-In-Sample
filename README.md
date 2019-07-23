---
page_type: sample
products:
- office-365
- office-excel
languages:
- typescript
- csharp
extensions:
  contentType: samples
  technologies:
  - Add-ins
  - Azure AD
  createdDate: 8/25/2015 7:34:03 AM
---
﻿# OData-Excel-Integration-App-Sample
This sample is an app for Office that reads and writes OData data to documents in an Office host application, such as Excel. This article provides sample code and procedures that show you how to design an app for office. The languages used are C# and TypeScript. The tools used to setup the sample are Visual Studio and the Azure Portal.

**Table of contents:**

[Deploy the sample app](#DeployTheSampleApp)<BR>
[Key components of the sample](#KeyComponents)<BR>
[Modify the Sample for your needs](#ModifySample)<BR>
 
<a name="DeployTheSampleApp"></a>
Use the steps in this section to test and debug the app.

1.	Open the sample and open the ExcelODataInterface.sln file in Visual Studio.
2.	Click Start or press F5 in Visual Studio.
3.	The first time that you use F5, you are prompted to grant permissions to the app. Click Trust It.
4.	In Start Action, select Internet Explorer to use an Office Online client, or select Office Desktop Client, then select Start Document to specify the kind of Office document to test with.
5.	The app page will show at the Excel right pane, choose “Products”, then, click “Connect”.
6.	The data will load in the Excel sheet, we can delete, compare and save data by the buttons in our app.
 
<a name="KeyComponents"></a>
#Key components of the sample
The Visual Studio solution contains the following:
- ODSampleData project, which contains the app's manifest configured to support hosting the app in Excel 2013, Excel Online.
- ODSampleData Web project, which contains the following components:
   home.aspx   The main page of the app
   ODataHelper.cs   A C# file to consume and update data using the OData. It contains the following parts: 
     Parsing metadata from OData
     Getting data form OData 
     Updating data to Odata 
- DataHelper.ts   A TypeScript file is implemented based on Javascript API for Office. It’s running on client side.  It contains the following parts:
   Data methods which are designed to interact with Excel data
   Format methods which can set data format
   Navigation Methods
   Error handler methods
- Diff.ts  A TypeScript file for solving differences
- UX.ts , UXHelpers.ts, UX.BulgingDiffPage.ts, UX.DiffPage, UXList.ts   for UI element and data object

<a name="ModifySample"></a>
# Modify the sample for your needs
The following procedures can help you to use your own data source.

1.	Open the web.config file in the ODSampleDataWeb project and change the '<add key="ida:ODataEndpointURL" value=" " />' 
2.	 Set the value of your data source url.

In our sample: 
 ```
<add key="ida:ODataEndpointURL"  
value="http://services.odata.org/V3/(S(omlwdrfviuvthgrncrmyko1m))/OData/OData.svc/" />
<add key="ida:ODataMetadataURL"     value="http://services.odata.org/V3/(S(omlwdrfviuvthgrncrmyko1m))/OData/OData.svc/$metadata" />
```
 The first value is the data source location, the second one is the data format.
 
You can use your own data by replace the values here.


# Please follow these instructions to create a Excel Add-In of your own:
 
See [How to: Create your first task pane or content app with Visual Studio] (http://msdn.microsoft.com/en-us/library/office/fp142161(v=office.15).aspx)
General work flow

1.	Office 'Add-In' requests available data feeds from OData via OData Helper on Azure
2.	OData Helper on Azure parses the metadata and send the table information to the Office 'Add-In'. UI Helper will render the data feeds.
3.	User choose a table and its columns to connect
4.	OData Helper on Azure retrieves data by OData Helper on Azure in JSON format
5.	OData Helper on Azure parses the JSON into arrays and send to Agave app. Excel Helper will write the data into Excel table.
6.	Excel Helper will read data from Excel. Diff Helper will analyze the changed and then send updated records to OData Helper on Azure
7.	OData Helper on Azure make batched OData call with JSON payload and send to OData

## Copyright ##

Copyright (c) Microsoft. All rights reserved.


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
