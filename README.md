# OData-Excel-Integration-App-Sample
This sample is an app for Office that reads and writes OData data to documents in an Office host application, such as Excel. This article provides sample code and procedures that show you how to design an app for office. The languages used are C# and TypeScript. The tools used to setup the sample are Visual Studio and the Azure Portal.

**Table of contents:**

[Prerequisites](#Prerequisites)<BR>
[Key components of the sample](#KeyComponents)<BR>
[Deploy the sample app](#DeployTheSampleApp)<BR>

<a name="Prerequisites"></a>
#Prerequisites:

This sample requires the following:

 - Microsoft Visual Studio 2013 Update 4 or later.

 - Office Developer Tools for Visual Studio 2013 March, 2014, version or later. (This is included in Update 2 of Visual Studio 2013.)

 - An Office 365 Developer Site in an Office 365 domain that is associated with a Azure AD tenancy. See [Sign up for an Office 365 Developer Site, set up your tools and environment, and start deploying apps](http://msdn.microsoft.com/en-us/library/office/fp179924(v=office.15).aspx) or [How to: Create a Developer Site within your existing Office 365 subscription](http://msdn.microsoft.com/en-us/library/office/jj692554(v=office.15).aspx).

 - An organization account in Microsoft Azure. See [Create an organizational user account](http://www.microsoft.com/en-us/download/details.aspx?id=44944).
 
 - Basic familiarity with Azure AD. See [Getting started with Azure AD] (http://msdn.microsoft.com/en-us/library/azure/dn655157.aspx)
 
 - Basic familiarity with creating apps for Office. See [How to: Create your first task pane or content app with Visual Studio] (http://msdn.microsoft.com/en-us/library/office/fp142161(v=office.15).aspx)
 
 - Basic familiarity with OAuth 2.0 in Azure AD. See the topics under [OAuth 2.0 in Azure AD] (http://msdn.microsoft.com/en-us/library/azure/dn645545.aspx)
 
# General work flow
1.	Office 'Add-In' requests available data feeds from OData via OData Helper on Azure
2.	OData Helper on Azure parses the metadata and send the table information to the Office 'Add-In'. UI Helper will render the data feeds.
3.	User choose a table and its columns to connect
4.	OData Helper on Azure retrieves data by OData Helper on Azure in JSON format
5.	OData Helper on Azure parses the JSON into arrays and send to Agave app. Excel Helper will write the data into Excel table.
6.	Excel Helper will read data from Excel. Diff Helper will analyze the changed and then send updated records to OData Helper on Azure
7.	OData Helper on Azure make batched OData call with JSON payload and send to OData
 
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

# Modify the sample for your needs
The following procedures can help you to use your own data source.
1.	Unzip the sample and open the *.sln file in Visual Studio.
2.	Open the web.config file in the ODSampleDataWeb project and change the '<add key="ida:ODataEndpointURL" value=" " />, Set the value of your data source url.'
In our sample: 
  <add key="ida:ODataEndpointURL"  
value="http://services.odata.org/V3/(S(omlwdrfviuvthgrncrmyko1m))/OData/OData.svc/" />
<add key="ida:ODataMetadataURL"     value="http://services.odata.org/V3/(S(omlwdrfviuvthgrncrmyko1m))/OData/OData.svc/$metadata" />

 The first value is the data source location, the second one is the data format.
 
 You can use your own data by replace the values here.

<a name="DeployTheSampleApp"></a>
# Test the app in Visual Studio
Use the steps in this section to test and debug the app.

1.	Unzip the sample and open the ExcelODataInterface.sln file in Visual Studio.
2.	Click Start or press F5 in Visual Studio.
3.	The first time that you use F5, you are prompted to grant permissions to the app. Click Trust It.
4.	In Start Action, select Internet Explorer to use an Office Online client, or select Office Desktop Client, then select Start Document to specify the kind of Office document to test with.
5.	The app page will show at the Excel right pane, choose “Products”, then, click “Connect”.
6.	The data will load in the Excel sheet, we can delete, compare and save data by the buttons in our app.

## Copyright ##

Copyright (c) Microsoft. All rights reserved.
