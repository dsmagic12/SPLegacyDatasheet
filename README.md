# SPLegacyDatasheet
A JavaScript library enabling the use of the SharePoint 2007/2010 Datasheet view that was based on MS Access

# Requirements
The end user must have Microsoft's Access 2010 database engine (32-bit) installed (https://www.microsoft.com/en-us/download/details.aspx?id=13255) to support the ListNet / STSList.dll ActiveX component (https://msdn.microsoft.com/en-us/library/ms416795(v=office.14).aspx)
The end user must use Microsoft Internet Explorer (32-bit) as their browser when trying to view legacy datasheets https://www.microsoft.com/en-us/download/internet-explorer.aspx

# Explicit (hard code your list and view GUIDs)
Download a copy of the "spLDS.js" file and open it in a text editor like Notepad
Set spLDS.listGUID to the GUID of your list
Set spLDS.viewGUID to the GUID of your view
Save your changes to the file and close it
Upload your modified "spLDS.js" file to your site's Style Library
Add a new page to your site
Edit the page, then add a Content Editor Web Part (CEWP) to the page
Edit the Content Editor Web Part, and set its Link to point it to the "wp_spLDS_replaceQuickEditView.html" web part file you uploaded to your Style Library a moment ago
Optionally, set the Chrome Type property (under 'Appearance') of the Content Editor Web Part to be "None"
Click OK to save your changes to the Content Editor Web Part
Stop editing the page to save your changes
Once the page loads, you should see that the list view you specified in the code file is replaced with a legacy datasheet

A working example is available here: http://1.dsmagicsp.cloudappsportal.com/SitePages/demo_datasheetView.aspx


# Simple (add web part to page containing Quick Edit views)
Upload both "spLDS_replaceQuickEditView.js" and "wp_spLDS_replaceQuickEditView.html" to your site's Style Library
Go to a page on your site that contains one or more Quick Edit (SharePoint 2013 datasheet) views
Edit the page, then add a Content Editor Web Part (CEWP) to the page
Edit the Content Editor Web Part, and set its Link to point it to the "wp_spLDS_replaceQuickEditView.html" web part file you uploaded to your Style Library a moment ago
Optionally, set the Chrome Type property (under 'Appearance') of the Content Editor Web Part to be "None"
Click OK to save your changes to the Content Editor Web Part
Stop editing the page to save your changes
Once the page loads, you should see messages in your browser's script console about the code working to replace each Quick Edit view on the page with a legacy datasheet

NOTE!!! The old datasheet view's performance wasn't great. If you add this to a page containing more than 1 or 2 Quick Edit views, you should expect that the page performance and initial load time will be less than ideal.
