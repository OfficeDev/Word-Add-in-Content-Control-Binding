# Word-Add-in-Content-Control-Binding
The sample shows how to use JavaScript to add bindings to the content controls in the document and verify that bindings are in place before getting values from them. 

Learn how to use JavaScript in a Word 2013 task pane app to bind to content controls in a document. The sample shows the binding process, and how to retrieve the content from those bindings so you can validate their values and submit them to a back-end data store or other system. 

**Description of the sample**
The CorporateBio.docx file is set as the StartAction property of the task pane app. The document has three content controls (Name, Position, and About Me) that the user should provide values for. The following screen shot shows how the document surface when the document is first opened. 

*Figure 1. CorporateBio.docx showing the task pane app*
![CorporateBio.docx showing the task pane app](/description/CG_CorpBioWd_fig01.gif)

The sample shows the following: 

* How to use JavaScript to add bindings to the content controls in the document.
* How to verify that bindings are in place before attempting to retrieve the values from them.
* How to retrieve values from content controls and validate that the user has entered required data.


**Prerequisites**

This sample requires:

* Visual Studio 2012 (RTM).
* Office 2013 tools for Visual Studio 2012 (RTM).
* Word 2013 (RTM).

**Key components of the sample**

The sample app contains:

* The CorporateBio project, which contains:
* The CorporateBio.xml manifest file.
* The CorporateBio.docx document, which is prepopulated with three RichTextContentControl objects.

**Note**

Each object has had its Title property set, which enables it to be bound to in JavaScript.

* The CorporateBioWeb project, which contains multiple template files. However, the two files that have been developed as part of this sample solution include:
* CorporateBio.html (in the Pages folder). This contains the HTML user interface that is displayed in the task pane. It consists of a <div> with an id of validationReport, and two buttons.
* CorporateBio.js (in the Scripts folder). This script file contains code that runs when the app is loaded. This startup script attempts to add bindings to the content controls in the document. The success or failure of this operation is reported back to the CorporateBio.html page. The script file also includes the Click event handlers for the two buttons in CorporateBio.html. One of these buttons retrieves and validates the content in the content controls by accessing the bindings that were added in the startup script. The other button provides a stub procedure that simulates submitting the data from the bindings to a back-end system or process, but only if the values from the bindings have been retrieved and validated by the first button. In all cases, a suitable report is added to the CorporateBio.html page.

All other files are automatically provided by the Visual Studio project template for apps for Office, and they have not been modified in the development of this sample app.

**Configure the sample**

1. Open the CorporateBio.sln file with Visual Studio.
2. Select the **CorporateBio** project in **Solution Explorer**.
3. In the **Properties** pane, set **Start Action** to **Office Desktop Client**.
4. Set **Start Document** to **New Word Document**.


**Build the sample**

To build the sample, choose the Ctrl+Shift+B keys.

**Run and test the sample**

1. Choose the F5 key. Word will open. There will be a **Corporate Bio** group on the **Home** ribbon with an **Open** button. *DO NOT CLICK THE BUTTON YET.*
2. In Word, open the file `{project root}\C#\CorporateBio\CorporateBio.docx`.
2. On the **Home** ribbon, click the **Open** button in the **Corporate Bio** group.

The following screen shots show examples of the document at various stages of the process. Figure 2 shows a document opened with content controls successfully bound to a custom XML part.

*Figure 2. The status for each binding has been reported in the task pane.*
![CorporateBio.docx showing the task pane app](/description/CG_CorpBioWd_fig02.gif)

Figure 3 shows the task pane app UI after the Validate button has been chosen.

*Figure 3. The bindings have been retrieved.*

![CorporateBio.docx showing the task pane app](/description/CG_CorpBioWd_fig03.gif)

Figure 4 shows the task pane app UI after the Submit button has been chosen.

*Figure 4. The data has passed validation.*

![CorporateBio.docx showing the task pane app](/description/CG_CorpBioWd_fig04.gif)


**Troubleshooting**

If the app starts with a blank document instead of the one shown in Figure 1, ensure that the StartAction property of the CorporateBio project is set to CorporateBio\CorporateBio.docx and not just to Word.

**Change log**


* First release: March 15, 2013.
* Release on GitHub: August 12, 2015

**Related content**

* [Build apps for Office](http://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Bindings object (apps for Office)](http://msdn.microsoft.com/en-us/library/office/apps/fp160966.aspx)
* [Binding to regions in a document or spreadsheet](http://msdn.microsoft.com/en-us/library/office/apps/fp123511(v=office.15).aspx)

