# Scenarios
I've got an M365 SharePoint and a public facing WordPress site. I want to use SharePoint as the source of truth for most documents, even public facing documents.

## PDF Magic on SharePoint
The SharePoint tenant has a flow defined in the doc library that automatically generates a PDF for a document every time the document is modified. The PDF is generated in a ./PDF/ directory.

So if you have `Shared Documents/Magic/Document1.docx` then when you modify the document, `Shared Documents/Magic/PDF/Document1.pdf` will be created/updated.

With a little Flow trickery, you can replace the contents of Document1.pdf without recreating the file, so any sharing links to that file will be preserved. 
This is great magic and lets you share the PDF version of a document and then update that document as much as you want without
worrying about the links becoming invalid.

## PDF Magic Pitfals

### Moving the document to a new subsite
Nothing you can do here. The document ID used in the sharing link is specific to the subsite. If you move a document to a new subsite on the SharePoint, you are out of luck. All the existing sharing links will be broken.

### Moving the document within the same SharePoint subsite
If you move the original document in the same subsite (even to a new document library), the PDF will stay behind. A *new* PDF will get created in the document's new location. Great, but all the sharing links
still point to the *OLD* PDF and there's nothing you can do to get the shared links to the old PDF to point to the new pdf.  (ok, there *is*, but I haven't implemented that yet).

To avoid this, you have to carefully orchestrate things when you move documents around on your sharepoint. If you carefully move the PDF to the correct new location, and then the document,
then any updates to that document will correctly update the PDF and all the existing sharing links will continue to work.

This is tedious and easy to do wrong. This was the motivation behind `Move-SPOFiles.ps1`


