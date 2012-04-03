About
============

The DocumentBarcodeSplitter is a windows service that can help you implementing the following workflow.

* Splitting a PDF-Document into multiple documents by barcodes
* Every barcode marks the start of a new document
* The splitted documents are named as the barcode-value used for the split
* The service implements a file system watcher, doing the above for every new file in a defined folder

DocumentBarcodeSplitter is using imagemagick and PDFSharp to accomplish these tasks. 

Why use this software?
For example if you like to mass-import documents into a Document Management Solution. Here is a sample workflow

* Users in the Finance department "labeling" the first page of every incoming invoice with a barcode label
* The users put one or more invoices to a scanner and scans to file
* The scanner produces on big PDF-File into a network share folder
* The DocumentBarcodeSplitter generates a single PDF for every Invoice and names the pdf like the barcode value
* The splitted PDF files are auto-imported to a DMS. The file name is used as an index field

Installation
============

There is no installer or binary package available at the moment, so you have to checkout and compile yourself.


Usage
============


