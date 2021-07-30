* Name......... WRITE the QUOTE LABELS from foxpro into word program
* Date......... 08/04/2003
* Caller....... writeppdoc.prg
* Notes........ This prints the Quote Labels

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

* Print the Quote Folder Labels
******
* Create Application & Document
WordApp = CreateObject("Word.application.8")
*!*	 * Open document template
*!*	WordDoc = WordApp.Documents.open('c:\my documents\ACSQtLabelspp.dot')
* Open edits the template and does not allow saving as a .doc file thus
* Add a new document using the AddressLabelspp template
*                          .Add(Template:                   , NewTemplate)
* WordDoc = WordApp.Documents.Add('c:\my documents\ACSQtLabelspp.dot', .f.)
* When the network template is used the above statement will change to
WordDoc = WordApp.Documents.Add('G:\Word Templates\Quote Templates\ACSQtLabelspp.dot', .f.)
* New caption
WordApp.Caption = "ACS Quote Labels"
* Make Visible
WordApp.Application.Visible = .t.

* WordApp.Selection.TypeParagraph = enter
WordApp.ActiveDocument.Paragraphs(1).Range.Select
WordApp.Selection.StartOf
WordApp.Selection.TypeText("Quotation #"+ cquote)
WordApp.Selection.TypeParagraph
WordApp.Selection.TypeText(ccustomer + ", " + clocation)
WordApp.Selection.TypeParagraph
WordApp.Selection.TypeText(ctitle)
WordApp.Selection.TypeParagraph
WordApp.Selection.TypeText(dtoc(date()))
* second label
WordApp.ActiveDocument.Paragraphs(6).Range.Select
WordApp.Selection.StartOf
WordApp.Selection.TypeText("Quotation #"+ cquote)
WordApp.Selection.TypeParagraph
WordApp.Selection.TypeText(ccustomer + ", " + clocation)
WordApp.Selection.TypeParagraph
WordApp.Selection.TypeText(ctitle)
WordApp.Selection.TypeParagraph
WordApp.Selection.TypeText(dtoc(date()))

*-- EOP WRITEQUOTELABELS