* Name......... write MOVE the Options To A New Page program
* Date......... 08/07/2006
* Caller....... pricepageavanti.prg, pricepagemaid.prg, pricepagemodnetonly.prg
* Caller....... pricepagemodshown.prg, pricepagend.prg, pricepagewey.prg, 
* Notes........ 

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

******
* there was more than one option group
if loptionnewpage = .t.
	* this selects the paragraph to add returns to force the option to their own page
	WordApp.ActiveDocument.Paragraphs(ptircount).Range.Select
	WordApp.Selection.EndKey
	WordApp.Selection.TypeParagraph
	* add 1 to ptircount
	ptircount = ptircount + 1
	* at this point you are at the top of the next page
	NewRange = WordApp.ActiveDocument.Paragraphs(ptircount).Range
	WordApp.ActiveDocument.Tables.Add(NewRange,4,2)
	* set style, font, paragraph format and border settings
	WordApp.ActiveDocument.Tables(4).Range.Select
	With WordApp.Selection
		.Style = WordApp.ActiveDocument.Styles("Body Text")
		With .Font
			.Bold = .t.
			.Italic = .f.
			.Name = "Arial"
			.Size = 11
			.Smallcaps = .t.
		EndWith
		With .ParagraphFormat
	        .SpaceBefore = 0
	        .SpaceAfter = 0
        EndWith
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
	EndWith
	* fill in table with information
	MyHeaderTable = WordApp.ActiveDocument.Tables(4)
	With MyHeaderTable
		.Cell(1,1).Range.InsertAfter("Price Proposal")
		.Cell(2,1).Range.InsertAfter(ccustomer)
		.Cell(3,1).Range.InsertAfter(clocation)
		.Cell(3,2).Range.InsertAfter(mdy(date()))
		.Cell(4,1).Range.InsertAfter("Quotation #" + cquote)
		.Cell(4,2).Range.InsertAfter("PAGE 2")
	EndWith
	* change paragraphs to right alignment
	WordApp.ActiveDocument.Paragraphs(ptircount+7).Range.Select
	With WordApp.Selection.ParagraphFormat
		.Alignment = wdAlignParagraphRight
	EndWith
	WordApp.ActiveDocument.Paragraphs(ptircount+10).Range.Select
	With WordApp.Selection.ParagraphFormat
		.Alignment = wdAlignParagraphRight
	EndWith
	* change style of paragraph
	WordApp.ActiveDocument.Paragraphs(ptircount+11).Range.Select
	With WordApp.Selection
		.Style = WordApp.ActiveDocument.Styles("Body Text")
	EndWith
	* move to another page
	WordApp.ActiveDocument.Paragraphs(ptircount).PageBreakBefore = .t.
endif
******

*-- EOP WRITEMOVEOTNP