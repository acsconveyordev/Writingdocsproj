* Name......... write the PRICE PAGE for AVANTI conveyors only from foxpro into word program
* Date......... 02/25/2004
* Caller....... writeppdoc.prg
* Notes........ This will be the one called for Avanti Conveyors only

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

*********************************
* Enter information on Price Page
*********************************
* Continue with New Page 4
* Continue with Old Page 3

******
* fill in table with group prices
MyGroupTable = WordApp.ActiveDocument.Tables(2)
tcount = 2
* paragraph count needed to find Option - Paragraph
pcount = 0
* clastgroup must have a value of at least "A"
store "A" to clastgroup
* change Material to Material Total
With MyGroupTable
	.Cell(1,3).Range.InsertAfter(" Total")
EndWith

* prices with the material discount will be shown in this table
* keep a running total for total group variable
store 0 to nnetttotal 
do while alltrim(a.group) # '$'
	if tcount => 6 .and. alltrim(a.group) # '$'
		MyGroupTable.Rows.Add
		* each row adds 8 paragraphs
		pcount = pcount + 8
	endif
	With MyGroupTable
		.Cell(tcount,1).Range.InsertAfter(a.group)
		.Cell(tcount,2).Range.InsertAfter(alltrim(a.gname1) + " " + alltrim(a.gname2))
		store round(a.material * nmaterialdisc,0) to ndismaterial
		do numberformat with 3,"a.material - ndismaterial"
		* no installation or controls thus the next 4 are not needed
		* do numberformat with 4,"a.install"
		* do numberformat with 5,"a.rtotal"
		* do numberformat with 6,"a.pc_tech"
		* do numberformat with 7,"a.total"
		* keep a running total for total group
		nnetttotal = nnetttotal + (a.material - ndismaterial)
	EndWith
	tcount = tcount + 1
	skip 1 in a
	* check next group
	if alltrim(a.group) # "$"
		store alltrim(a.group) to clastgroup
	endif
	* Get recno() for updating progress report
	store recno("a") to nrecnum
	progressreport.refresh
enddo
******

******
* Paragraph 92 is the "Total For Groups :" line
*  before any group lines are added which increase the value of pcount
*  go to Paragraph 92 plus pcount and insert information
WordApp.ActiveDocument.Paragraphs(85+pcount).Range.Select
WordApp.Selection.EndKey
WordApp.Selection.TypeBackspace
if clastgroup = "A"
	WordApp.Selection.InsertAfter("A:")
else
	WordApp.Selection.InsertAfter("A - " + clastgroup + ":")
endif
******

******
* totals changed to a table
* fill in totals
* Tables(3) and insert/edit information
tcount = 1
MyTotalTable = WordApp.ActiveDocument.Tables(3)
With MyTotalTable
	do numberformat with 3,"nnetttotal"
	tcount = tcount + 1
	* determine if discount is needed
	if a.material # 0
		WordApp.Selection.MoveDown(wdLine,1)
		* change Material to Net Material
		WordApp.Selection.MoveRight(wdCharacter,1)
		.Cell(tcount-1,2).Range.InsertBefore("Net ")
	endif
	* change Installation to Freight
	WordApp.Selection.MoveDown(wdLine,1)
	WordApp.Selection.Delete(wdCharacter,12)
	.Cell(tcount,2).Range.InsertBefore("Estimated Freight")
	* determine where the freight cost will be stored in the quote
	if fcount("a") < 17
		* table does not have the freight field yet
		set message to "Quotation table does not have the freight field yet."
		wait "Quotation table does not have the freight field yet." window at 6,10 timeout 3
		store 0 to ntotfreight
	else
		store a.freight to ntotfreight
		do numberformat with 3,"a.freight"
	endif
	tcount = tcount + 1
	WordApp.Selection.MoveDown(wdLine,1)
	WordApp.Selection.SelectRow
	WordApp.Selection.Rows.Delete
	pcount = pcount - 4
	WordApp.Selection.SelectRow
	WordApp.Selection.Rows.Delete
	pcount = pcount - 4
	.Cell(tcount,2).Range.InsertBefore("NET ")
	do numberformat with 3,"nnetttotal + ntotfreight"
EndWith
******

* Notes
* delete the first note of the template
* go to Paragraph 106+pcount
WordApp.ActiveDocument.Paragraphs(106+pcount).Range.Select
WordApp.Selection.Delete
pcount = pcount - 1
* edit the next note
* there is no installation or r's prices on these quotes
* go to Paragraph 107+pcount
WordApp.ActiveDocument.Paragraphs(107+pcount).Range.Select
WordApp.Selection.Find.ClearFormatting
With WordApp.Selection.Find
	.Text = "freight o"
	.Forward = .t.
	.Wrap = wdFindContinue
	.Format = .f.
	.MatchCase = .t.
	.MatchWholeWord = .t.
	.MatchWildcards = .f.
	.MatchSoundsLike = .f.
	.MatchAllWordForms = .f.
EndWith
WordApp.Selection.Find.Execute
* erases the template information include
WordApp.Selection.Delete
* insert revision
WordApp.Selection.InsertAfter("include installation, installation materials, controls, custom clearance o")
WordApp.Selection.EndKey
* add a paragraph
WordApp.Selection.TypeParagraph
pcount = pcount + 1
WordApp.Selection.InsertAfter("Prices quoted include 10% material reduction.")
WordApp.Selection.EndKey
* add a paragraph
WordApp.Selection.TypeParagraph
pcount = pcount + 1
WordApp.Selection.InsertAfter("Prices quoted are in U.S. Dollars.")
WordApp.Selection.EndKey
* add a paragraph
WordApp.Selection.TypeParagraph
pcount = pcount + 1
WordApp.Selection.InsertAfter("Estimated freight does not include offloading of equipment at final destination.")
* Need a variable to find the paragraph to start adding returns until the options,
*  if there is more than 1, are moved to a new page.
* at this point the pcount variable will be used to determine the correct paragraph.
ptircount = 109 + pcount - 2
******

******
* fill in option 01 information if required
skip 1 in a
loptionnewpage = .f.
if eof() = .f.
	* variable to move the options to a new page
	loptionnewpage = .t.
	* always add the group name
	* go to Paragraph 116+pcount and insert/edit information won't work paragraph will change
	WordApp.ActiveDocument.Paragraphs(109+pcount).Range.Select
	WordApp.Selection.EndKey
	WordApp.Selection.InsertAfter(alltrim(a.gname1) + " " + alltrim(a.gname2))
	* edit/write Option 01
	do writeopt01
else    && there are no options
	* this statement deletes the next three paragraphs
	* they are the template information paragraphs for option 01
	WordApp.ActiveDocument.Range(WordApp.ActiveDocument.Paragraphs(108+pcount).Range.Start,WordApp.ActiveDocument.Paragraphs(110+pcount).Range.End).Delete
endif
if recno("a") <= reccount("a")
	skip 1 in a
endif
* Get recno() for updating progress report
store recno("a") to nrecnum
progressreport.refresh
******

******
do while eof() = .f.
	* Add additional options if required
	do writeoptadd
	skip 1 in a
	* Get recno() for updating progress report
	store recno("a") to nrecnum
	progressreport.refresh
enddo
******

******
* there was more than one option group
if loptionnewpage = .t.
	do writemoveotanp
endif
******

* edit the size of the table if necessary
do case
	case nnumofgroups = 3
		* delete last row of table
		WordApp.ActiveDocument.Tables(2).Select
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.MoveDown(wdLine,5)
		WordApp.Selection.SelectRow
		WordApp.Selection.Rows.Delete
	case nnumofgroups = 2
		* delete last two rows of table
		WordApp.ActiveDocument.Tables(2).Select
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.MoveDown(wdLine,4)
		WordApp.Selection.SelectRow
		WordApp.Selection.Rows.Delete
		WordApp.Selection.SelectRow
		WordApp.Selection.Rows.Delete
	case nnumofgroups = 1    && real one
		* delete table and the following two paragraphs
		WordApp.ActiveDocument.Tables(2).Select
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.Delete(wdCharacter,20)
		WordApp.ActiveDocument.Tables(2).Select
		WordApp.ActiveDocument.Tables(2).Delete
		WordApp.Selection.Delete(wdCharacter,1)
	otherwise
		* leave as is
endcase

* go back to table and remove the last four columns
if nnumofgroups # 1    && if 1 the table is removed thus no need to edit it
	With MyGroupTable
		.Cell(1,4).Range.Select
	EndWith
	WordApp.Selection.Collapse
	WordApp.Selection.SelectColumn
	WordApp.Selection.Columns.Delete
	WordApp.Selection.SelectColumn
	WordApp.Selection.Columns.Delete
	WordApp.Selection.SelectColumn
	WordApp.Selection.Columns.Delete
	WordApp.Selection.SelectColumn
	WordApp.Selection.Columns.Delete
	WordApp.Selection.SelectColumn
	* adjust the width of the columns 3 and then 2
	WordApp.Selection.Tables(1).Columns(3).SetWidth(104,wdAdjustNone)
	WordApp.Selection.Moveleft(wdCharacter,1)
	WordApp.Selection.SelectColumn
	WordApp.Selection.Tables(1).Columns(2).SetWidth(260,wdAdjustNone)
	WordApp.Selection.Collapse
endif

*-- EOP PRICEPAGEAVANTI