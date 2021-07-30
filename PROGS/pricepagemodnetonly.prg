* Name......... write the PRICE PAGE with Material Only Discount shown as NET totals ONLY from foxpro into word program
* Date......... 05/29/2012
* Caller....... writeppdoc.prg
* Notes........ This will be the one called for material discount only.
*               Does not have the ability to do Estimated Freight

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

*********************************
* Enter information on Price Page
*********************************

******
* fill in table with group prices
MyGroupTable = WordApp.ActiveDocument.Tables(2)
tcount = 2
* paragraph count needed to find Option - Paragraph
pcount = 0
* clastgroup must have a value of at least "A"
store "A" to clastgroup

* change Material to Net Material and Total to Net Total
With MyGroupTable
	.Cell(1,3).Range.Select
	WordApp.Selection.MoveLeft(wdCharacter,1)
	WordApp.Selection.MoveRight(wdCharacter,1)
	WordApp.Selection.InsertAfter("Net ")
	WordApp.Selection.MoveLeft(wdCharacter,1)
	WordApp.Selection.TypeBackspace
	.Cell(1,7).Range.Select
	WordApp.Selection.MoveLeft(wdCharacter,1)
	WordApp.Selection.MoveRight(wdCharacter,1)
	WordApp.Selection.InsertAfter("Net ")
	WordApp.Selection.Collapse
EndWith

* variable for keeping a running total
store 0 to nmdisrunningtotal

do while alltrim(a.group) # '$'
	if tcount => 6 .and. alltrim(a.group) # '$'
		MyGroupTable.rows.add
		* each row adds 8 paragraphs
		pcount = pcount + 8
	endif
	With MyGroupTable
		.Cell(tcount,1).Range.InsertAfter(a.group)
		.Cell(tcount,2).Range.InsertAfter(alltrim(a.gname1) + " " + alltrim(a.gname2))
		* as of 7/30/03 the net material and net total will be shown for each group
		* thus a running total must be created in order for the following
		* TOTAL FOR GROUPS numbers to match
		store round(a.material - (a.material * nmaterialdisc),0) to ndismaterialg
		do numberformat with 3,"ndismaterialg"    && was do numberformat with 3,"a.material"
		do numberformat with 4,"a.install"
		do numberformat with 5,"a.rtotal"
*!*			do numberformat with 6,"a.trnprice"
		do numberformat with 6,"a.pc_tech"
*!*			do numberformat with 8,"ndismaterialg + a.install + a.rtotal + a.trnprice + a.pc_tech"    && was do numberformat with 7,"a.total"
		do numberformat with 7,"ndismaterialg + a.install + a.rtotal + a.pc_tech"
		* running total updater
		store ndismaterialg + nmdisrunningtotal to nmdisrunningtotal
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

*!*	******
*!*	* does not include discount
*!*	WordApp.ActiveDocument.Paragraphs(84+pcount).Range.Select
*!*	if nnumofgroups = 1
*!*		* table and following paragraph will be deleted later
*!*	else
*!*		WordApp.Selection.MoveLeft(wdCharacter,1)
*!*		WordApp.Selection.InsertAfter("Above Prices do not include discount.")
*!*	endif
*!*	******

******
* Paragraph 85 is the "Total For Groups :" line
*  before any group lines are added which increase the value of pcount
*  go to Paragraph 85 plus pcount and insert information
*!*	WordApp.ActiveDocument.Paragraphs(90+pcount).Range.Select
WordApp.ActiveDocument.Paragraphs(85+pcount).Range.Select
WordApp.Selection.EndKey
WordApp.Selection.TypeBackspace
if like("A", clastgroup)
	WordApp.Selection.InsertAfter("A:")
else
	WordApp.Selection.InsertAfter("A - " + clastgroup + ":")
endif
WordApp.Selection.Collapse
******

******
* totals changed to a table
* fill in totals
* Tables(3) and insert/edit information
tcount = 1

MyTotalTable = WordApp.ActiveDocument.tables(3)
With MyTotalTable
*!*		store round(a.material * nmaterialdisc,0) * -1 to ndismaterial
*!*		store a.material + ndismaterial to nnetmaterial
	do numberformat with 3,"nmdisrunningtotal"    && was do numberformat with 3,"nnetmaterial"
	tcount = tcount + 1
	* determine if discount is needed
	if a.material # 0
		WordApp.Selection.MoveDown(wdLine,1)
		* change Material to Net Material
		WordApp.Selection.MoveRight(wdCharacter,1)
		.Cell(tcount-1,2).Range.InsertBefore("Net ")
	endif
	do numberformat with 3,"a.install"
	tcount = tcount + 1
	do numberformat with 3,"a.rtotal"
	tcount = tcount + 1
*!*		do numberformat with 3,"a.trnprice"
*!*		tcount = tcount + 1
	do numberformat with 3,"a.pc_tech"
	tcount = tcount + 1
	WordApp.Selection.MoveDown(wdLine,1)
	WordApp.Selection.MoveLeft(wdCharacter,6)
	if a.material # 0
		.Cell(tcount,2).Range.InsertBefore("NET ")
	endif
*!*		do numberformat with 3,"nmdisrunningtotal + a.install + a.rtotal + a.trnprice + a.pc_tech" && was do numberformat with 3,"a.total + ndismaterial"
	do numberformat with 3,"nmdisrunningtotal + a.install + a.rtotal + a.pc_tech"
EndWith
******

* note about includes discount
* for International Paper
* or Smurfit
if ncompnumber = 1098
* .or. ncompnumber = 546
	WordApp.ActiveDocument.Paragraphs(106+pcount).Range.Select
*!*		WordApp.ActiveDocument.Paragraphs(111+pcount).Range.Select
	* go to front
	WordApp.Selection.MoveLeft(wdCharacter,1)
	* add a paragraph
	WordApp.Selection.TypeParagraph
	* go back to the previous paragraph
	WordApp.Selection.MoveLeft(wdCharacter,1)
	* insert information
	WordApp.Selection.InsertAfter("Net total includes corporate discount.")
	pcount = pcount + 1
endif
* for Rock-Tenn
if ncompnumber = 0879
* was 1089
	WordApp.ActiveDocument.Paragraphs(106+pcount).Range.Select
*!*		WordApp.ActiveDocument.Paragraphs(111+pcount).Range.Select
	* go to front
	WordApp.Selection.MoveLeft(wdcharacter,1)
	* add a paragraph
	WordApp.Selection.TypeParagraph
	* go back to the previous paragraph
	WordApp.Selection.MoveLeft(wdcharacter,1)
	* insert information
	WordApp.Selection.InsertAfter("Net total includes " + alltrim(str(nmaterialdisc*100)) + "% material discount.")
	* was WordApp.Selection.InsertAfter("Net total includes 7% material discount.")
	pcount = pcount + 1
endif

******
if a.install = 0 .and. a.rtotal = 0

	* As of 01/07/2009 the first note will be removed
	* Prices quoted are based on all groups being purchased and installed at the same time. If groups
	*  are purchased or installed separately, prices must be re-evaluated.
	* delete the first note of the template
	WordApp.ActiveDocument.Paragraphs(104+pcount).Range.Select
*!*		WordApp.ActiveDocument.Paragraphs(109+pcount).Range.Select
	WordApp.Selection.Delete
	pcount = pcount - 1

	* revise note if there is no installation and no r's prices
	* Prices quoted do not include freight or taxes.
	WordApp.ActiveDocument.Paragraphs(105+pcount).Range.Select
*!*		WordApp.ActiveDocument.Paragraphs(110+pcount).Range.Select
	WordApp.Selection.Find.ClearFormatting
	With WordApp.Selection.Find
		.Text = "include f"
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
	WordApp.Selection.InsertAfter("include installation, installation materials, f")

	* add two notes
	WordApp.ActiveDocument.Paragraphs(107+pcount).Range.Select
*!*		WordApp.ActiveDocument.Paragraphs(112+pcount).Range.Select
	WordApp.Selection.EndKey
	WordApp.Selection.TypeParagraph
	WordApp.Selection.InsertAfter('An Automated Conveyor Systems, Inc. supervisor can be provided on a time and expenses basis to assist plant personnel with installation. Please see the enclosed "Installation Services Sheet" for applicable rates and expenses.')
	WordApp.Selection.EndKey
	WordApp.Selection.TypeParagraph
	WordApp.Selection.InsertAfter('A recommended list of installation materials will be provided for your use.')

	* add 2 to paragraph count
	pcount = pcount + 2
endif
******

******
* Need a variable to find the paragraph to start adding returns until the options,
*  if there is more than 1, are moved to a new page.
* at this point the pcount variable will be used to determine the correct paragraph.
*!*	ptircount = 118 + pcount - 2
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
	* go to Paragraph 109+pcount and insert/edit information won't work paragraph will change
*!*		WordApp.ActiveDocument.Paragraphs(118+pcount).Range.Select
	WordApp.ActiveDocument.Paragraphs(109+pcount).Range.Select
	WordApp.Selection.EndKey
	WordApp.Selection.InsertAfter(alltrim(a.gname1) + " " + alltrim(a.gname2))
	* edit/write Option 01
	do writeopt01
else    && there are no options
	* this statement deletes the next three paragraphs
	* they are the template information paragraphs for option 01
*!*		WordApp.ActiveDocument.Range(WordApp.ActiveDocument.Paragraphs(117+pcount).Range.Start,WordApp.ActiveDocument.Paragraphs(119+pcount).Range.End).Delete
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
* Add additional options if required
do while eof() = .f.
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

*-- EOP PRICEPAGEMODNETONLY