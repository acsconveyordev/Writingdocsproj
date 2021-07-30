* Name......... write the PRICE PAGE with Material And Installation Discount
*               from foxpro into word program
* Date......... 10/22/2010
* Caller....... writeppdoc.prg
* Notes........ This will be the one called for material and installation discount
*               Does not have the ability to do Estimated Freight

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

PUBLIC discr
store 0 to ndiscr
PUBLIC ndisnettotal
store 0 to ndisnettotal

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
do while alltrim(a.group) # '$'
	if tcount => 6 .and. alltrim(a.group) # '$'
		MyGroupTable.Rows.Add
		* each row adds 8 paragraphs
		pcount = pcount + 8
	endif
	With MyGroupTable
		.Cell(tcount,1).Range.InsertAfter(a.group)
		.Cell(tcount,2).Range.InsertAfter(alltrim(a.gname1) + " " + alltrim(a.gname2))
		do numberformat with 3,"a.material"
		do numberformat with 4,"a.install"
		do numberformat with 5,"a.rtotal"
*!*			do numberformat with 6,"a.trnprice"
		do numberformat with 6,"a.pc_tech"
		do numberformat with 7,"a.total"
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
* does not include discount
WordApp.ActiveDocument.Paragraphs(84+pcount).Range.Select
*!*	WordApp.ActiveDocument.Paragraphs(89+pcount).Range.Select

if nnumofgroups = 1
	* table and following paragraph will be deleted later
else
	WordApp.Selection.MoveLeft(wdCharacter,1)
	WordApp.Selection.InsertAfter("Above Prices do not include discount.")
endif
******

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
******

******
* totals changed to a table
* fill in totals
* Tables(3) and insert/edit information
tcount = 1

MyTotalTable = WordApp.ActiveDocument.Tables(3)
With MyTotalTable
	do numberformat with 3,"a.material"
	tcount = tcount + 1
	* determine if discount is needed
	if a.material # 0
		WordApp.Selection.MoveDown(wdLine,1)
		* change Material to List Material
		WordApp.Selection.MoveRight(wdCharacter,1)
		.Cell(tcount-1,2).Range.InsertBefore("List ")
		* Office/Word command not in Word 97
    	* WordApp.Selection.SelectRow
    	* WordApp.Selection.InsertRowsBelow(3)		
		WordApp.Selection.MoveDown(wdLine,1)
    	WordApp.Selection.InsertRows(3)
		* each row adds 4 paragraphs
		pcount = pcount + 12
		* the next commands places you at row 2 column 1
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.MoveRight(wdCharacter,1)
		WordApp.Selection.Font.Italic = .t.
		store str(nmaterialdisc * 100,5,2) to cdisc
		.cell(tcount,2).range.insertafter("Less " + ltrim(cdisc) + "% Material Discount")
		WordApp.Selection.EndKey
		WordApp.Selection.MoveRight(wdCharacter,1)
		WordApp.Selection.Font.Underline = wdUnderlineSingle
		store round(a.material * nmaterialdisc,0) * -1 to nmatdisc
		do numberformat with 3,"nmatdisc"
   		tcount = tcount + 1
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.MoveLeft(wdCharacter,1)
		.cell(tcount,2).range.insertafter("Net Material")
		store a.material + nmatdisc to nnetmaterial
		do numberformat with 3,"nnetmaterial"
		tcount = tcount + 2
		WordApp.Selection.MoveDown(wdLine,1)
	else
		store 0 to nmatdisc
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.MoveRight(wdCharacter,1)
	endif
	do numberformat with 3,"a.install"
	tcount = tcount + 1
	* determine if discount is needed
	if a.install # 0
		WordApp.Selection.MoveDown(wdLine,1)
		* change Installation to List Installation
		WordApp.Selection.MoveRight(wdCharacter,1)
		.Cell(tcount-1,2).Range.InsertBefore("List ")
		WordApp.Selection.MoveLeft(wdCharacter,6)
	else
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.MoveRight(wdCharacter,1)
	endif
	do numberformat with 3,"a.rtotal"
	tcount = tcount + 1	
	if a.rtotal # 0
		WordApp.Selection.MoveDown(wdLine,1)
		* change Relocation to List Relocation
		WordApp.Selection.MoveRight(wdCharacter,1)
		.Cell(tcount-1,2).Range.InsertBefore("List ")
		WordApp.Selection.MoveLeft(wdCharacter,6)
	else
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.MoveRight(wdCharacter,1)
	endif
	if a.install # 0 .and. a.rtotal # 0
		* Office/Word command not in Word 97
		* WordApp.Selection.SelectRow
    	* WordApp.Selection.InsertRowsBelow(3)
    	WordApp.Selection.MoveDown(wdLine,1)
    	WordApp.Selection.InsertRows(3)
   		* each row adds 4 paragraphs
		pcount = pcount + 12
		* the next commands places you at row 2 column 1
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.MoveRight(wdCharacter,1)
		store str(ninstalldisc * 100,5,2) to cidisc
		.Cell(tcount,2).Range.InsertAfter("Less " + ltrim(cidisc) + "% Installation, Relocation, Rework and Removal Discount")
		WordApp.Selection.MoveRight(wdCell,1)
		WordApp.Selection.Font.Underline = wdUnderlineSingle
		store round(((a.install + a.rtotal) * ninstalldisc),0) * -1 to ninstdisc
		* ? "ninstdisc is " + str(ninstdisc)
		do numberformat with 3,"ninstdisc"
		tcount = tcount + 1

		WordApp.Selection.MoveLeft(wdCell,1)
		WordApp.Selection.MoveLeft(wdCharacter,1)
		* enter net Installation, Relocation, Rework and Removal Total
		.Cell(tcount,2).Range.InsertAfter("Net Installation, Relocation, Rework and Removal")
		* added
		.Cell(tcount,3).Range.Font.Underline = wdUnderlineNone
		store ((a.install + a.rtotal) + ninstdisc) to nnetinst
		* ? "nnetinst is " + str(nnetinst)
		do numberformat with 3,"nnetinst"
		tcount = tcount + 2
		WordApp.Selection.MoveRight(wdCell,6)
	else
		store 0 to ninstdisc
	endif
*!*		do numberformat with 3,"a.trnprice"
*!*		tcount = tcount + 1
	do numberformat with 3,"a.pc_tech"
	tcount = tcount + 1
	WordApp.Selection.MoveRight(wdCell,3)
	WordApp.Selection.MoveLeft(wdCharacter,1)
	if a.material # 0
		.Cell(tcount,2).Range.InsertBefore("NET ")
	endif
*!*		Change done on 07/27/2009
*!*		do case
*!*			case nmatdisc # 0 .and. ninstdisc # 0
*!*				store a.total to nnettotal
*!*			case nmatdisc # 0
*!*				store a.total + ninstdisc to nnettotal
*!*			case ninstdisc # 0
*!*				store a.total + nmatdisc to nnettotal
*!*			otherwise
*!*				store a.total + nmatdisc + ninstdisc to nnettotal
*!*		endcase
	store a.total + nmatdisc + ninstdisc to nnettotal
	do numberformat with 3,"nnettotal"
EndWith
******

******
if a.install = 0 .and. a.rtotal = 0
	* As of 01/07/2009 the first note will be removed
	* Prices quoted are based on all groups being purchased and installed at the same time. If groups
	*  are purchased or installed separately, prices must be re-evaluated.
	* delete the first note of the template
	* go to Paragraph 106+pcount
*!*		WordApp.ActiveDocument.Paragraphs(111+pcount).Range.Select
	WordApp.ActiveDocument.Paragraphs(106+pcount).Range.Select
	WordApp.Selection.Delete
	pcount = pcount - 1

	* revise note if there is no installation and no r's prices
	* go to Paragraph 107+pcount
*!*		WordApp.ActiveDocument.Paragraphs(112+pcount).range.select
	WordApp.ActiveDocument.Paragraphs(107+pcount).range.select
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
*!*		WordApp.ActiveDocument.Paragraphs(112+pcount).Range.Select
	WordApp.ActiveDocument.Paragraphs(107+pcount).Range.Select
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

*-- EOP PRICEPAGEMAID