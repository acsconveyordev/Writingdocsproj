* Name......... write the PRICE PAGE for WEYerhaeuser
* Date......... 11/04/2010
* Caller....... writeppdoc.prg
* Notes........ This will be the one called for Weyerhaeuser
*               when the previous quote showed a discount.
*               Does not have the ability to do Estimated Freight

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
* at this point the nmaterialdisc and the ninstalldisc are both 0
* since the discount is determined by the total cost you must first
* determine that then fill in the table.
* keep a running total to ensure the numbers add up
* discount levels
* 6% Overall up to $200,000
* 8% Overall $200,001 to $1,000,000
* 10% Overall $1,000,001 and up
* store 0.94 to noveralldisc   && default
* determine discount
do while eof() = .f.
	if alltrim(a.group) = "$"
		do case
		case a.total > 1000000
			 store 0.90 to noveralldisc
		case a.total > 200000
			 store 0.92 to noveralldisc
		otherwise
			store 0.94 to noveralldisc
		endcase
	endif
	skip 1 in a
enddo

* return to the top before continuing
go top

* clastgroup must have a value of at least "A"
store "A" to clastgroup
* running totals required to ensure totals
store 0 to nmatlpc, nnetmatlpc
store 0 to ninstrrr, nnetinstrrr
do while alltrim(a.group) # '$'
	if tcount => 6 .and. alltrim(a.group) # '$'
		MyGroupTable.Rows.Add
		* each row adds 8 paragraphs
		pcount = pcount + 8
	endif
	With MyGroupTable
		.Cell(tcount,1).Range.InsertAfter(a.group)
		.Cell(tcount,2).Range.InsertAfter(alltrim(a.gname1) + " " + alltrim(a.gname2))
*!*			store round((a.material + a.trnprice + a.pc_tech)* noveralldisc,0) to nmatlpc
		store round((a.material + a.pc_tech)* noveralldisc,0) to nmatlpc
		do numberformat with 3,"nmatlpc"
		store nnetmatlpc + nmatlpc to nnetmatlpc
		store round((a.install + a.rtotal)* noveralldisc,0) to ninstrrr
		do numberformat with 4,"ninstrrr"
		store nnetinstrrr + ninstrrr to nnetinstrrr
		do numberformat with 7,"nmatlpc + ninstrrr"
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
	* columns 4 and 5 to be removed from table later
enddo
******

******
* does not include discount
WordApp.ActiveDocument.Paragraphs(84+pcount).Range.Select
if nnumofgroups = 1
	* table and following paragraph will be deleted later
else
	WordApp.Selection.MoveLeft(wdCharacter,1)
	WordApp.Selection.InsertAfter("Above Prices include corporate discount.")
endif
******

******
* Paragraph 85 is the "Total For Groups :" line
*  before any group lines are added which increase the value of pcount
*  go to Paragraph 85 plus pcount and insert information
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
	* the material and pc_tech to be added together and discounted together
	do numberformat with 3,"nnetmatlpc"
	tcount = tcount + 1
	if a.material # 0
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.MoveRight(wdCharacter,1)
		.Cell(tcount-1,2).Range.InsertBefore("Net ")
	endif
	* the installation and rtotal to be added together and discounted together
	do numberformat with 3,"nnetinstrrr"
	if a.install + a.rtotal # 0
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.MoveRight(wdCharacter,1)
		.Cell(tcount,2).Range.InsertBefore("Net ")
	endif
	tcount = tcount + 3
	WordApp.Selection.MoveDown(wdLine,3)
	WordApp.Selection.MoveLeft(wdCharacter,6)
	if a.material # 0
		.Cell(tcount,2).Range.InsertBefore("NET ")
	endif
	do numberformat with 3,"nnetmatlpc + nnetinstrrr"
	* at this point there are 2 rows in the above table that need to be removed
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
	* Prices quoted do not include freight or taxes.
*!*		WordApp.ActiveDocument.Paragraphs(112+pcount).Range.Select
	WordApp.ActiveDocument.Paragraphs(107+pcount).Range.Select
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
*!*	WordApp.ActiveDocument.Paragraphs(111+pcount).Range.Select
WordApp.ActiveDocument.Paragraphs(106+pcount).Range.Select
* go to front
WordApp.Selection.MoveLeft(wdCharacter,1)
* add a paragraph
WordApp.Selection.TypeParagraph
* go back to the previous paragraph
WordApp.Selection.MoveLeft(wdCharacter,1)
* insert information
WordApp.Selection.InsertAfter("Net total includes corporate discount.")
pcount = pcount + 1
*!*	WordApp.ActiveDocument.Paragraphs(112+pcount).Range.Select
WordApp.ActiveDocument.Paragraphs(107+pcount).Range.Select
WordApp.Selection.EndKey
WordApp.Selection.TypeParagraph
WordApp.Selection.InsertBefore("Payment Terms: Normal terms are per Master Business Agreement between Weyerhaeuser ")
WordApp.Selection.InsertAfter("& ACS for equipment purchased are 90% when items are received at ordering site & 10% ")
WordApp.Selection.InsertAfter("retainage not to exceed 60 days. Payment terms for parts and services purchased ")
WordApp.Selection.InsertAfter("separately from an equipment purchase shall be net thirty days.")

pcount = pcount + 1
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
	if isblank(option_ref) = .t.    && no reference
*!*			if a.material = 0 .and. a.install = 0 .and. a.rtotal = 0 .and. a.trnprice = 0 .and. a.pc_tech = 0
		if a.material = 0 .and. a.install = 0 .and. a.rtotal = 0 .and. a.pc_tech = 0
			* template information must be removed then add a paragraph
			WordApp.Selection.Movedown(wdLine,1)
			WordApp.Selection.HomeKey(wdline,1) &&    Selection.HomeKey Unit:=wdLine
			WordApp.Selection.Delete(wdWord,21)
			* all template information cleared at this point
		else
			* move cursor to the beginning of the Add Material statement
			WordApp.Selection.MoveRight(wdCharacter,4)
*!*				if a.material + a.trnprice + a.pc_tech = 0    && remove the line for material
			if a.material + a.pc_tech = 0    && remove the line for material
				* at this time the cursor is somewhere on the Add Material line thus delete
				WordApp.Selection.StartOf
				WordApp.Selection.TypeBackspace
				WordApp.Selection.TypeBackspace
				WordApp.Selection.Delete(wdCharacter,15)
				* cursor is at the beginning
			else
*!*					if a.material + a.trnprice + a.pc_tech < 0
				if a.material + a.pc_tech < 0
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Deduct Net")
				else
					WordApp.Selection.MoveRight(wdCharacter,4)
					WordApp.Selection.InsertBefore("Net ")
				endif
				WordApp.Selection.EndKey
*!*					store round((a.material + a.trnprice + a.pc_tech) * noveralldisc,0) to ndismaterialpc
				store round((a.material + a.pc_tech) * noveralldisc,0) to ndismaterialpc
				do numberformatopt with "ndismaterialpc"
				WordApp.Selection.MoveRight(wdCharacter,2)
				* cursor is at the beginning
			endif
			* at his point the cursor is at the beginning of the line
			if a.install + a.rtotal = 0    && remove the line for install
				* at this time the cursor is at the beginning of the line
				WordApp.Selection.Delete(wdCharacter,21)
				* at this time the cursor is at the beginning of the ADD TOTAL line
			else
				if a.install + a.rtotal < 0
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Deduct Net")
				else
					WordApp.Selection.MoveRight(wdCharacter,6)
					WordApp.Selection.InsertBefore("Net ")
				endif
				store round((a.install + a.rtotal) * noveralldisc,0) to ndisinstallrrr
				WordApp.Selection.EndKey
				do numberformatopt with "ndisinstallrrr"
				WordApp.Selection.MoveRight(wdCharacter,2)
				* cursor is at the beginning of the ADD TOTAL line
			endif
			* at his point the cursor is at the beginning of the ADD TOTAL line
			* a.total = a.material + a.pc_tech .or. a.total = a.install + a.rtotal
			* don't print this line
*!*				if a.total = a.material + a.trnprice + a.pc_tech .or. a.total = a.install + a.rtotal
			if a.total = a.material + a.pc_tech .or. a.total = a.install + a.rtotal
				* don't print this line
				* at this time the cursor is at the beginning of the line
				WordApp.Selection.Delete(wdCharacter,14)
				WordApp.Selection.TypeBackspace
				* at this time the cursor is at the beginning of the ADD TOTAL line
			else
			    WordApp.Selection.MoveRight(wdCharacter,5)
				if a.total < 0
					WordApp.Selection.InsertBefore("DEDUCT NET")
				else
				    if a.total # 0
						WordApp.Selection.InsertBefore("NET")
					endif
				endif
				* Deduct must be written within the paragraph
				*  thus you must go back to the start of and delete the word add
				WordApp.Selection.StartOf
				WordApp.Selection.Delete(wdCharacter,3)
				WordApp.Selection.EndKey
				do numberformatopt with "ndismaterialpc + ndisinstallrrr"
				* a.total was formatted and must be written within the paragraph
				* thus you must go back to the start of and delete the '$'
				WordApp.Selection.StartOf
			    WordApp.Selection.Delete(wdCharacter,1)
			endif
		endif   && material, install, rtotal, and pc_tech all equal zero if
	else    && reference used
		* add to group name
		WordApp.Selection.InsertAfter(" IN LIEU OF GROUP " + alltrim(option_ref))
*!*			if (a.material + a.trnprice + a.pc_tech - g.material - g.trnprice - g.pc_tech ) = 0 .and. ;
*!*			   (a.install + a.rtotal - g.install - g.rtotal) = 0
		if (a.material + a.pc_tech - g.material - g.pc_tech ) = 0 .and. ;
		   (a.install + a.rtotal - g.install - g.rtotal) = 0     && template information must be removed then add a paragraph
			WordApp.Selection.Movedown(wdLine,1)
			WordApp.Selection.HomeKey(wdline,1) &&    Selection.HomeKey Unit:=wdLine
			WordApp.Selection.Delete(wdWord,21)
			* all template information cleared at this point
		else
			* move cursor to the beginning of the Add Material statement
			WordApp.Selection.MoveRight(wdCharacter,4)
*!*				if a.material + a.trnprice + a.pc_tech - g.material - g.trnprice - g.pc_tech = 0    && remove the line for material
			if a.material + a.pc_tech - g.material - g.pc_tech = 0    && remove the line for material
				* at this time the cursor is somewhere on the Add Material line thus delete
				WordApp.Selection.StartOf
				WordApp.Selection.TypeBackspace
				WordApp.Selection.TypeBackspace
				WordApp.Selection.Delete(wdCharacter,15)
				* cursor is at the beginning
				store 0 to ndismaterialpc
			else
			    WordApp.Selection.Delete(wdCharacter,3)
				if a.material + a.pc_tech - g.material - g.pc_tech < 0
					WordApp.Selection.InsertBefore("Deduct Net")
				else
					WordApp.Selection.InsertBefore("Add Net")
				endif
*!*					store round((a.material + a.trnprice + a.pc_tech - g.material - g.trnprice - g.pc_tech) * noveralldisc,0) to ndismaterialpc
				store round((a.material + a.pc_tech - g.material - g.pc_tech) * noveralldisc,0) to ndismaterialpc
				WordApp.Selection.EndKey
				do numberformatopt with "ndismaterialpc"
				WordApp.Selection.MoveRight(wdCharacter,2)
				* cursor is at the beginning
			endif
			* at his point the cursor is at the beginning of the line
			if a.install + a.rtotal - g.install - g.rtotal = 0    && remove the line for install
				* at this time the cursor is at the beginning of the line
				WordApp.Selection.Delete(wdCharacter,21)
				* at this time the cursor is at the beginning of the ADD TOTAL line
			else
			    WordApp.Selection.MoveRight(wdCharacter,2)
			    WordApp.Selection.Delete(wdCharacter,3)
				if a.install + a.rtotal - g.install - g.rtotal < 0
					WordApp.Selection.InsertBefore("Deduct Net")
				else
					WordApp.Selection.InsertBefore("Add Net")
				endif
				store round((a.install + a.rtotal - g.install - g.rtotal) * noveralldisc,0) to ndisinstallrrr
				WordApp.Selection.EndKey
				do numberformatopt with "ndisinstallrrr"
				WordApp.Selection.MoveRight(wdCharacter,2)
				* cursor is at the beginning of the ADD TOTAL line
			endif
			* at this point the cursor is at the beginning of the ADD TOTAL line
			* a.total - g.total = a.material - g.material + a.pc_tech - g.pc_tech .or. ;
			* a.total - g.total = a.install - g.install + a.rtotal - g.rtotal
			* don't print this line
*!*				if a.total - g.total = a.material - g.material + a.trnprice - g.trnprice + a.pc_tech - g.pc_tech .or. ;
*!*				   a.total - g.total = a.install - g.install + a.rtotal - g.rtotal
			if a.total - g.total = a.material - g.material + a.pc_tech - g.pc_tech .or. ;
			   a.total - g.total = a.install - g.install + a.rtotal - g.rtotal				&& don't print this line
				* at this time the cursor is at the beginning of the line
				WordApp.Selection.Delete(wdCharacter,14)
				WordApp.Selection.TypeBackspace
				* at this time the cursor is at the beginning of the ADD TOTAL line
			else
				if a.total - g.total < 0
					WordApp.Selection.InsertBefore("DEDUCT NET")
				else
					WordApp.Selection.InsertBefore("ADD NET")
				endif
				* Deduct net and Add net must be written within the paragraph
				*  thus you must go back to the start of and delete the word add
				WordApp.Selection.StartOf
			    WordApp.Selection.Delete(wdCharacter,3)
				WordApp.Selection.EndKey
				do numberformatopt with "ndismaterialpc + ndisinstallrrr"
				* g.total was formatted and must be written within the paragraph
				* thus you must go back to the start of and delete the '$'
				WordApp.Selection.StartOf
			    WordApp.Selection.Delete(wdCharacter,1)
			endif
		endif   && (a.material - g.material + a.pc_tech - g.pc_tech ) = 0 .and. (a.install - g.install + a.rtotal - g.rtotal) = 0
	endif    && option reference
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
	if eof() = .f.
		* go to the end then add a paragraph
		* always add the group name
 	    WordApp.Selection.EndKey
		WordApp.Selection.TypeParagraph
		* Enter information
		if isblank(option_ref) = .t.    && no reference
			WordApp.Selection.TypeText("Option " + a.group + " - " + alltrim(a.gname1) + " " + alltrim(a.gname2))
		else
			* add to group name
			WordApp.Selection.TypeText("Option " + a.group + " - " + alltrim(a.gname1) + " " + alltrim(a.gname2) + " IN LIEU OF GROUP " + alltrim(option_ref))
		endif
		* go to end then add a paragraph
 	    WordApp.Selection.EndKey
		WordApp.Selection.MoveDown(wdLine,1)
		* now go back and format previous paragraph
		* paragraph 111
*!*			WordApp.ActiveDocument.Paragraphs(111+pcount).Range.Select
		WordApp.ActiveDocument.Paragraphs(106+pcount).Range.Select
		With WordApp.Selection
			With .Shading
				.Texture = wdTexture5Percent
				.ForegroundPatternColorIndex = wdAuto
				.BackgroundPatternColorIndex = wdWhite
			EndWith
			With .Font
				.Underline = wdUnderlineSingle
				.Bold = .t.
			EndWith
		EndWith
		* move down the paragraph that was added and insert information
		WordApp.Selection.MoveDown(wdLine,1)
		* determine if a reference has been used
		if isblank(option_ref) = .t.    && no reference
*!*				if a.material = 0 .and. a.install = 0 .and. a.rtotal = 0 .and. a.trnprice = 0 .and. a.pc_tech = 0
			if a.material = 0 .and. a.install = 0 .and. a.rtotal = 0 .and. a.pc_tech = 0
				* just add a line between titles
			else
				if a.material + a.pc_tech = 0
					store 0 to ndismaterial
					* don't add the line for material
				else
					* add information
*!*						if a.material + a.trnprice + a.pc_tech < 0
					if a.material + a.pc_tech < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Net Material"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Net Material"+Chr(9)+Chr(9))
					endif
*!*						store round((a.material + a.trnprice + a.pc_tech) * noveralldisc,0) to ndismaterialpc
					store round((a.material + a.pc_tech) * noveralldisc,0) to ndismaterialpc
					WordApp.Selection.EndKey
					do numberformatopt with "ndismaterialpc"
				endif
				if a.install + a. rtotal = 0    && don't add the line for installation
					* don't add the line for installation
				else
					* add information
					WordApp.Selection.EndKey
					iif(a.material # 0,WordApp.Selection.Typetext(Chr(11)),"")
					if a.install + a.rtotal < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Net Installation"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Net Installation"+Chr(9)+Chr(9))
					endif
					store round((a.install + a.rtotal) * noveralldisc,0) to ndisinstallrrr
					WordApp.Selection.EndKey
					do numberformatopt with "ndisinstallrrr"
				endif
				WordApp.Selection.EndKey
				WordApp.Selection.TypeText(Chr(11)+Chr(9)+Chr(9))
				* a.total = a.material + a.pc_tech .or. a.total = a.install + a.rtotal
				* don't print this line
*!*					if a.total = a.material + a.trnprice + a.pc_tech .or. a.total = a.install + a.rtotal
				if a.total = a.material + a.pc_tech .or. a.total = a.install + a.rtotal
					* don't print this line
					* at this time the cursor is at the beginning of the line
					WordApp.Selection.TypeBackSpace
					WordApp.Selection.TypeBackSpace
					WordApp.Selection.TypeBackSpace
					* at this time the cursor is at the beginning of the ADD TOTAL line
				else
					if a.total < 0
						WordApp.Selection.InsertAfter("DEDUCT NET TOTAL")
					else
						WordApp.Selection.InsertAfter("ADD NET TOTAL")
					endif
					* selection must be formatted to bold
					WordApp.Selection.Font.Bold = .t.
					WordApp.Selection.EndKey
					WordApp.Selection.TypeText(Chr(9)+Chr(9))
					do numberformatopt with "ndismaterialpc + ndisinstallrrr"
					* selection must be formatted to add double underline
					WordApp.Selection.Font.Bold = .t.
					WordApp.Selection.Font.Underline = wdUnderlineDouble
					WordApp.Selection.EndKey
				endif
			endif
		else
*!*				if a.material - g.material = 0 .and. a.install - g.install = 0 .and. ;
*!*				   a.rtotal - g.rtotal = 0 .and. a.trnprice - gtrnprice = 0 .and. a.pc_tech - g.pc_tech = 0
			if a.material - g.material = 0 .and. a.install - g.install = 0 .and. ;
			   a.rtotal - g.rtotal = 0 .and. a.pc_tech - g.pc_tech = 0				&& do nothing
			else
*!*					if a.material - g.material + a.trnprice - gtrnprice + a.pc_tech - g.pc_tech = 0
				if a.material - g.material + a.pc_tech - g.pc_tech = 0
					store 0 to ndismaterialpc
					* don't add the line for material
				else
					* add information
*!*						if a.material - g.material + a.trnprice - gtrnprice + a.pc_tech - g.pc_tech < 0
					if a.material - g.material + a.pc_tech - g.pc_tech < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Net Material"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Net Material"+Chr(9)+Chr(9))
					endif
					WordApp.Selection.EndKey
*!*						store round(((a.material - g.material + a.trnprice - gtrnprice + a.pc_tech - g.pc_tech) * noveralldisc),0) to ndismaterialpc
					store round(((a.material - g.material + a.pc_tech - g.pc_tech) * noveralldisc),0) to ndismaterialpc
					do numberformatopt with "ndismaterialpc"
				endif
				if a.install - g.install + a.rtotal - g.rtotal = 0    && don't add the line for installation
					* don't add the line for installation
				else
					iif(a.material # 0,WordApp.Selection.Typetext(Chr(11)),"")
					* add information
					WordApp.Selection.EndKey
					if a.install - g.install + a.rtotal - g.rtotal < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Net Installation"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Net Installation"+Chr(9)+Chr(9))
					endif
					store round((a.install - g.install + a.rtotal - g.rtotal) * noveralldisc,0) to ndisinstallrrr
					WordApp.Selection.EndKey
					do numberformatopt with "ndisinstallrrr"
				endif
				WordApp.Selection.EndKey
				WordApp.Selection.TypeText(Chr(11) + Chr(9) + Chr(9))
				* a.total - g.total = a.material - g.material + a.pc_tech - g.pc_tech .or. ;
				* a.total - g.total = a.install - g.install + a.rtotal - g.rtotal 
				* don't print this line
*!*					if a.total - g.total = a.material - g.material + a.trnprice - gtrnprice + a.pc_tech - g.pc_tech .or. ;
*!*					   a.total - g.total = a.install - g.install + a.rtotal - g.rtotal
				if a.total - g.total = a.material - g.material + a.pc_tech - g.pc_tech .or. ;
				   a.total - g.total = a.install - g.install + a.rtotal - g.rtotal				   	&& don't print this line
					* at this time the cursor is at the beginning of the line
					WordApp.Selection.TypeBackspace
					WordApp.Selection.TypeBackspace
					WordApp.Selection.TypeBackspace
					* at this time the cursor is at the beginning of the ADD TOTAL line
				else
					if a.total - g.total < 0
						WordApp.Selection.insertafter("DEDUCT NET TOTAL")
					else
						WordApp.Selection.insertafter("ADD NET TOTAL")
					endif
					* selection must be formatted to bold
					WordApp.Selection.Font.Bold = .t.
					WordApp.Selection.EndKey
					WordApp.Selection.TypeText(Chr(9)+Chr(9))
					do numberformatopt with "ndismaterialpc + ndisinstallrrr"
					* selection must be formatted to add double underline
					WordApp.Selection.Font.Bold = .t.
					WordApp.Selection.Font.Underline = wdUnderlineDouble
					WordApp.Selection.EndKey
				endif
			endif
		endif
		* do again
		* add a paragraph, then go back to the previous one
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveUp(wdLine,1)
		pcount = pcount + 2
		* the EndKey and TypeParagraph are at the beginning
	endif
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

* go back to Group table and remove the r/r/r column and the pc tech column
if nnumofgroups # 1    && if 1 the table is removed thus no need to edit it
	With MyGroupTable
		.cell(1,5).range.select
	EndWith
	WordApp.Selection.Collapse
	WordApp.Selection.SelectColumn
	WordApp.Selection.Columns.Delete
	WordApp.Selection.SelectColumn
	WordApp.Selection.Columns.Delete
	* adjust the width of the column 2
	WordApp.Selection.Tables(1).Columns(2).SetWidth(220,wdAdjustNone)
	WordApp.Selection.Collapse
endif
* go back to the Total table and remove the r/r/r row and the pc tech row
With MyTotalTable
	.Cell(3,1).Range.Select
EndWith
WordApp.Selection.SelectRow
WordApp.Selection.Rows.Delete
WordApp.Selection.SelectRow
WordApp.Selection.Rows.Delete

*-- EOP PRICEPAGEWEY