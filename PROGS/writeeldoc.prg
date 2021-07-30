* Name......... WRITE the Equipment List DOCument from foxpro into word program
* Date......... 10/05/2005
* Caller....... writingdocs_app.prg
* Notes........ This uses Visual Basic language to create a WORD document.

* check to see if the tables exist before running the program
* variables
PUBLIC lpdbfnotfound, lgdbfnotfound, lidbfnotfound, lvdbfnotfound, lcontinuewriting
store .f. to lpdbfnotfound, lgdbfnotfound, lidbfnotfound, lvdbfnotfound
store .f. to lcontinuewriting

if file(ctabledir + cdrawing + 'P.DBF') = .f.
	store .t. to lpdbfnotfound
else
	store .f. to lpdbfnotfound
endif
if file(ctabledir + cdrawing + 'G.DBF') = .f.
	store .t. to lgdbfnotfound
else
	store .f. to lgdbfnotfound
endif
if file(ctabledir + cdrawing + 'I.DBF') = .f.
	store .t. to lidbfnotfound
else
	store .f. to lidbfnotfound
endif
* V.DBF is not needed but is here because of the form
if file(ctabledir + cdrawing + 'V.DBF') = .f.
	store .t. to lvdbfnotfound
else
	store .f. to lvdbfnotfound
endif
* show form if a file does not exist
if lpdbfnotfound = .t. .or. lgdbfnotfound = .t. .or. lidbfnotfound = .t. .or. lvdbfnotfound = .t.
	do form cformspath + 'wfilesnotfound'
	* 
	* lcontinuewriting set in the click event of the exit button
else
	store .t. to lcontinuewriting
endif
* release form variables
release lpdbfnotfound, lgdbfnotfound, lidbfnotfound, lvdbfnotfound

if lcontinuewriting = .t.
	* open tables
	use ctabledir + cdrawing + 'P.DBF' in a exclusive   && price page
	use ctabledir + cdrawing + 'G.DBF' in b exclusive   && group information
	use ctabledir + cdrawing + 'I.DBF' in c exclusive   && item information

	* create an item index
	set safety off
	select c
	index on item to ctemppath + 'item'
	set order to item
	set safety on

	* this is required to get the item comment from 'I.DBF'
	select a
	set relation to item into c

	* set message
	set message to 'Creating An Equipment List Table.'
	wait 'Creating An Equipment List Table.' window at 5,10 timeout 2

	* new set of procedures to create equipment list information
	set procedure to eqlqflagprcd

	* erase the cdrawing + 'E.DBF' table if it exists
	if file(ctabledir+cdrawing+'E.DBF') = .t.
		erase ctabledir+cdrawing+'E.DBF'
	endif

	* create the cdrawing + 'E.DBF' table in a new work area
	select d
	use cmoldpath + 'moldeqlist' in d shared noupdate
	copy structure to ctabledir + cdrawing + 'E.DBF'
	* use table and add a record
	use ctabledir + cdrawing + 'E.DBF' in d exclusive
	append blank

	* select a and return all the tables to the top
	select a
	go top in a
	go top in b
	go top in c

	* variables that are needed for the next DO command
	PUBLIC cprevitem, cprevgroup
	store '' to cprevitem, cprevgroup
	PUBLIC cpart,nxscale,nyscale
	PUBLIC cscale,clength,cdescr,clayer,nqty,ntestflag
	PUBLIC lconversion, ncversrecnos, ncversrecnoe
	store .f. to lconversion
	store 0 to ncversrecnos, ncversrecnoe
	PUBLIC cconvgname
	store '' to cconvgname
	PUBLIC clayerdesc
	set exact on
	**************
	* use ctemppath + "ourdebug" in h
	


	* fill out the appended table
	do while eof("a") = .f.
*!*			select h
*!*			append blank
*!*			replace record with recno("a")
*!*			replace timestamp with datetime()
*!*			replace message with "next record"
		select a
		
		* fill in the item, layer and description from A (P.DBF)
		* only want to replace the item one time
		store a.ITEM to citem
		replace d.EQITEMREF	with a.ITEM
		if a.ITEM = cprevitem
			* do not replace, leave blank
		else
			replace d.EQFITEM with a.ITEM
			replace d.EQITEM with a.ITEM
		endif
		replace d.EQLAYER with a.LAYER
		replace d.EQDESC with alltrim(a.DESCR)

		* reset xscale and yscale variables
		store a.LAYER to clayer
		if substr(a.LAYER,8,1) = "W" .and. lconversion = .f.
			store .t. to lconversion
			store recno("d") to ncversrecnos     && first record of conversion
		endif  
		store alltrim(a.DESCR) to cdescr
		store a.PART to cpart
		store a.XSCALE to nxscale
		store a.YSCALE to nyscale
		store mod(a.QFLAG,100) to ntestflag

		* get the group information from B (G.DBF)
		select b

		* only want to replace the group name one time
		if substr(a.ITEM,1,2) = cprevgroup
			* do not replace leave blank
		else
			locate for alltrim(substr(a.ITEM,1,2)) = alltrim(b.GROUP)
			if found() = .t.
				replace d.GROUPNAME with alltrim(b.GROUP) + ' GROUP - ' + ;
				 alltrim(b.GNAME1) + ' ' + alltrim(b.GNAME2)
			else
				store alltrim(b.GROUP) + ' GROUP - ' + ;
				 alltrim(b.GNAME1) + ' ' + alltrim(b.GNAME2) to cconvgname
			endif
			go top
		endif

		* return to a
		select a

		* the EQQTY, EQSCALE, EQLENGTH fields
		*  must be filled in from the EQLQFLAGPRCD program
		*!*	store a.LAYER to clayer
		*!*	store a.PART to cpart
		*!*	store a.XSCALE to nxscale
		*!*	store a.YSCALE to nyscale
		*!*	store mod(a.QFLAG,100) to ntestflag
		store '' to cscale
		store '' to clength
		store 0 to nqty

		* since skip 1 is the last command
		* you must get the previous item and group information now
		store a.ITEM to cprevitem
		store substr(a.ITEM,1,2) to cprevgroup
		do case
			case ntestflag = 0
				do while a.ITEM = citem
					skip in a
				enddo
			case ntestflag = 1 .or. ntestflag = 13
				do elqflag1
			case ntestflag = 2 .or. ntestflag = 10 .or. ntestflag = 11
				do elqflag2
			case ntestflag = 3
				do elqflag3
			case ntestflag = 4
				do elqflag4
			case ntestflag = 5
				do elqflag5
			case ntestflag = 6
				do elqflag6
			case ntestflag = 7 .or. ntestflag = 12
				do elqflag7
			case ntestflag = 9
				do elqflag9
		endcase
		if eof() = .t.
			skip -1 in a
			replace d.EQQTY with nqty
			replace d.EQSCALE with cscale
			replace d.EQLENGTH with clength
			* the EQLAYERDES field must be filled in
			store '' to clayerdesc
			do case
			* describe the layers
			case clayer = 'C_RELO__' .or. clayer = 'C_RELO_C'
				store '(Relocated)' to clayerdesc
			case clayer = 'C_REMV__' .or. clayer = 'C_REMV_C'
				store '(Removed)' to clayerdesc
			case clayer = 'C_REWK__' .or. clayer = 'C_REWK_C'
				store '(Reworked)' to clayerdesc
			case clayer = 'C_RUSE__' .or. clayer = 'C_RUSE_C'
				store '(Reused)' to clayerdesc
			case clayer = 'C_PEIO__' .or. clayer = 'C_PEIO_C'
				store '(Install only)' to clayerdesc
			case lconversion = .t. .and. (clayer = 'C_PLAN__' .or. clayer = 'C_PLAN_C')
				store '(New)' to clayerdesc
			endcase
			replace d.EQLAYERDES with clayerdesc
			store '' to clayerdesc
			skip 1 in a
		else
			replace d.EQQTY with nqty
			replace d.EQSCALE with cscale
			replace d.EQLENGTH with clength
		endif
		* reset variables
		store 0 to nqty
		store '' to cscale,clength

		* the EQLAYERDES field must be filled in
		store '' to clayerdesc
		do case
			* describe the layers
			case clayer = 'C_RELO__' .or. clayer = 'C_RELO_C'
				store '(Relocated)' to clayerdesc
			case clayer = 'C_REMV__' .or. clayer = 'C_REMV_C'
				store '(Removed)' to clayerdesc
			case clayer = 'C_REWK__' .or. clayer = 'C_REWK_C'
				store '(Reworked)' to clayerdesc
			case clayer = 'C_RUSE__' .or. clayer = 'C_RUSE_C'
				store '(Reused)' to clayerdesc
			case clayer = 'C_PEIO__' .or. clayer = 'C_PEIO_C'
				store '(Install only)' to clayerdesc
			case (clayer = 'C_PLAN__' .or. clayer = 'C_PLAN_C') .and. lconversion = .t.
				store '(New)' to clayerdesc
		endcase
		replace d.EQLAYERDES with clayerdesc
		store '' to clayerdesc

		* things that will need to be done before returning to DO command
		select a

		do while a.QFLAG = 8
			skip 1 in a
		enddo
		if recno("a") <= reccount("a")
			* check the item to determine if it is the last one
			if a.item # cprevitem
				if lconversion = .t.
					store recno("d") to ncversrecnoe     && last record of conversion
					select d
					copy to ctemppath + cdrawing + 'E' for recno() < ncversrecnos
					copy to ctemppath + 'convert1' for ;
					 recno() => ncversrecnos .and. eqlayer = 'C_EXST_W'
					copy to ctemppath + 'convert2' for ;
					 recno() => ncversrecnos .and. eqlayer = 'C_PLAN_W'
					copy to ctemppath + 'convert3' for ;
					 recno() => ncversrecnos .and. eqlayer # 'C_EXST_W' .and. eqlayer # 'C_PLAN_W'
					use ctemppath + cdrawing + 'E' in d
					erase ctabledir + cdrawing + 'E.DBF'
					copy to ctabledir + cdrawing + 'E'
					use ctabledir + cdrawing + 'E' in d exclusive
					erase ctemppath + cdrawing + 'E.DBF'
					* fix the remaining records
					append blank
					* replace d.GROUPNAME with cconvgname
					replace d.EQFITEM with citem
					replace d.EQITEMREF with citem
					replace d.EQDESCC with 'EXISTING'
					store recno("d") to nrecback
					append from ctemppath + 'convert1'
					go nrecback
					skip 1
					blank field d.EQFITEM
					* check for group name
					if isblank(d.GROUPNAME) = .f.
						store d.GROUPNAME to cconvgname
						blank field d.GROUPNAME
						skip -1
						replace d.GROUPNAME with cconvgname
					endif
					append blank
					* added to remove 'CONVERTED TO' from printing if there are no records
					use ctemppath + 'convert2' in f
					if reccount('f') > 0
						replace d.EQDESCC with 'CONVERTED TO'
						replace d.EQITEMREF with citem
						append from ctemppath + 'convert2'
					endif
					use in f
					append blank
					replace d.EQDESCC with 'WITH'
					replace d.EQITEMREF with citem
					append from ctemppath + 'convert3'
					erase ctemppath + 'convert*.*'
					store .f. to lconversion
					store 0 to ncversrecnos, ncversrecnoe
				endif
				* get the comment from the 'C' work area
				select c
				locate for alltrim(cprevitem) = alltrim(c.ITEM)
				if found() = .t.
					if memlines(c.comment) > 0
						replace d.EQCOMMENT with 'YES'
						replace d.EQITEM with citem
					else
						replace d.EQCOMMENT with 'NO'
					endif
				endif
				go top
			endif
			select d
			append blank
			select a
		else
			* eof() = .t. .and. lconversion = .t.
			* if eof() = .t.
			if lconversion = .t.
				store recno("d") to ncversrecnoe     && last record of conversion
				select d
				copy to ctemppath + cdrawing + 'E' for recno() < ncversrecnos
				copy to ctemppath + 'convert1' for ;
				 recno() => ncversrecnos .and. eqlayer = 'C_EXST_W'
				copy to ctemppath + 'convert2' for ;
				 recno() => ncversrecnos .and. eqlayer = 'C_PLAN_W'
				copy to ctemppath + 'convert3' for ;
				 recno() => ncversrecnos .and. eqlayer # 'C_EXST_W' .and. eqlayer # 'C_PLAN_W'
				use ctemppath + cdrawing + 'E' in d
				erase ctabledir + cdrawing + 'E.DBF'
				copy to ctabledir + cdrawing + 'E'
				use ctabledir + cdrawing + 'E' in d exclusive
				erase ctemppath + cdrawing + 'E.DBF'
				* fix the remaining records
				append blank
				* replace d.GROUPNAME with cconvgname
				replace d.EQFITEM with citem
				replace d.EQITEMREF with citem
				replace d.EQDESCC with 'EXISTING'
				store recno("d") to nrecback
				append from ctemppath + 'convert1'
				go nrecback
				skip 1
				blank field d.EQFITEM
				* check for group name
				if isblank(d.GROUPNAME) = .f.
					store d.GROUPNAME to cconvgname
					blank field d.GROUPNAME
					skip -1
					replace d.GROUPNAME with cconvgname
				endif
				append blank
				* added to remove 'CONVERTED TO' from printing if there are no records
				use ctemppath + 'convert2' in f
				if reccount('f') > 0
					replace d.EQDESCC with 'CONVERTED TO'
					replace d.EQITEMREF with citem
					append from ctemppath + 'convert2'
				endif
				use in f
				append blank
				replace d.EQDESCC with 'WITH'
				replace d.EQITEMREF with citem
				append from ctemppath + 'convert3'
				erase ctemppath + 'convert*.*'
				store .f. to lconversion
				store 0 to ncversrecnos, ncversrecnoe
			endif
			* endif
			* get the comment from the 'C' work area
			select c
			locate for alltrim(cprevitem) = alltrim(c.ITEM)
			if found() = .t.
				if memlines(c.comment) > 0
					replace d.EQCOMMENT with 'YES'
					replace d.EQITEM with citem
				else
					replace d.EQCOMMENT with 'NO'
				endif
			endif
			go top
		endif
	enddo    && of the fill out the table
	set exact off

	* clean up
	set message to

	* close new set of procedures
	close procedure eqlqflagprcd
	* release variables
	release cprevitem, cprevgroup
	release cpart,nxscale,nyscale
	release cscale,clength,cdescr,clayer,nqty,ntestflag
	release lconversion, ncversrecnos, ncversrecnoe
	release cconvgname
	release clayerdesc

	* new variable for page title
	PUBLIC cpagetitle
	store "EQUIPMENT LIST" to cpagetitle

	* create an item index
	set safety off
	select c
	index on item to ctemppath + 'item'
	set order to item
	set safety on

	* this is required to get the item comment from 'I.DBF'
	select d
	set relation to eqitem into c

	* set message
	set message to 'Writing the Equipment List in Word.'
	wait 'Writing the Equipment List in Word.' window at 5,10 timeout 2	

	* Create Application & Document
	* Dimensioning of variables not required
	*!*	Dim WordApp As Object
	WordApp = CreateObject("Word.application")
	* Add to it
	WordDoc = WordApp.Documents.Add

	* Define caption
	WordApp.Caption = "Quotation Equipment List"
	* Make Visible moved to the end
	*!*	WordApp.Application.Visible = .t.

	* Setup Page Margins 72 points = 1"
	WordApp.Selection.Pagesetup.LeftMargin = 36    && 1/2"
	WordApp.Selection.Pagesetup.RightMargin = 36
	WordApp.Selection.Pagesetup.TopMargin = 18     && 1/4"
	WordApp.Selection.Pagesetup.BottomMargin = 18
	WordApp.ActiveDocument.Paragraphs(1).Alignment = 1

	* Sets header distance from top of page
	WordApp.ActiveDocument.PageSetup.HeaderDistance = 18     && 1/4"
	* Sets footer distance from bottom of page
	WordApp.ActiveDocument.PageSetup.FooterDistance = 18     && 1/4"

	* Set up header and footer
	With WordApp.ActiveDocument.Sections(1)
		* Header
		.Headers(1).Range.Font.Bold = .t.
		.Headers(1).Range.Font.Name = "Arial"
		.Headers(1).Range.Font.Size = 10
		* Alignment = left
		.Headers(1).Range.Paragraphs(1).Alignment = 0
	    .Headers(1).Range.Text = "  ITEM       QTY     DESCRIPTION"
		* change the constant wdundelinewords to underline only the words
		wdunderlinewords = 2
		.Headers(1).Range.Font.Underline = wdUnderlinewords	
		* Footer
		.Footers(1).Range.Font.Name = "Verdana"
		.Footers(1).Range.Font.Size = 9
		* Alignment = center
		.Footers(1).Range.Paragraphs(1).Alignment = 1
		.Footers(1).Range.Text = "QUOTATION # " + cquote
	EndWith

	* Erase header and footer from first page
	With WordApp.ActiveDocument
	    .PageSetup.DifferentFirstPageHeaderFooter = .t.
	    .Sections(1).Headers(1).Range.InsertBefore("")
	EndWith

	* Set up the object MyRange
	MyRange = WordApp.ActiveDocument.Paragraphs(1).Range
	Wordapp.ActiveDocument.Paragraphs(1).SpaceAfter = 0

	* Set the Font Attributes for MyRange
	With MyRange.Font
		.Bold = .t.
		.Name = "Arial"    && .Name = "Verdana"
		.Size = 12         && .Size = 11
	EndWith

	* Insert Equipment List Heading Information   *** NOTE: chr(13) is a carraige return ***
	*!*	WordApp.Selection.InsertAfter ("EQUIPMENT LIST" + chr(13))
	WordApp.Selection.InsertBefore (cpagetitle + chr(13))
	*!*	WordApp.Selection.InsertAfter ("COMPANY NAME" + chr(13))
	WordApp.Selection.InsertAfter (upper(ccustomer) + chr(13))
	*!*	WordApp.Selection.InsertAfter ("LOCATION NAME" + chr(13))
	WordApp.Selection.Insertafter (upper(clocation) + chr(13))
	*!*	WordApp.Selection.InsertAfter ("QUOTATION # " + chr(13))
	WordApp.Selection.InsertAfter ("QUOTATION # " + cquote + chr(13))
	*!*	WordApp.Selection.InsertAfter ("LAYOUT #" + chr(13) + chr(13))
	WordApp.Selection.InsertAfter ("LAYOUT #" + clayout + chr(13) + chr(13))

	* Release selection
	With WordApp.Selection
		.Collapse
		.InsertParagraph
	EndWith

	* Move the selection to the end
	WordApp.ActiveDocument.Paragraphs(WordApp.ActiveDocument.Paragraphs.Count).Range.Select
	WordApp.Selection.MoveEnd

	* Insert Table to contain Item, Qty, Description
	MyTable = WordApp.ActiveDocument.Tables.Add(WordApp.Selection.Paragraphs(1).Range,1,3)

	* Set column 3 alignment to left, leave others to center
	WordApp.ActiveDocument.Paragraphs(10).Alignment = 0
	* Set MyRange to the new table row 1
	MyRange = WordApp.ActiveDocument.Tables(1).Rows(1).Range
	* Change Font in first column (i.e. MyRange)
	With MyRange.Font
		.Bold = .t.
		.Name = "Arial"    && .Name = "Verdana"
		.Size = 10
		.Underline = .t.
	EndWith

	* Set Table's Column Widths and First Page Heading
	With MyTable
		.Borders.Enable = .f.
		.Columns(1).Width = 44
		.Columns(2).Width = 40
		.Columns(3).Width = 472
		* insert information
		.Cell(1,1).Range.InsertAfter("ITEM")
		.Cell(1,2).Range.InsertAfter("QTY")
		.Cell(1,3).Range.InsertAfter("DESCRIPTION")
	EndWith

	* Prepare table for inserting data
	select d
	go top

	* screen updating
	*!*		WordApp.Application.ScreenUpdating = .f.

	* variable for determining if a conversion is to be done
	store .f. to lconversion

	* add a row to MyTable before you start
	MyTable.Rows.Add

	* Insert Data into Cells, adding a new row after each entry
	With MyTable
	*!*		This example determines whether the rows in the current table can be split across pages.
	*!*     If the insertion point isn't in a table, a message box is displayed.
	*!*		Selection.Collapse Direction:=wdCollapseStart
	*!*		If Selection.Tables.Count = 0 Then
	*!*		    MsgBox "The insertion point is not in a table."
	*!*		Else
	*!*		    allowBreak = Selection.Rows.AllowBreakAcrossPages
	*!*		End If
	    
		* Progress report form variables
		store reccount("d") to nreccount
		store recno("d") to nrecnum

		do form cformspath + "progressreport"
		* 

		* do while not end of file
		do while eof() = .f.
			* select the previous row
			.Rows(.Rows.Count-1).Select
			WordApp.Selection.MoveEnd (1,-1)
			* Do not let comments cross a page
			WordApp.Selection.Rows.AllowBreakAcrossPages = .f.
			* group name when required
			if len(alltrim(d.groupname)) > 0
				.Rows.Add
				.cell(.Rows.Count,3).Range.InsertAfter(alltrim(d.groupname))
				.Rows(.Rows.Count).Range.Bold = .t.
				.Rows(.Rows.Count).Range.Underline = .t.
				.Rows.Add
			endif   && end of group name selection

			* print the first item number when required
			if len(alltrim(d.eqfitem)) > 0
				* reset font of the item number
				* reset font to item information
				.Rows(.Rows.count).Range.Bold =.f.
				.Rows(.Rows.count).Range.Font.Name = "Verdana"
				.Rows(.Rows.count).Range.Font.Size = 10
				.Rows(.Rows.count).Range.Underline = .f.
				.Rows.Add
				.Cell(.Rows.Count,1).Range.InsertAfter(alltrim(d.eqitemref))
			endif
			* quantity must be printed on the first item and if not first item
			* conversion must be dealt with on the first item also
			* since a conversion item creates a merge for cells 2 and 3
			* it must be determined before proceeding	
			if len(alltrim(d.eqdescc)) > 0
				.Cell(.Rows.Count,2).Range.InsertAfter(alltrim(d.eqdescc))
				* a row must be added to keep the format
				.Rows.Add
				* then go back to the previous row and merge cells
				* select the previous row
				.Rows(.Rows.Count-1).Select
				WordApp.Selection.MoveEnd (1,-1)
				.Cell(.Rows.Count-1,2).Merge (.Cell(.Rows.Count-1,3))
			else
				.Cell(.Rows.Count,2).Range.InsertAfter(d.eqqty)
				* Comment required if a.eqcomment equals 'YES" and the quantity is not equal to zero(Deleted)
				if alltrim(d.eqcomment) = "YES" .and. d.eqqty # 0
					.Cell(.Rows.Count,3).Range.InsertAfter(alltrim(alltrim(d.eqlength) + " " + alltrim(d.eqdesc) + " " + alltrim(d.eqscale)) + " " + alltrim(d.eqlayerdes) + chr(13) + c.comment)
				else
					.Cell(.Rows.Count,3).Range.InsertAfter(alltrim(alltrim(d.eqlength) + " " + alltrim(d.eqdesc) + " " + alltrim(d.eqscale)) + " " + alltrim(d.eqlayerdes))
				endif
				.Rows.Add
			endif

			skip 1
			* If for some reason the next record is blank
			if isblank(d.eqitemref) = .t. .and. eof() = .f.
				skip 1
			endif
			* Next item so reset conversion
			if len(alltrim(d.eqfitem)) > 0
				store .f. to lconversion
			endif

			* Get recno() for updating progress report
			store recno("d") to nrecnum
			progressreport.refresh

		enddo
		progressreport.release
	EndWith

	* Make visible
	WordApp.Application.Visible = .t.

	* Save document
	* document directory
	cdocdir = cserver + "\SALESSRV\WORK IN PROGRESS\DOCUMENTS\WORD\"
	* ctabledir = cserver + "\SALESSRV\WORK IN PROGRESS\DRAWINGS\QUOTE\"
	* save in tabledir for now
	* WordDoc.SaveAs(ctabledir + cdrawing +'EL')

	* Quit Application, Set to nil
	* WordApp.Quit
	* WordApp = "nil"
endif   && if lcontinuewriting = .t.

* clean up
close tables all
set message to
erase ctabledir + cdrawing + 'E.DBF'
erase ctemppath + 'item.idx'
release cpagetitle
release lcontinuewriting

*-- EOP WRITEELDOC