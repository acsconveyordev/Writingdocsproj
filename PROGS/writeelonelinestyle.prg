* Name......... WRITE an Equipment List with ONE LINE per item STYLE program
* Date......... 10/19/2010
* Caller....... quotationneweq_app.prg
* Notes........ Updated and removed Overview  10/15/2010


PUBLIC csavepath
*csavepath = "C:\writedocs\temp"

csavepath =  cserver + "\SALESSRV\DOCUMENTS\writedocs\"


* ensure that there are no tables open
close tables all

* ensure work area
select a

* this temporary table created elsewhere
* needs to be added to the program
do createeltable

* open the drawing number group table
use ctabledir + cdrawing + 'G.DBF' in a alias ngrnm

* open the group names table
use cmoldpath + 'groupnames' in b

use ctabledir + cdrawing + 'E.DBF' in c

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

* Create Application & Document
WordApp = CreateObject("Word.application.8")
* Make Visible
WordApp.Application.Visible = .t.
WordApp.Application.WindowState = wdWindowStateMaximize

select a
go top in a

do while alltrim(a.group) # '$'
	* All the names in the group name table must have a template
	store alltrim(alltrim(gname1) + " " + alltrim(gname2)) to cgrpname
	* check to see if there is a document for this group name
	if file(coverviewdir + cgrpname + '.doc') = .t.
		? 'Have the document for ' + cgrpname
		store .t. to lgotdoc
	else
		? 'Do not have a document for ' + cgrpname
		store .f. to lgotdoc
	endif

	if lgotdoc = .t.
		* get the group named template
		WordDoc = WordApp.Documents.Add(coverviewdir + cgrpname + '.doc', .f.)
	else
		* get the group new template with overview
		* WordDoc = WordApp.Documents.Add(coverviewdir + 'newpagetemplate.doc', .f.)
		* get the group new template without overview
		WordDoc = WordApp.Documents.Add(coverviewdir + 'newpagetemplatenooverview.doc', .f.)
	endif

	* fill in the lettered group
	WordApp.Selection.InsertBefore('"'+alltrim(a.group)+'" Group')
	if lgotdoc = .f.
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.InsertBefore(alltrim(alltrim(gname1) + " " + alltrim(gname2)))
	endif

	* variable to determine when a table row needs to be added
	store 2 to trcount

	* equipment information
	* do while substr(alltrim(c.eqitemref),1,1) = substr(alltrim(a.group),1,1)
	do while substr(c.eqitemref,1,2) = substr(a.group,1,2)
		MyGroupTable = WordApp.ActiveDocument.Tables(1)
		With MyGroupTable
			* skips over deleted items
			* can change items and/or groups
			*******
			store alltrim(c.eqitemref) to citemref
			*******
			do while c.eqqty < 1 .and. eof("c") = .f.
				skip 1 in c
			enddo
			if substr(c.eqitemref,1,2) = substr(a.group,1,2)
				store alltrim(c.eqitemref) to citemref
				.Cell(trcount,1).Range.InsertAfter(alltrim(c.eqitemref))
				.Cell(trcount,2).Range.InsertAfter(c.eqqty)
				if substr(c.eqdesc,1,8) = 'Existing'
					skip 1 in c
				endif
				if substr(c.eqdesc,1,7) = 'Unknown'
					skip 1 in c
				endif
				if c.eqlayer = 'C_EXST_W'
					.Cell(trcount,3).Range.InsertAfter("Equipment to be converted.")
				else
					.Cell(trcount,3).Range.InsertAfter(alltrim(alltrim(c.eqlength) + " " + alltrim(c.eqdesc) + " " + alltrim(c.eqscale)) + " " + alltrim(c.eqlayerdes))
				endif
				* + chr(13) + c.comment)
				trcount = trcount + 1
			endif
			* one line per item
			do while alltrim(c.eqitemref) = citemref
				skip 1 in c
			enddo
			* only adds a row when necessary
			if substr(c.eqitemref,1,2) = substr(a.group,1,2).and. trcount > 3
				.Rows.Add
			endif
		EndWith   && MyGroupTable
	enddo
	WordDoc.SaveAs(csavepath + alltrim(substr(a.group,1,2)) + '.doc')
	WordApp.Documents.Close()
	skip 1 in a
enddo

* table is on the $ group name record
skip 1 in a
if eof() = .f.
	do while eof() = .f.
		* all the names in the group name table must have a template
		store alltrim(alltrim(gname1) + " " + alltrim(gname2)) to cgrpname
		* ? cgrpname
		* check to see if there is a document for this group name
		if file(coverviewdir + cgrpname + '.doc') = .t.
			? 'Have the document for ' + cgrpname
			store .t. to lgotdoc
		else
			? 'Do not have a document for ' + cgrpname
			store .f. to lgotdoc
		endif
		if lgotdoc = .t.
			* get the group named template
			WordDoc = WordApp.Documents.Add(coverviewdir + cgrpname + '.doc', .f.)
		else
			* get the group new template
			* WordDoc = WordApp.Documents.Add(coverviewdir + 'newpagetemplate.doc', .f.)
			* get the group new template without overview
			WordDoc = WordApp.Documents.Add(coverviewdir + 'newpagetemplatenooverview.doc', .f.)
endif
		* fill in the option group
		WordApp.Selection.InsertBefore('OPTION "' + alltrim(a.group)+'"')
		if lgotdoc = .f.
			WordApp.selection.MoveDown(wdLine,1)
			WordApp.selection.InsertBefore(alltrim(alltrim(gname1) + " " + alltrim(gname2)))
		endif
		* variable to determine when a table row needs to be added
		store 2 to trcount
		* equipment information
		do while substr(c.eqitemref,1,2) = substr(a.group,1,2)
			MyGroupTable = WordApp.ActiveDocument.Tables(1)
			With MyGroupTable
				* do while c.eqqty < 1 .and. eof("c") = .f.
					* skip 1 in c
				* enddo
				do while c.eqqty < 0 .and. eof("c") = .f.
					skip 1 in c
				enddo
				if substr(c.eqitemref,1,2) = substr(a.group,1,2)
					store alltrim(c.eqitemref) to citemref
					.Cell(trcount,1).Range.InsertAfter(alltrim(c.eqitemref))
					.Cell(trcount,2).Range.InsertAfter(c.eqqty)
					if substr(c.eqdesc,1,8) = 'Existing'
						skip 1 in c
					endif
					if c.eqlayer = 'C_EXST_W'
						.Cell(trcount,3).Range.InsertAfter("Equipment to be converted.")
					else
						.Cell(trcount,3).Range.InsertAfter(alltrim(alltrim(c.eqlength) + " " + alltrim(c.eqdesc) + " " + alltrim(c.eqscale)) + " " + alltrim(c.eqlayerdes))
					endif
					* + chr(13) + c.comment)
					trcount = trcount + 1
				endif
				* one line per item
				do while alltrim(c.eqitemref) = citemref
					skip 1 in c
				enddo
				* only adds a row when necessary
				if substr(c.eqitemref,1,2) = substr(a.group,1,2).and. trcount > 3
					.Rows.Add
				endif
			EndWith   && MyGroupTable
		enddo
		WordDoc.SaveAs(csavepath + alltrim(substr(a.group,1,2)) + '.doc')
		WordApp.Documents.Close()
		skip 1 in a
	enddo
endif

WordDoc = WordApp.Documents.Add(coverviewdir + 'new.doc', .f.)
select a
go top
do while eof("a") = .f.
	if substr(a.group,1,2) = "01"
	    WordApp.Selection.InsertBreak(wdPageBreak)
	   	WordApp.Selection.InsertAfter("OPTIONS")
		With WordApp.Selection.Font
			* .Bold = wdToggle
			.Name = "Arial"
			.Size = 24
			.Underline = .t.
			.Italic = wdToggle
		EndWith
		WordApp.Selection.MoveRight(wdcharacter,7)
		* enter, enter
		WordApp.Selection.TypeParagraph
		WordApp.Selection.TypeParagraph
   	endif
   	if alltrim(a.group) # '$'
		if alltrim(a.group) = 'A'
   			WordApp.Selection.MoveDown(wdLine,1)
   		endif
		WordApp.Selection.InsertFile(csavepath + alltrim(substr(a.group,1,2)) + '.doc', "", .f.,.f.,.f.)
*!*			WordApp.Selection.InsertFile(coverviewdir + alltrim(substr(a.group,1,2)) + '.doc', "", .f.,.f.,.f.)
	endif
	skip 1 in a
enddo
* there is always one return, thus delete it
WordApp.Selection.Delete(wdCharacter,1)

WordDoc.SaveAs(coverviewdir + cdrawing + '.doc')

SET DEBUG ON
SET STEP ON

* clean up
select a
go top
do while eof("a") = .f.
   	if alltrim(a.group) # '$'
		erase coverviewdir + alltrim(substr(a.group,1,2)) + '.doc'
	endif
	skip 1 in a
enddo
close tables all
set message to
*erase ctabledir + cdrawing + 'E.DBF'
*clear
release csavepath

*-- EOP WRITEELONELINESTYLE