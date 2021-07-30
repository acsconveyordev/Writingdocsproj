* Name......... GET the write the Price Page DOCument information program
* Date......... 04/22/2010
* Caller....... writingdocs_app.prg
* Notes........ Changed the writeppdoc program into two programs.
*               This program is the first one and the writeppdoc is the second one.

* Close all open tables
close tables all

* check to see if the tables exist before running the program
* variables
PUBLIC lpdbfnotfound, lgdbfnotfound, lidbfnotfound, lvdbfnotfound, lcontinuewriting
store .f. to lpdbfnotfound, lgdbfnotfound, lidbfnotfound, lvdbfnotfound
store .f. to lcontinuewriting

* P.DBF is not needed but is here because of the form
store iif(file(ctabledir + cdrawing + 'P.DBF') = .f.,.t.,.f.) to lpdbfnotfound
* G.DBF is needed
store iif(file(ctabledir + cdrawing + 'G.DBF') = .f.,.t.,.f.) to lgdbfnotfound
* I.DBF is not needed but is here because of the form
store iif(file(ctabledir + cdrawing + 'I.DBF') = .f.,.t.,.f.) to lidbfnotfound
* V.DBF is needed
store iif(file(ctabledir + cdrawing + 'V.DBF') = .f.,.t.,.f.) to lvdbfnotfound

* show form if a file does not exist
if lpdbfnotfound = .t. .or. lgdbfnotfound = .t. .or. lidbfnotfound = .t. .or. lvdbfnotfound = .t.
	if lonnetwork = .t.
		do form cformspath + 'wfilesnotfound'
		* 
	else
		do form cformspath + 'wfilesnotfoundnn'
		* 
	endif
	* lcontinuewriting set in the forms exit button (click event)
else
	store .t. to lcontinuewriting
endif

* release form variables
release lpdbfnotfound, lgdbfnotfound, lidbfnotfound, lvdbfnotfound

if lcontinuewriting = .t.
	*******
	* add a messagebox() here to ask if a fax sheet is required
	cMessageTitle = 'FAX SHEET'
	cMessageText = 'Do you want a Fax Sheet to be included?'
	nDialogType = 4 + 32 + 0
	*  4 = Yes and No buttons
	*  32 = Question mark icon
	*  0 = First button is default
	*  256 = Second button is default
	nAnswer = MESSAGEBOX(cMessageText, nDialogType, cMessageTitle)
	DO CASE
	   CASE nAnswer = 6
	      store .t. to ldofaxsheet
	   CASE nAnswer = 7
	      store .f. to ldofaxsheet
	ENDCASE
	*******
	use cdatapath + 'people' in b order num_per shared
	use cdatapath + 'mailtore' in c order num_per shared
	if upper(alias("B")) = 'PEOPLE' .and. upper(alias("C")) = 'MAILTORE'
		* Get the person information
		* Variables needed
		PUBLIC lgotpernumber, nfromlist, lnoname
		store .f. to lgotpernumber,lnoname
		store 0 to nfromlist
		* get the peoples names
		do while lgotpernumber = .f.
			* open tables if it exists
			if file(ctabledir + cdrawing + 'L.DBF') = .t.
				use ctabledir + cdrawing + 'L.DBF' in a
				if a.NUM_ACS = 0
					* erase the old one and recreate it
					use in a
					erase ctabledir + cdrawing + 'L.DBF'
					* set message
					set message to 'Creating a mailing list.'
					wait 'Creating a mailing list.' window at 5,10 timeout 2
					* create the table
					do createltable
				endif
			else
				* set message
				set message to 'Creating a mailing list.'
				wait 'Creating a mailing list.' window at 5,10 timeout 2
				* create the table
				do createltable
			endif   && if file(ctabledir + cdrawing + 'L.DBF') = .t.
			* use cdatapath + 'people'   in b order num_per shared
			* use cdatapath + 'mailtore' in c order num_per shared
			use cdatapath + 'location' in d order num_acs shared
			if upper(alias("B")) = 'PEOPLE' .and. upper(alias("D")) = 'LOCATION'
			   	select a    && L.DBF
				go top
				* lgotpernumber, nfromlist variables are set in form
				do form cformspath + 'selectaname'
				* 
				* clean up
				release selectaname
				if lnoname = .f.
					* since the structure change the information is in A
					* the record was selected in the form thus it is in the right place
					* store the name information
					PUBLIC cppsal, cppnamef, cppnamel, cpptitle
					store rtrim(a.sal) to cppsal
					store rtrim(a.name_f) to cppnamef
					store rtrim(a.name_l) to cppnamel
					store rtrim(a.title) to cpptitle
					* get mail to location
					store a.mailto to nmailto
					* get the mail to information, address
					select d    && location
					if seek(nmailto,"location")
						PUBLIC cppcustomer, cpplocation, cppaddress, cppphone, cppfax, cppzip
						store alltrim(d.company) to cppcustomer
						cpplocation = iif(empty(d.M_CITY),"",trim(d.M_CITY)) + ;
						              iif(empty(d.M_STATE),"",(", " + trim(d.M_STATE))) + ;
						              iif(trim(d.M_COUNTRY) # "USA",(", " + trim(d.M_COUNTRY)),"")
		 				store alltrim(d.m_address1) + " " + alltrim(d.m_address2) to cppaddress
						if alltrim(d.m_country) = "USA"
							store substr(phone,1,3) + "-" + substr(phone,4,3) + "-" + substr(phone,7,4) to cppphone
							store substr(fax,1,3) + "-" + substr(fax,4,3) + "-" + substr(fax,7,4) to cppfax
						else
							store alltrim(phone) to cppphone
							store alltrim(fax) to cppfax
						endif
						store alltrim(d.m_code) to cppzip
					endif
				else
					PUBLIC cppsal, cppnamef, cppnamel, cpptitle
					PUBLIC cppcustomer, cpplocation, cppaddress, cppphone, cppfax, cppzip
					store "" to cppsal, cppnamef, cppnamel, cpptitle
					store "" to cppcustomer, cpplocation, cppaddress, cppphone, cppfax, cppzip
					* else means no name was selected
					* assumes original is going to the plant
					store ncompleteacsno to nmailto
				endif

				* copies of Price Page
				* get the person information
				* variables needed
				PUBLIC lgotcopypernumber, ncopyfromlist, lcopynoname
				store .f. to lgotcopypernumber,lcopynoname
				store 0 to ncopyfromlist
				* number count
				nNumSelected = 0
				* tables already open thus
				select a
				go top   && to move from person previouly sselected
				* for the cancel button on form
				PUBLIC lstoprunning
				store .f. to lstoprunning
				* lgotcopypernumber, ncopyfromlist, lstoprunning variables are set in form
				do form cformspath + 'selectcopytoname'
				* 
				* since the people,mailtore and location tables are open
				*  and set to the correct index fill in the rest of the information
				if lgotcopypernumber = .t.
					for f = 1 to 11
						if len(tempselectarray[f,1]) > 0
							if seek(tempselectarray[f,1],"people") = .t.
								store rtrim(people.sal) to tempselectarray[f,2]
								store rtrim(people.name_f) to tempselectarray[f,3]
								store rtrim(people.name_l) to tempselectarray[f,4]
							endif
							if seek(tempselectarray[f,1],"mailtore") = .t.
								store mailtore.mailto to tempselectarray[f,5]
							endif
							if seek(tempselectarray[f,5],"location") =.t.
								store alltrim(d.company) to tempselectarray[f,6] 
								store alltrim(d.m_address1) + " " + alltrim(d.m_address2) to tempselectarray[f,7]
								ccclocation = iif(empty(d.M_CITY),"",trim(d.M_CITY)) + ;
						                      iif(empty(d.M_STATE),"",(", " + trim(d.M_STATE))) + ;
						                      iif(trim(d.M_COUNTRY) # "USA",(", " + trim(d.M_COUNTRY)),"")
						        store ccclocation + " " + alltrim(d.m_code) to tempselectarray[f,8]
							endif
						endif
					endfor
				else
					* no one selected thus do nothing
				endif
			else
				set message to "All the tables were not opened."
				wait "All the tables were not opened." window at 6,10 timeout 3
				set message to "Selection from the mailing list table cannot be done now - Please try again."
				wait "Selection from the mailing list table cannot be done now - Please try again." window at 7,10 timeout 3
			endif    && if upper(aliases are not correct)
		enddo   && do while lgotpernumber = .f.
		release lgotpernumber, nfromlist
	endif  && if upper(alias("B")) = 'PEOPLE'
	******

	* the form cformspath + 'selectcopytoname' has a cancel button on it
	* if it is pressed lstoprunning will be set to true (.t.)
	if lstoprunning = .f.
		if lonnetwork = .t.
			******
			* open tables
			* if sold at existing use it, else use the group table
			if file(ctabledir + cdrawing + 'S.DBF') = .t.
				use ctabledir + cdrawing + 'S' in a
			else
				use ctabledir + cdrawing + 'G' in a
			endif
			use ctabledir + cdrawing + 'V' in b
			use quotetrk in c order num_quote shared
			use quothist in d order num_quote shared
			use salesser in e order ssi       shared
			use saleterr in f order st        shared
			* got all the tables
			if upper(alias("C")) = 'QUOTETRK' .and. upper(alias("D")) = 'QUOTHIST' .and. ;
			   upper(alias("E")) = 'SALESSER' .and. upper(alias("F")) = 'SALETERR'
				* determine if a reference is needed
				select a
				go bottom
				* if group not equal to '$' (total) or group is a number
				if alltrim(group) # '$' .and. asc(alltrim(group)) < 58
					* since the report can not LOCATE from within the program
					* create a new table with the option reference information in it
					select a
					copy structure to ctemppath + cdrawing + 'REF'
					select g
					use ctemppath + cdrawing + 'REF' in g
					index on group to ctemppath + cdrawing + 'group'
				endif
				* find references
				select a
				go top
				PUBLIC lreffound    && reference found
				store .f. to lreffound
				locate for isblank(option_ref) = .f.
				do while found() = .t.
					store .t. to lreffound
					select g
					append blank
					replace g.GROUP with A.GROUP
					select a
					store a.OPTION_REF to cfind
					store recno("a") to nreturnto
					go top
					do while alltrim(a.GROUP) # alltrim(cfind)
						skip 1 in a
					enddo
					replace g.MATERIAL with a.MATERIAL
					replace g.INSTALL  with a.INSTALL
					replace g.RELOCATE with a.RELOCATE
					replace g.REMOVE   with a.REMOVE
					replace g.REWORK   with a.REWORK
					replace g.RTOTAL   with a.RTOTAL
					replace g.PC_TECH  with a.PC_TECH
					replace g.TRNPRICE with a.TRNPRICE
					replace g.TOTAL    with a.TOTAL
					goto nreturnto
					continue
				enddo
				* reset table in a to the top
				select a
				go top
				* set relation here
				if lreffound = .t.
					set relation to group into g
				endif
				* get the sales support and sales territory initials
				PUBLIC cssi,csti
				store "" to cssi,csti
				select a
				if seek(b.num_quote,"quotetrk") = .t.
					store quotetrk.ssi to cssi
					store quotetrk.st to csti
				else
					if seek(b.num_quote,"quothist") = .t.
						store quothist.ssi to cssi
						store quothist.st to csti
					endif
				endif
				* get the sales support first and last name
				PUBLIC cssname,csssname
				if seek(cssi,"salesser") = .t.
					store e.sal + rtrim(e.name_f) + " " + rtrim(e.name_l) to cssname
					store rtrim(e.name_f) + " " + rtrim(e.name_l) to csssname
				else
					store "no name" to cssname,csssname
				endif
				* get the sales support full name for those that want to use it
				if seek(cssi,"salesser") = .t.
					if alltrim(cssi) = "MTM"
						store e.sal + rtrim(e.name_f) + " " + substr(e.name_m,1,1) + ". " + rtrim(e.name_l) to cssname
						store rtrim(e.name_f) + " " + substr(e.name_m,1,1) + ". " +  + rtrim(e.name_l) to csssname
					endif
					if alltrim(cssi) = "JTH"
						store e.sal + substr(e.name_m,9,3) + " " + substr(e.name_l,1,4) to cssname
						store substr(e.name_m,9,3) + " " + substr(e.name_l,1,4) to csssname
					endif
				endif
				* get the sales territory first and last name
				PUBLIC cstname
				if seek(csti,"saleterr") = .t.
					store rtrim(f.name_f) + " " + rtrim(f.name_l) to cstname
					** set Mike Lucado's to Mike Lucado
					*if csti = "019"
						*store "Mike Lucado" to cstname
					*endif
					if csti = "028"
						store "Rich Morrison" to cstname
					endif
					* get the sales territory full name for those that want to use it
					if seek(csti,"saleterr") = .t.
						* Michael P. Shenigo and Terry D. Davis
						if csti = "013" .or. csti = "021"
							store rtrim(f.name_f) + " " + substr(f.name_m,1,1) + ". " + rtrim(f.name_l) to cstname
						endif
					endif
				else
					store "no name" to cstname
				endif

				* for weyerhaeuser
				* As of 01/01/2007 The discount will be shown
				* thus the default is now .t.
				lshowdiscount = .t.    && default for original quotes
				** this is a revision thus there must be a choice
				*if ncompnumber = 198 .and. len(alltrim(cquote)) > 9
					** select whether to show discount
					*do form cformspath + 'showdiscount'
					** 
					** Yes or No selected on form
					** this sets lshowdiscount to true or false
				*endif

				* Progress report form variables
				store reccount("a") + 1 to nreccount
				store recno("a") to nrecnum
				do form cformspath + "progressreport"
				* 
			else
				set message to "All the tables were not opened."
				wait "All the tables were not opened." window at 6,10 timeout 3
				set message to "The Quote Tracking or Quote History table was not opened. - Please try again."
				wait "The Quote Tracking or Quote History table was not opened. - Please try again." window at 7,10 timeout 3
			endif   && if upper(aliases = 'QUOTETRK', 'QUOTHIST', 'SALESSER', 'SALETERR'
		endif    && if lonnetwork = .t.
	else
		* cancel selected in the form cformspath + 'selectcopytoname'
		* do nothing
		MESSAGEBOX('NO Printout will be done', 32, 'CANCEL PRESSED')
	endif   && if lstoprunning = .f.

	* the form cformspath + 'selectcopytoname' has a cancel button on it
	* if it is pressed lstoprunning will be set to true (.t.)
	if lstoprunning = .f.
		if lonnetwork = .f.
			******
			* open tables
			* if sold at existing use it, else use the group table
			if file(ctabledir + cdrawing + 'S.DBF') = .t.
				use ctabledir + cdrawing + 'S' in a
			else
				use ctabledir + cdrawing + 'G' in a
			endif
			use ctabledir + cdrawing + 'V' in b
			* use quotetrk in c order num_quote shared
			* use quothist in d order num_quote shared

			* use cdatapath + 'location' in d order num_acs shared
			* already open

			* use salesser in e order ssi       shared
			use cdatapath + 'saleterr' in f order st        shared
			* got all the tables
			if upper(alias("F")) = 'SALETERR'
				* determine if a reference is needed
				select a
				go bottom
				* if group not equal to '$' (total) or group is a number
				if alltrim(group) # '$' .and. asc(alltrim(group)) < 58
					* since the report can not LOCATE from within the program
					* create a new table with the option reference information in it
					select a
					copy structure to ctemppath + cdrawing + 'REF'
					select g
					use ctemppath + cdrawing + 'REF' in g
					index on group to ctemppath + cdrawing + 'group'
				endif
				* find references
				select a
				go top
				PUBLIC lreffound    && reference found
				store .f. to lreffound
				locate for isblank(option_ref) = .f.
				do while found() = .t.
					store .t. to lreffound
					select g
					append blank
					replace g.GROUP with A.GROUP
					select a
					store a.OPTION_REF to cfind
					store recno("a") to nreturnto
					go top
					do while alltrim(a.GROUP) # alltrim(cfind)
						skip 1 in a
					enddo
					replace g.MATERIAL with a.MATERIAL
					replace g.INSTALL  with a.INSTALL
					replace g.RELOCATE with a.RELOCATE
					replace g.REMOVE   with a.REMOVE
					replace g.REWORK   with a.REWORK
					replace g.RTOTAL   with a.RTOTAL
					replace g.PC_TECH  with a.PC_TECH
					replace g.TRNPRICE with a.TRNPRICE
					replace g.TOTAL    with a.TOTAL
					goto nreturnto
					continue
				enddo
				* reset table in a to the top
				select a
				go top
				* set relation here
				if lreffound = .t.
					set relation to group into g
				endif

				select b       && v.DBF table
				set relation to num_acs into d

				* get the sales support and sales territory initials
				PUBLIC cssi,csti				
				store "" to cssi
				store d.st to csti

				* get the sales territory first and last name
				PUBLIC cstname
				if seek(csti,"saleterr") = .t.
					store rtrim(f.name_f) + " " + rtrim(f.name_l) to cstname
					* set Richard Morrison to Rich Morrison
					if csti = "028"
						store "Rich Morrison" to cstname
					endif
					* get the sales territory full name for those that want to use it
					if seek(csti,"saleterr") = .t.
						* Michael P. Shenigo and Terry D. Davis
						if csti = "013" .or. csti = "021"
							store rtrim(f.name_f) + " " + substr(f.name_m,1,1) + ". " + rtrim(f.name_l) to cstname
						endif
					endif
					* get the initials
					store alltrim(f.si) to csti
				else
					store "no name" to cstname
				endif

				? cssi
				? csti

				* no sales support name thus
				store cstname to cssname, csssname

				* for weyerhaeuser
				lshowdiscount = .f.    && default for original quotes
				* this is a revision thus there must be a choice
				if ncompnumber = 198 .and. len(alltrim(cquote)) > 9
					* select whether to show discount
					do form cformspath + 'showdiscount'
					* 
					* Yes or No selected on form
					* this sets lshowdiscount to true or false
				endif

				* Progress report form variables
				store reccount("a") + 1 to nreccount
				store recno("a") to nrecnum
				do form cformspath + "progressreport"
				* 
			else
				set message to "All the tables were not opened."
				wait "All the tables were not opened." window at 6,10 timeout 3
				set message to "The Quote Tracking or Quote History table was not opened. - Please try again."
				wait "The Quote Tracking or Quote History table was not opened. - Please try again." window at 7,10 timeout 3
			endif   && if upper(aliases = 'QUOTETRK', 'QUOTHIST', 'SALESSER', 'SALETERR'
		endif    && if lonnetwork = .f.
	else
		* cancel selected in the form cformspath + 'selectcopytoname'
		* do nothing
		MESSAGEBOX('NO Printout will be done', 32, 'CANCEL PRESSED')
	endif   && if lstoprunning = .f.

SET DEBUG ON
SET STEP ON


	* use Visual Basic language to create a WORD document
	do writeppdoc

endif   && lcontinuewriting = .t.

* clean up
close tables all
set message to

* Variables no longer needed
*!*	PUBLIC lpdbfnotfound, lgdbfnotfound, lidbfnotfound, lvdbfnotfound, lcontinuewriting
*!*	release lpdbfnotfound, lgdbfnotfound, lidbfnotfound, lvdbfnotfound
*!*	* done in program
release lcontinuewriting
*!*	PUBLIC lgotpernumber, nfromlist, lnoname
*!*	release lgotpernumber, nfromlist
* done in program
release lnoname
release cppsal, cppnamef, cppnamel, cpptitle
release cppcustomer, cpplocation, cppaddress, cppphone, cppfax, cppzip
release lgotcopypernumber, ncopyfromlist, lcopynoname
release lstoprunning
release lreffound
release cssi,csti
release cssname,csssname
release cstname
* release nmaterialdisc, ninstalldisc    && if two types of prices are run the second can't find the variable

* erase the temporary tables
erase ctemppath + cdrawing + 'REF.DBF'
erase ctemppath + cdrawing + 'group.IDX'

* release the array
release tempselectarray

*-- EOP GETPPDOCINFO

Procedure createltable
* create the table
use cmoldpath + 'moldmls' in a shared noupdate
copy structure to ctabledir+cdrawing+'L.DBF'
use ctabledir+cdrawing+'L.DBF' in a exclusive
* the A, B, C and D work areas are being used
use ctabledir+cdrawing+'V.DBF' in e
select a
* do not create table if b.num_acs = 0
if e.NUM_ACS > 0
	append from mailtore for NUM_ACS = e.NUM_ACS
	* table has all the numbers now enter the information
	go top
	set relation to num_per into people
	do while eof("a") = .f.
		replace sal with b.sal
		replace name_f with b.name_f
		replace name_m with b.name_m
		replace name_l with b.name_l
		replace title  with b.title
		skip 1 in a
	enddo
	go top
else
	* do nothing
endif
use in e
endproc