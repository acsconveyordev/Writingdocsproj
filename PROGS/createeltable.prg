* Name......... CREATE the Equipment List TABLE program
* Date......... 04/16/2010
* Caller....... writeelonelinestyle.prg
* Notes........ This program to creates a table for use by the equipment list creation program.

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
* check for 'G'
if file(ctabledir+cdrawing+'G.DBF') = .t.
	select a
	* foxpro converts the table automatically
	use ctabledir + cdrawing + 'G.DBF' in a exclusive
	* update the structure if there is only 15 fields
	*  the new file as of 9/17/97 should have 16 fields
	*  field size for gname1 & gname2 has changed from 24 to 27
	*  field size for gname1 has changed from 27 to 40
	*  field size for gname2 has changed from 27 to 49
	*  the new file as of 04/16/2010 should have 18 fields
	if fcount("a") < 18 .or. fsize('gname1') = 24 .or. fsize('gname1') = 27 
		set message to 'Updating ' + upper(cdrawing) + 'G.DBF'
		wait 'Updating ' + upper(cdrawing) + 'G.DBF' window at 6,10 timeout 2
		use in a
		rename ctabledir + cdrawing + 'G.DBF' to ctabledir + 'a' + cdrawing + 'G.DBF'
		use cmoldpath + 'moldttl' in a shared noupdate
		copy structure to ctabledir + cdrawing + 'G.DBF'
		use ctabledir + cdrawing + 'G.DBF' in a
		append from ctabledir + "a" + cdrawing + 'G.DBF'
		erase ctabledir + 'a' + cdrawing + 'G.DBF'
	endif
	use in a
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
	* fill out the appended table
	do while eof("a") = .f.
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
						append blank
					endif
					use in f
					replace d.EQDESCC with 'WITH'
					replace d.EQITEMREF with citem
					append from ctemppath + 'convert3'
					store recno("d") to nrecback   && ADDED
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
					append blank
				endif
				use in f
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
endif

* clean up
close tables all
set message to
erase ctemppath + 'item.idx'
release lcontinuewriting

*-- EOP CREATEELTABLE