* Name......... SELECT DraWinG NUMBER program
* Date......... 12/23/2008
* Caller....... writingdocs_app.prg
* Notes........ Lets the operator selects a drawing number.
*               Then sets all the variables.
*               04/15/03 Added getting the discount information from the v,dbf table
*               04/21/03 Added getting the variable to determine general notes 20 and 21
*               Added the variable linmex to check for Mexico quotes Juan or Eddie

* ensure a clean work area
close tables all

* variables for form and/or program
* five variables for the four possible directorys
PUBLIC lgotdwgnfq      && cserver + \salessrv\work in progress\drawings\quote
* has not been put in its folder yet
PUBLIC lgotdwgnfbd     && cserver + \salessrv\backup\drawings\
* above 10000 in its folder  10000\10000, 10100, 10200, etc..
PUBLIC lgotdwgnfbdd    && cserver + \salessrv\backup\drawings\" + ;
*                alltrim(str(int(val(cgetdrawingnum)/1000))) + "000\ + ;
*                alltrim(str(int(val(cgetdrawingnum)/100))) + "00\"
PUBLIC lgotdwgnfbdd0   && cserver + \salessrv\backup\drawings\" + "0" + ;
*                alltrim(str(int(val(cgetdrawingnum)/1000))) + "000\0" + ;
*                alltrim(str(int(val(cgetdrawingnum)/100))) + "00\"
PUBLIC lgotdwgnfcqq    && c:\quotedata\quotes
* set all the variables to false
store .f. to lgotdwgnfq, lgotdwgnfbd, lgotdwgnfbdd, lgotdwgnfbdd0, lgotdwgnfcqq
* drawing number variable
* PUBLIC cgetdrawingnum
store "" to cgetdrawingnum
* cancel pressed in form
PUBLIC lcancelit
store .f. to lcancelit

* open form to get drawing number from operator
if lonnetwork = .t.
	do form cformspath + 'drawingnumgetter'
else
	do form cformspath + 'drawingnumgetternn'
endif
*!*	? "lgotdwgnfq is "
*!* ? lgotdwgnfq
*!*	? "lgotdwgnfbd is "
*!*	? lgotdwgnfbd
*!*	? "lgotdwgnfbdd is "
*!*	? lgotdwgnfbdd
*!*	? "lgotdwgnfbdd0 is "
*!*	? lgotdwgnfbdd0
*!*	? "lgotdwgnfcqq is "
*!*	? lgotdwgnfcqq

* information about the number after the /n
*!*	Value Application attributes 
*!*	1 Active and normal size 
*!*	2 Active and minimized 
*!*	3 Active and maximized 
*!*	4 Inactive and normal size 
*!*	7 Inactive and minimized 
* if cancel button was not pressed

STORE .F. TO LTESTIT 
if lcancelit = .f.
	* at this point there are five choices
	* 1 file in the quote directory
	* 2 file in the backup\drawings directory
	* 3 file in the backup\drawings\int(drawing number/100)*100 directory
	* 4 file in the backup\drawings\"0"+ int(drawing number/100)*100 directory
	* 5 file in the c:\quotedata\quotes directory  **** no network ****
	do case
	case lgotdwgnfq = .t.
		* use the files from this directory
	case lgotdwgnfbd = .t.
		* unzip the files to the quote directory
		* note erase later
		store cserver + "\salessrv\backup\drawings\"+alltrim(cgetdrawingnum)+".zip" to ctempvar
		* run /n7 "c:\program files\winzip\wzunzip" -n -o &ctempvar cserver + "\salessrv\work in progress\drawings\quote" *.dbf
		* run /n7 "c:\program files\winzip\wzunzip" -n -o &ctempvar cserver + "\salessrv\work in progress\drawings\quote" *.fpt
		run /n7 unzipme.bat &ctempvar

	case lgotdwgnfbdd = .t.
		* unzip the files to the quote directory
		* note erase later
		* number over 9999
		store cserver + "\salessrv\backup\drawings\" + ;
				alltrim(str(int(val(cgetdrawingnum)/1000))) + "000\" + ;
				alltrim(str(int(val(cgetdrawingnum)/100))) + "00\" + ;
				alltrim(cgetdrawingnum) + ".zip" to ctempvar
		* run /n7 "c:\program files\winzip\wzunzip" -n -o &ctempvar cserver + "\salessrv\work in progress\drawings\quote" *.dbf
		* run /n7 "c:\program files\winzip\wzunzip" -n -o &ctempvar cserver + "\salessrv\work in progress\drawings\quote" *.fpt
		run /n7 unzipme.bat &ctempvar

	case lgotdwgnfbdd0
		* number less than 10000
		store cserver + "\salessrv\backup\drawings\" + "0" + ;
				alltrim(str(int(val(cgetdrawingnum)/1000))) + "000\0" + ;
				alltrim(str(int(val(cgetdrawingnum)/100))) + "00\" + ;
				alltrim(cgetdrawingnum) + ".zip" to ctempvar
		* run /n7 "c:\program files\winzip\wzunzip" -n -o &ctempvar cserver + "\salessrv\work in progress\drawings\quote" *.dbf
		* run /n7 "c:\program files\winzip\wzunzip" -n -o &ctempvar cserver + "\salessrv\work in progress\drawings\quote" *.fpt
		run /n7 unzipme.bat &ctempvar

	case lgotdwgnfcqq
		* no network use files from c:

	endcase
	* at this point the files are in the directory
	* use tables and create/define the needed variables
	* like the openfile program in the quotation application
	* define the cdrawing variable
	PUBLIC cdrawing
	store alltrim(cgetdrawingnum) to cdrawing
	* define the clayout variable
	if len(alltrim(cdrawing)) > 5
		store substr(cdrawing,1,4) + '-' + upper(substr(cdrawing,5,1)) ;
		 + '-' + upper(substr(cdrawing,6,2)) to clayout
	else
		store substr(cdrawing,1,4) + '-' + upper(substr(cdrawing,5,1)) to clayout
	endif
	* added for drawing numbers above 9,999
	if val(cdrawing) > 9999
		if len(cdrawing) > 6 
			store substr(cdrawing,1,5) + '-' + upper(substr(cdrawing,6,1)) ;
			 + '-' + upper(substr(cdrawing,7,2)) to clayout
		else
			store substr(cdrawing,1,5) + '-' + upper(substr(cdrawing,6,1)) to clayout
		endif
	endif
	if lonnetwork = .f.
		store cdrawing to clayout
	endif
	* the group table was updated to 17 fields FREIGHT ADDED 05/01/03
	if file(ctabledir+cdrawing+'G.DBF') = .t.
		* open file and check then close it
		use ctabledir+cdrawing+'G.DBF' in a
		if fcount("a") < 17 .or. fsize('gname1') = 24
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
	* from the openfile program in the quotation application
	* restore memory variables by reading variables from the cdrawing + 'V.DBF' file
	* open tables
	* because windows continues this program while the run statement is being executed
	*  the file may not be in place yet - thus a do while is being used
	store .f. to lopened
	store .t. to lfilethere
	store 0 to ndowhilecounter
	do while lopened = .f.
		if file(ctabledir + cdrawing + 'V.DBF') = .t.
			use ctabledir + cdrawing + 'V.DBF' in a
			store .t. to lopened
		else
			* do nothing
			store ndowhilecounter + 1 to ndowhilecounter
			if ndowhilecounter > 1000
				store .t. to lopened
				set message to 'No pricing files available.'
				wait 'No pricing files available.' window at 5,10 timeout 2
				store .f. to lfilethere
			endif
		endif
	enddo
	if lonnetwork = .t.
		if lfilethere = .t.
			use location in b order num_acs shared
			use quotetrk in c order dwgnum  shared
			use quothist in d order dwgnum  shared
			PUBLIC lnodrawing
			store .f. to lnodrawing
			if upper(alias("A")) = 'A' .and. upper(alias("B")) = 'LOCATION' .and. ;
			   upper(alias("C")) = 'QUOTETRK' .and. upper(alias("D")) = 'QUOTHIST'
				* set message
				set message to 'Restoring variables.'
				wait 'Restoring variables.' window at 5,10 timeout 2
				* ensure work area
				select a
				* ensure that the relation is to the correct quote table
				if seek(DWGNUM,"quotetrk") = .t.
					set relation to dwgnum into quotetrk, num_acs into location
					use in d    && quotetrk selected thus close quothist
					ctitle = alltrim(C.QTITLE)
				else
					if seek(DWGNUM,"quothist") = .t.
						set relation to dwgnum into quothist, num_acs into location
						use in c    && quothist selected thus close quotetrk
						ctitle = alltrim(D.QTITLE)
					else
						cMessageTitle = 'Drawing Not Found'
						cMessageText = 'Please check quote files for correct drawing number ?'
						nDialogType = 0+16+0
						*   0 = OK button
						*  16 = Stop sign
						*   0 = First button is default
						nnodwg = MESSAGEBOX(cMessageText, nDialogType, cMessageTitle)
						lnodrawing = .t.
					endif
				endif
				* still in a
				go top
				* if lnodrawing = .f. continue
				if lnodrawing = .f.
					* variables from a (ctabledir + cdrawing + 'v.DBF')
					cquote = alltrim(a.NUM_QUOTE)
					* variable to determine if the customer has a discount
					PUBLIC nmaterialdisc, ninstalldisc
					store a.discount to nmaterialdisc     && material discount
					store a.discounti to ninstalldisc     && installation discount
					* variable for determining the proper company template
					PUBLIC ncompnumber,ncompleteacsno
					ncompnumber = floor(a.NUM_ACS/1000)
					ncompleteacsno = a.NUM_ACS
					* variables from b (location)
					ccustomer = alltrim(b.COMPANY)
					clocation = iif(empty(b.S_CITY),"",trim(b.S_CITY)) + ;
					            iif(empty(b.S_STATE),"",(", " + trim(b.S_STATE))) + ;
					            iif(trim(b.S_COUNTRY) # "USA",(", " + trim(b.S_COUNTRY)),"")
					* Added 9/25/03 to check for Mexico quotes Juan or Eddie
					PUBLIC linmex
					linmex = iif(upper(alltrim(b.S_COUNTRY)) = "MEXICO",.t.,.f.)
					* variable to determine if the note about US Dollars will be used
					PUBLIC ldeletedollars
					if upper(alltrim(b.m_country)) = "USA"
						ldeletedollars = .t.
					else
						ldeletedollars = .f.
					endif
					* variable for Guarantee & Terms for Non USA Quotes
					* will not include Canada or Avanti quotes
					* lcoocfgt = Logical Company Out Of Country For Guarantee & Terms
					PUBLIC lcoocfgt
					if upper(alltrim(b.s_country)) = "USA" .or. ;
					   upper(alltrim(b.s_country)) = "CANADA"
						lcoocfgt = .f.
					else
						lcoocfgt = .t.
					endif
					* variable that determines the quote is for AVANTI
					* if the quote is for Avanti lcoocfgt will become false
					PUBLIC lforavanti
					if upper(substr(b.b_address1,1,6)) = "AVANTI"
						lforavanti = .t.
						lcoocfgt = .f.
					else
						lforavanti = .f.
					endif
					close tables all
				endif
			else
				set message to "All the tables were not opened."
				wait "All the tables were not opened." window at 6,10 timeout 3
				set message to "The variable information was not updated."
				wait "The variable information was not updated." window at 7,10 timeout 3
			endif
		endif
	else    && not on network
		use cdatapath + 'location' in b order num_acs shared
		if upper(alias("B")) = 'LOCATION'
			* use ctabledir + cdrawing + 'V.DBF' in a
			* no Quotetrk or quothist
			* need a value for ctitle
			if file(ctabledir+cdrawing+'N.DBF') = .f.
				set message to 'Creating ' + upper(cdrawing) + 'N.DBF'
				wait 'Creating ' + upper(cdrawing) + 'N.DBF' window at 6,10 timeout 2
				use in c
				select c
				use cmoldpath + 'moldnonet' in c shared noupdate
				copy structure to ctabledir + cdrawing + 'N.DBF'
				use ctabledir + cdrawing + 'N.DBF' in c exclusive
				append blank
				replace C.QUOTETITLE with 'Please name this quotation.'
			else
				use ctabledir + cdrawing + 'N.DBF' in c exclusive
			endif
			* open table and revise C.QUOTETITLE
			do form cformspath + 'gettitle'
			* 
			ctitle = alltrim(C.QUOTETITLE)
			select a
			go top
			* variables from a (ctabledir + cdrawing + 'v.DBF')
			cquote = alltrim(a.NUM_QUOTE)
			* variable to determine if the customer has a discount
			PUBLIC nmaterialdisc, ninstalldisc
			store a.discount to nmaterialdisc     && material discount
			store a.discounti to ninstalldisc     && installation discount
			* set relation for location
			set relation to num_acs into location
			* variable for determining the proper company template
			PUBLIC ncompnumber,ncompleteacsno
			ncompnumber = floor(a.NUM_ACS/1000)
			ncompleteacsno = a.NUM_ACS
			* variables from b (location)
			ccustomer = alltrim(b.COMPANY)
			clocation = iif(empty(b.S_CITY),"",trim(b.S_CITY)) + ;
			            iif(empty(b.S_STATE),"",(", " + trim(b.S_STATE))) + ;
			            iif(trim(b.S_COUNTRY) # "USA",(", " + trim(b.S_COUNTRY)),"")
			* Added 9/25/03 to check for Mexico quotes Juan or Eddie
			PUBLIC linmex
			linmex = iif(upper(alltrim(b.S_COUNTRY)) = "MEXICO",.t.,.f.)
			* variable to determine if the note about US Dollars will be used
			PUBLIC ldeletedollars
			if upper(alltrim(b.m_country)) = "USA"
				ldeletedollars = .t.
			else
				ldeletedollars = .f.
			endif
			* variable for Guarantee & Terms for Non USA Quotes
			* will not include Canada or Avanti quotes
			* lcoocfgt = Logical Company Out Of Country For Guarantee & Terms
			PUBLIC lcoocfgt
			if upper(alltrim(b.s_country)) = "USA" .or. ;
			   upper(alltrim(b.s_country)) = "CANADA"
				lcoocfgt = .f.
			else
				lcoocfgt = .t.
			endif
			* variable that determines the quote is for AVANTI
			* if the quote is for Avanti lcoocfgt will become false
			PUBLIC lforavanti
			if upper(substr(b.b_address1,1,6)) = "AVANTI"
				lforavanti = .t.
				lcoocfgt = .f.
			else
				lforavanti = .f.
			endif
			close tables all
		else
			set message to "All the tables were not opened."
			wait "All the tables were not opened." window at 6,10 timeout 3
			set message to "The variable information was not updated."
			wait "The variable information was not updated." window at 7,10 timeout 3
		endif
	endif
endif

* set message
set message to 'Completed.'
wait 'Completed.' window at 5,10 timeout 5

* clean up
set message to

*-- EOP SELECTDWGNUMBER