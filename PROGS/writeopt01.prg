* Name......... WRITE the OPTion 01 information from foxpro into word program
* Date......... 10/22/2010
* Caller....... See Notes
* Notes........ Consolidating the Option 01 printing
*               Does Includes pricepageavanti.prg
*               Does Includes pricepagemaid.prg
*               Does Includes pricepagemodnetonly
*               Does Includes pricepagemodshown
*               Does Includes pricepagend

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

if isblank(option_ref) = .t.    && no reference
	* if a.material = 0 and a.freight = 0 and a.install = 0 and a.rtotal = 0 and a.trnprice = 0 and a.pc_tech = 0
	if a.material = 0 and a.freight = 0 and a.install = 0 and a.rtotal = 0 and a.pc_tech = 0
		* template information must be removed then add a paragraph
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.HomeKey(wdline,1) &&    Selection.HomeKey Unit:=wdLine
		WordApp.Selection.Delete(wdWord,21)
		* all template information cleared at this point
	else
		* move cursor to the beginning of the Add Material statement
		WordApp.Selection.MoveRight(wdCharacter,4)
		store 0 to ndismaterial
		if a.material = 0    && remove the line for material
			* at this time the cursor is somewhere on the Add Material line thus delete
			WordApp.Selection.StartOf
			WordApp.Selection.TypeBackspace
			WordApp.Selection.TypeBackspace
			WordApp.Selection.Delete(wdCharacter,15)
			* cursor is at the beginning
		else
			do case
*!*				case ncompnumber = 546
*!*					* store 0 to nmaterialdisc to get list
*!*					nmaterialdisc = 0
*!*					WordApp.Selection.MoveRight(wdCharacter,4)
*!*					WordApp.Selection.InsertBefore("List ")
			case nmaterialdisc > 0
				if a.material < 0
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Deduct Net")
				else
					WordApp.Selection.MoveRight(wdCharacter,4)
					WordApp.Selection.InsertBefore("Net ")
				endif
			otherwise
				if a.material < 0
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Deduct")
				endif
			endcase
			WordApp.Selection.EndKey
			store round(a.material * nmaterialdisc,0) to ndismaterial
			do numberformatopt with "a.material - ndismaterial"
			WordApp.Selection.MoveRight(wdCharacter,2)
			* cursor is at the beginning
		endif    && if a.material = 0
		* at his point the cursor is at the beginning of the line
		store 0 to ndisinstall
		if a.install = 0    && remove the line for install
			* at this time the cursor is at the beginning of the line
			WordApp.Selection.Delete(wdCharacter,21)
			* at this time the cursor is at the beginning of the ADD TOTAL line
		else
			do case
			case lforavanti = .t. .or. ninstalldisc = 0
				if a.install < 0
				    WordApp.Selection.MoveRight(wdCharacter,2)
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Deduct")
				endif
			case ninstalldisc > 0
				if a.install < 0
				    WordApp.Selection.MoveRight(wdCharacter,2)
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Deduct Net")
				else
				    WordApp.Selection.MoveRight(wdCharacter,2)
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Add Net")
				endif
			endcase
			WordApp.Selection.EndKey
			store round(a.install * ninstalldisc,0) to ndisinstall
			do numberformatopt with "a.install - ndisinstall"
			WordApp.Selection.MoveRight(wdCharacter,2)
			* cursor is at the beginning of the ADD TOTAL line
		endif   && if a.install = 0
		* at his point the cursor is at the beginning of the ADD TOTAL line
		* add rtotal if necessary
		store 0 to ndiscr
		if a.rtotal # 0    && don't add the line for rtotal
	    	WordApp.Selection.TypeText(Chr(9)+Chr(9))
			do case
			case lforavanti = .t. .or. ninstalldisc = 0
				if a.rtotal < 0
					WordApp.Selection.TypeText("Deduct Relocation, Rework & Removal"+Chr(9))
				else
					WordApp.Selection.TypeText("Add Relocation, Rework & Removal"+Chr(9))
				endif
			case ninstalldisc > 0
				if a.rtotal < 0
					WordApp.Selection.TypeText("Deduct Net Relocation, Rework & Removal"+Chr(9))
				else
					WordApp.Selection.TypeText("Add Net Relocation, Rework & Removal"+Chr(9))
				endif
			endcase
			store round(a.rtotal * ninstalldisc,0) to ndiscr
	    	do numberformatopt with "a.rtotal - ndiscr"
			WordApp.Selection.MoveRight
	    	WordApp.Selection.TypeText(Chr(11))
		endif
*!*			* if rtotal not used then the cursor is at the beginning of the ADD TOTAL line
*!*			* if rtotal is used then the cursor is at the beginning of the ADD TOTAL line
*!*			if a.trnprice # 0    && don't add the line for training
*!*		    	WordApp.Selection.TypeText(Chr(9)+Chr(9))

*!*					if a.trnprice < 0
*!*						WordApp.Selection.TypeText("Deduct Training"+Chr(9))
*!*					else
*!*						WordApp.Selection.TypeText("Add Training"+Chr(9))
*!*					endif

*!*				* store round(a.rtotal * ninstalldisc,0) to ndiscr
*!*		    	do numberformatopt with "a.trnprice"
*!*				WordApp.Selection.MoveRight
*!*		    	WordApp.Selection.TypeText(Chr(11))
*!*			endif
		* if rtotal not used then the cursor is at the beginning of the ADD TOTAL line
		* if rtotal is used then the cursor is at the beginning of the ADD TOTAL line
		* add pc_tech if necessary
		if a.pc_tech # 0    && don't add the line for pc_tech
	    	WordApp.Selection.TypeText(Chr(9)+Chr(9))
			if a.pc_tech < 0
				WordApp.Selection.TypeText("Deduct PLC Start Up"+Chr(9)+Chr(9))
			else
				WordApp.Selection.TypeText("Add PLC Start Up"+Chr(9)+Chr(9))
			endif
	    	do numberformatopt with "a.pc_tech"
			WordApp.Selection.MoveRight
	    	WordApp.Selection.TypeText(Chr(11))
		endif
		* AVANTI QUOTES ARE THE ONLY ONES THAT USE THIS FOR NOW
		* at his point the cursor is at the beginning of the ADD TOTAL line
		* add freight if necessary
		if a.freight # 0 .and. lforavanti = .t.   && don't add the line for freight
	    	WordApp.Selection.TypeText(Chr(9)+Chr(9))
			if a.freight < 0
				WordApp.Selection.TypeText("Deduct Estimated Freight"+Chr(9)+Chr(9))
			else
				WordApp.Selection.TypeText("Add Estimated Freight"+Chr(9)+Chr(9))
			endif
	    	do numberformatopt with "a.freight"
			WordApp.Selection.MoveRight
	    	WordApp.Selection.TypeText(Chr(11))
		endif
		* if freight not used then the cursor is at the beginning of the ADD TOTAL line
		* if freight is used then the cursor is at the beginning of the ADD TOTAL line
		if a.total = a.material .or. a.total = a.freight .or. a.total = a.install .or. ;
		   a.total = a.rtotal .or. a.total = a.pc_tech
			* don't print this line
			* at this time the cursor is at the beginning of the line
			WordApp.Selection.Delete(wdCharacter,14)
			WordApp.Selection.TypeBackspace
			* at this time the cursor is at the beginning of the ADD TOTAL line
		else
			do case
*!*				case ncompnumber = 546
*!*					* store 0 to nmaterialdisc to get list
*!*					* nmaterialdisc = 0
*!*					WordApp.Selection.MoveRight(wdCharacter,5)
*!*					WordApp.Selection.InsertBefore(" NET")
			case nmaterialdisc > 0 .or. ninstalldisc > 0
				if a.total < 0
				    WordApp.Selection.MoveRight(wdCharacter,5)
					WordApp.Selection.InsertBefore("DEDUCT NET")
					* Deduct must be written within the paragraph
					*  thus you must go back to the start of and delete the word add
					WordApp.Selection.StartOf
				    WordApp.Selection.Delete(wdCharacter,3)
				else
				    WordApp.Selection.MoveRight(wdCharacter,5)
				    if a.material # 0
						WordApp.Selection.InsertBefore(" NET")
					endif
				endif
			otherwise
				if a.total < 0
				    WordApp.Selection.MoveRight(wdCharacter,5)
					WordApp.Selection.InsertBefore("DEDUCT")
					* Deduct must be written within the paragraph
					*  thus you must go back to the start of and delete the word add
					WordApp.Selection.StartOf
				    WordApp.Selection.Delete(wdCharacter,3)
				else
				endif
			endcase
			WordApp.Selection.EndKey
			do numberformatopt with "a.total - ndismaterial - ndisinstall - ndiscr"
			* a.total was formatted and must be written within the paragraph
			* thus you must go back to the start of and delete the '$'
			WordApp.Selection.StartOf
		    WordApp.Selection.Delete(wdCharacter,1)
		endif
	endif   && if material, freight, install, rtotal, training and pc_tech all equal zero
else    && reference used
	* add to group name
	WordApp.Selection.InsertAfter(" IN LIEU OF GROUP " + alltrim(option_ref))
*!*		if (a.material - g.material) = 0 .and. (a.freight - g.freight) = 0 .and. ;
*!*		   (a.install - g.install) = 0 .and. (a.rtotal - g.rtotal) = 0 .and. ;
*!*		   (a.pc_tech - g.pc_tech) = 0 .and. (a.trnprice - g.trnprice) = 0
	if (a.material - g.material) = 0 .and. (a.freight - g.freight) = 0 .and. ;
	   (a.install - g.install) = 0 .and. (a.rtotal - g.rtotal) = 0 .and. ;
	   (a.pc_tech - g.pc_tech) = 0
		* template information must be removed then add a paragraph
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.HomeKey(wdline,1) &&    Selection.HomeKey Unit:=wdLine
		WordApp.Selection.Delete(wdWord,21)
		* all template information cleared at this point
	else
		* move cursor to the beginning of the Add Material statement
		WordApp.Selection.MoveRight(wdCharacter,4)
		store 0 to ndismaterial
		if a.material - g.material = 0    && remove the line for material
			* at this time the cursor is somewhere on the Add Material line thus delete
			WordApp.Selection.StartOf
			WordApp.Selection.TypeBackspace
			WordApp.Selection.TypeBackspace
			WordApp.Selection.Delete(wdCharacter,15)
			* cursor is at the beginning
		else
			do case
			case ncompnumber = 546
				* store 0 to nmaterialdisc to get list
				nmaterialdisc = 0
				WordApp.Selection.MoveRight(wdCharacter,4)
				WordApp.Selection.InsertBefore("List ")
			case nmaterialdisc > 0
				if a.material - g.material < 0
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Deduct Net")
				else
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Add Net")
				endif
			otherwise
				if a.material - g.material < 0
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Deduct")
				endif				
			endcase
			WordApp.Selection.EndKey
			store round(((a.material - g.material) * nmaterialdisc),0) to ndismaterial
			do numberformatopt with "(a.material - g.material) - ndismaterial"
			WordApp.Selection.MoveRight(wdCharacter,2)
			* cursor is at the beginning
		endif
		* at his point the cursor is at the beginning of the line
		store 0 to ndisinstall
		if a.install - g.install = 0    && remove the line for install
			* at this time the cursor is at the beginning of the line
			WordApp.Selection.Delete(wdCharacter,21)
			* at this time the cursor is at the beginning of the ADD TOTAL line
		else
			do case
			case lforavanti = .t. .or. ninstalldisc = 0
				if a.install - g.install < 0
				    WordApp.Selection.MoveRight(wdCharacter,2)
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Deduct")
				endif
			case ninstalldisc > 0
				if a.install - g.install < 0
				    WordApp.Selection.MoveRight(wdCharacter,2)
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Deduct Net")
				else
				    WordApp.Selection.MoveRight(wdCharacter,2)
				    WordApp.Selection.Delete(wdCharacter,3)
					WordApp.Selection.InsertBefore("Add Net")
				endif
			endcase
			WordApp.Selection.EndKey
			store round((a.install - g.install) * ninstalldisc,0) to ndisinstall
			do numberformatopt with "(a.install - g.install) - ndisinstall"
			WordApp.Selection.MoveRight(wdCharacter,2)
			* cursor is at the beginning of the ADD TOTAL line
		endif
		* at his point the cursor is at the beginning of the ADD TOTAL line
		* add rtotal if necessary
		store 0 to ndiscr
		if a.rtotal - g.rtotal # 0    && don't add the line for rtotal
	    	WordApp.Selection.TypeText(Chr(9)+Chr(9))
			do case
			case lforavanti = .t. .or. ninstalldisc = 0
				if a.rtotal - g.rtotal < 0
					WordApp.Selection.TypeText("Deduct Relocation, Rework & Removal"+Chr(9))
				else
					WordApp.Selection.TypeText("Add Relocation, Rework & Removal"+Chr(9))
				endif
			case ninstalldisc > 0
				if a.rtotal - g.rtotal < 0
					WordApp.Selection.TypeText("Deduct Net Relocation, Rework & Removal"+Chr(9))
				else
					WordApp.Selection.TypeText("Add Net Relocation, Rework & Removal"+Chr(9))
				endif
			endcase
			store round((a.rtotal - g.rtotal) * ninstalldisc,0) to ndiscr
	    	do numberformatopt with "(a.rtotal - g.rtotal) - ndiscr"
			WordApp.Selection.MoveRight
	    	WordApp.Selection.TypeText(Chr(11))
		endif
*!*			* if rtotal not used then the cursor is at the beginning of the ADD TOTAL line
*!*			* if rtotal is used then the cursor is at the beginning of the ADD TOTAL line
*!*			* add training if necessary
*!*			if a.trnprice - g.trnprice # 0    && don't add the line for training
*!*				WordApp.Selection.TypeText(Chr(9)+Chr(9))
*!*				if a.trnprice - g.trnprice < 0
*!*					WordApp.Selection.TypeText("Deduct Training"+Chr(9)+Chr(9))
*!*				else
*!*					WordApp.Selection.TypeText("Add Training"+Chr(9)+Chr(9))
*!*				endif
*!*				do numberformatopt with "(a.trnprice - g.trnprice)"
*!*				WordApp.Selection.MoveRight
*!*				WordApp.Selection.TypeText(Chr(11))
*!*			endif	
		* add pc_tech if necessary
		if a.pc_tech - g.pc_tech # 0    && don't add the line for pc_tech
	    	WordApp.Selection.TypeText(Chr(9)+Chr(9))
			if a.pc_tech - g.pc_tech < 0
				WordApp.Selection.TypeText("Deduct PLC Start Up"+Chr(9)+Chr(9))
			else
				WordApp.Selection.TypeText("Add PLC Start Up"+Chr(9)+Chr(9))
			endif
	    	do numberformatopt with "a.pc_tech - g.pc_tech"
			WordApp.Selection.MoveRight
	    	WordApp.Selection.TypeText(Chr(11))
		endif
		* AVANTI QUOTES ARE THE ONLY ONES THAT USE THIS FOR NOW
		* at this point the cursor is at the beginning of the ADD TOTAL line
		* add freight if necessary
		if a.freight - g.freight # 0 .and. lforavanti = .t.   && don't add the line for freight
	    	WordApp.Selection.TypeText(Chr(9)+Chr(9))
			if a.freight - g.freight < 0
				WordApp.Selection.TypeText("Deduct Estimated Freight"+Chr(9)+Chr(9))
			else
				WordApp.Selection.TypeText("Add Estimated Freight"+Chr(9)+Chr(9))
			endif
	    	do numberformatopt with "a.freight - g.freight"
			WordApp.Selection.MoveRight
	    	WordApp.Selection.TypeText(Chr(11))
		endif
		* if freight not used then the cursor is at the beginning of the ADD TOTAL line
		* if freight is used then the cursor is at the beginning of the ADD TOTAL line
		* a.total - g.total = a.material - g.material .or. a.total - g.total = a.install - g.install .or. ;
		* a.total - g.total = a.rtotal - g.rtotal .or. a.total - g.total = a.pc_tech - g.pc_tech
		* don't print this line
		if a.total - g.total = a.material - g.material .or. a.total - g.total = a.install - g.install .or. ;
		   a.total - g.total = a.rtotal - g.rtotal .or. a.total - g.total = a.pc_tech - g.pc_tech
			* don't print this line
			* at this time the cursor is at the beginning of the line
			WordApp.Selection.Delete(wdCharacter,14)
			WordApp.Selection.TypeBackspace
			* at this time the cursor is at the beginning of the ADD TOTAL line
		else
			do case
*!*				case ncompnumber = 546
*!*					* store 0 to nmaterialdisc to get list
*!*					* nmaterialdisc = 0
*!*					WordApp.Selection.MoveRight(wdCharacter,5)
*!*					WordApp.Selection.InsertBefore(" NET")
			case nmaterialdisc > 0 .or. ninstalldisc > 0
				if a.total - g.total < 0
				    WordApp.Selection.MoveRight(wdCharacter,5)
					WordApp.Selection.InsertBefore("DEDUCT NET")
				else
				    WordApp.Selection.MoveRight(wdCharacter,5)
					if a.material - g.material # 0 .or. a.install - g.install # 0 .or. a.rtotal - g.rtotal # 0
						WordApp.Selection.InsertBefore(" NET")
					else
						WordApp.Selection.InsertBefore("ADD")
					endif
				endif
				* Deduct net and Add net must be written within the paragraph
				*  thus you must go back to the start of and delete the word add
				WordApp.Selection.StartOf
				WordApp.Selection.Delete(wdCharacter,3)
			otherwise
				if a.total - g.total < 0
				    WordApp.Selection.MoveRight(wdCharacter,5)
					WordApp.Selection.InsertBefore("DEDUCT")
					* Deduct must be written within the paragraph
					*  thus you must go back to the start of and delete the word add
					WordApp.Selection.StartOf
				    WordApp.Selection.Delete(wdCharacter,3)
				endif
			endcase
			WordApp.Selection.EndKey
			do numberformatopt with "a.total - g.total - ndismaterial - ndisinstall - ndiscr"
			* g.total was formatted and must be written within the paragraph
			* thus you must go back to the start of and delete the '$'
			WordApp.Selection.StartOf
		    WordApp.Selection.Delete(wdCharacter,1)
		endif
	endif   && material, freight, install, rtotal, training and pc_tech all equal zero if
endif    && option reference

*-- EOP WRITOPT01