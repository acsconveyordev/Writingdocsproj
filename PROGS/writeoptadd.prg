* Name......... WRITE the OPTions ADDed after option 01 information from foxpro into word program
* Date......... 10/22/2010
* Caller....... See Notes
* Notes........ Consolidating the Additional Options printing
*               Does Includes pricepageavanti.prg
*               Does Includes pricepagemaid.prg
*               Does Includes pricepagemodnetonly
*               Does Includes pricepagemodshown
*               Does Not Includes pricepagend

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

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
	* go to end then move down one line
    WordApp.Selection.EndKey
	WordApp.Selection.MoveDown(wdLine,1)
	* now go back and format previous paragraph
	* paragraph 111
	wordapp.ActiveDocument.Paragraphs(111+pcount).Range.Select
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
	* move down to the paragraph that was added and insert information
	WordApp.Selection.MoveDown(wdLine,1)
	* determine if a reference has been used
	if isblank(option_ref) = .t.    && no reference
		store 0 to ndismaterial
*!*			if a.material = 0 .and. a.freight = 0 .and. a.install = 0 .and. a.rtotal = 0 .and. a.pc_tech = 0 .and. a.trnprice = 0
		if a.material = 0 .and. a.freight = 0 .and. a.install = 0 .and. a.rtotal = 0 .and. a.pc_tech = 0
			* just add a line between titles
		else
			* add material if necessary
			do case
*!*				case ncompnumber = 546
*!*					if a.material = 0
*!*						* don't add the line for material
*!*					else
*!*						* add information
*!*						if a.material < 0
*!*							WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct List Material"+Chr(9)+Chr(9))
*!*						else
*!*							WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add List Material"+Chr(9)+Chr(9))
*!*						endif
*!*						WordApp.Selection.EndKey
*!*						store 0 to ndismaterial
*!*						do numberformatopt with "a.material"
*!*					endif
			case nmaterialdisc > 0
				if a.material = 0
					* don't add the line for material
				else
					* add information
					if a.material < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Net Material"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Net Material"+Chr(9)+Chr(9))
					endif
					WordApp.Selection.EndKey
					store round(a.material * nmaterialdisc,0) to ndismaterial
					do numberformatopt with "a.material - ndismaterial"
				endif
			otherwise
				if a.material = 0
					* don't add the line for material
				else
					* add information
					if a.material < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Material"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Material"+Chr(9)+Chr(9))
					endif
					WordApp.Selection.EndKey
					do numberformatopt with "a.material"
				endif
			endcase
			* add installation if necessary
			store 0 to ndisinstall
			if a.install = 0    && don't add the line for installation
				* don't add the line for installation
			else
				do case
				case lforavanti = .t. .or. ninstalldisc = 0
					* add information
					WordApp.Selection.EndKey
					iif(a.material # 0,WordApp.Selection.TypeText(Chr(11)),"")
					if a.install < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Installation"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Installation"+Chr(9)+Chr(9))
					endif
					WordApp.Selection.EndKey
					do numberformatopt with "a.install"
				case ninstalldisc > 0
					* add information
					WordApp.Selection.EndKey
					iif(a.material # 0,WordApp.Selection.TypeText(Chr(11)),"")
					if a.install < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Net Installation"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Net Installation"+Chr(9)+Chr(9))
					endif
					WordApp.Selection.EndKey
					store round(a.install * ninstalldisc,0) to ndisinstall
					do numberformatopt with "a.install - ndisinstall"
				endcase
			endif   && if a.install = 0
			* add rtotal if necessary
			store 0 to ndiscr
			if a.rtotal # 0    && don't add the line for rtotal
				* there are four cases here to determine if a chr(11) is used
				* 1 material=0 and installation=0  not needed
				* all others needed
				WordApp.Selection.EndKey
				if a.material = 0 .and. a.install = 0
				else
					WordApp.Selection.TypeText(Chr(11))
				endif
				do case
				case lforavanti = .t. .or. ninstalldisc = 0
					* add information
					if a.rtotal < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Relocation, Rework & Removal"+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Relocation, Rework & Removal"+Chr(9))
					endif
				case ninstalldisc > 0
					* add information
					if a.rtotal < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Net Relocation, Rework & Removal"+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Net Relocation, Rework & Removal"+Chr(9))
					endif
				endcase
				WordApp.Selection.EndKey
				store round(a.rtotal * ninstalldisc,0) to ndiscr
		    	do numberformatopt with "a.rtotal - ndiscr"
			endif   && if a.rtotal # 0
*!*				* add training if necessary
*!*				if a.trnprice - g.trnprice # 0    && don't add the line for training
*!*		  			WordApp.Selection.TypeText(Chr(9)+Chr(9))
*!*					if a.trnprice - g.trnprice < 0
*!*						WordApp.Selection.TypeText("Deduct Training"+Chr(9)+Chr(9))
*!*					else
*!*						WordApp.Selection.TypeText("Add Training"+Chr(9)+Chr(9))
*!*					endif
*!*					do numberformatopt with "(a.trnprice - g.trnprice)"
*!*					WordApp.Selection.MoveRight
*!*					WordApp.Selection.TypeText(Chr(11))
*!*				endif
			* add pc_tech if necessary
			if a.pc_tech # 0    && don't add the line for pc_tech
				* there are eight cases here to determine if a chr(11) is used
				* 1 material=0 and installation=0 and rtotal=0  not needed
				* all others needed
				WordApp.Selection.EndKey
				if a.material = 0 .and. a.install = 0 .and. a.rtotal = 0
				else
					WordApp.Selection.TypeText(Chr(11))
				endif
				* add information
				if a.pc_tech < 0
					WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct PLC Start Up"+Chr(9))
				else
					WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add PLC Start Up"+Chr(9)+Chr(9))
				endif
				WordApp.Selection.EndKey
				do numberformatopt with "a.pc_tech"
			endif    && if a.pc_tech # 0
			* AVANTI QUOTES ARE THE ONLY ONES THAT USE THIS FOR NOW
			* at his point the cursor is at the beginning of the ADD TOTAL line
			* add freight if necessary
			if a.freight # 0 .and. lforavanti = .t.   && don't add the line for freight
				* there are eight cases here to determine if a chr(11) is used
				* 1 material=0 and freight=0 installation=0 and rtotal=0  not needed
				* all others needed
				WordApp.Selection.EndKey
				if a.material = 0 .and. a.pc_tech = 0 .and. a.install = 0 .and. a.rtotal = 0
				else
					WordApp.Selection.TypeText(Chr(11))
				endif
				* add information
				if a.freight < 0
					WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Estimated Freight"+Chr(9)+Chr(9))
				else
					WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Estimated Freight"+Chr(9)+Chr(9))
				endif
				WordApp.Selection.EndKey
				do numberformatopt with "a.freight"
			endif
			WordApp.Selection.EndKey
			WordApp.Selection.TypeText(Chr(11)+Chr(9)+Chr(9))
			* add total if necessary
			* a.total = a.material .or. a.total = a.freight .or. a.total = a.install .or. ;
			* a.total = a.rtotal .or. a.total = a.pc_tech
			if a.total = a.material .or. a.total = a.freight .or. a.total = a.install .or. ;
			   a.total = a.rtotal .or. a.total = a.pc_tech
				* don't print this line
				* at this time the cursor is at the beginning of the line
				WordApp.Selection.TypeBackSpace
				WordApp.Selection.TypeBackSpace
				WordApp.Selection.TypeBackSpace
				* at this time the cursor is at the beginning of the ADD TOTAL line
			else
				do case
				case ncompnumber = 546
					if a.total < 0
						WordApp.Selection.InsertAfter("DEDUCT NET TOTAL")
					else
						if a.material # 0
							WordApp.Selection.InsertAfter("ADD NET TOTAL")
						else
							WordApp.Selection.InsertAfter("ADD TOTAL")
						endif
					endif
				case nmaterialdisc > 0 .or. ninstalldisc > 0
					if a.total < 0
						WordApp.Selection.InsertAfter("DEDUCT NET TOTAL")
					else
						if a.material # 0
							WordApp.Selection.InsertAfter("ADD NET TOTAL")
						else
							WordApp.Selection.InsertAfter("ADD TOTAL")
						endif
					endif
				otherwise
					if a.total < 0
						WordApp.Selection.InsertAfter("DEDUCT TOTAL")
					else
						WordApp.Selection.InsertAfter("ADD TOTAL")
					endif
				endcase
				* selection must be formatted to bold
				WordApp.Selection.Font.Bold = .t.
				WordApp.Selection.EndKey
				WordApp.Selection.TypeText(Chr(9)+Chr(9))
				do numberformatopt with "a.total - ndismaterial - ndisinstall - ndiscr"
				* selection must be formatted to add double underline
				WordApp.Selection.Font.Bold = .t.
				WordApp.Selection.Font.Underline = wdUnderlineDouble
				WordApp.Selection.EndKey
			endif
		endif
	else    && reference used
*!*			if (a.material - g.material) = 0 .and. (a.freight - g.freight) = 0 .and. ;
*!*			   (a.install - g.install) = 0 .and. (a.rtotal - g.rtotal) = 0 .and. ;
*!*			   (a.pc_tech - g.pc_tech) = 0 .and. (a.trnprice - g.trnprice) = 0
		if (a.material - g.material) = 0 .and. (a.freight - g.freight) = 0 .and. ;
		   (a.install - g.install) = 0 .and. (a.rtotal - g.rtotal) = 0 .and. ;
		   (a.pc_tech - g.pc_tech) = 0
			* do nothing
		else
			* add material if necessary
			store 0 to ndismaterial
			if a.material - g.material = 0
				* don't add the line for material
			else
				do case
*!*					case ncompnumber = 546
*!*						* add information
*!*						if a.material - g.material < 0
*!*							WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct List Material"+Chr(9)+Chr(9))
*!*						else
*!*							WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add List Material"+Chr(9)+Chr(9))
*!*						endif
*!*						WordApp.Selection.EndKey
*!*						store 0 to ndismaterial
*!*						do numberformatopt with "(a.material - g.material)"
				case nmaterialdisc > 0 
					* add information
					if a.material - g.material < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Net Material"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Net Material"+Chr(9)+Chr(9))
					endif
					WordApp.Selection.EndKey
					store round((a.material - g.material) * nmaterialdisc,0) to ndismaterial
					do numberformatopt with "(a.material - g.material) - ndismaterial"
				otherwise
					* add information
					if a.material - g.material < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Material"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Material"+Chr(9)+Chr(9))
					endif
					WordApp.Selection.EndKey
					do numberformatopt with "a.material - g.material"
				endcase
			endif   && a.material - g.material = 0
			* add installation if necessary
			store 0 to ndisinstall
			if a.install - g.install = 0    && don't add the line for installation
				* don't add the line for installation
			else
				do case
				case lforavanti = .t. .or. ninstalldisc = 0
					* add information
					WordApp.Selection.EndKey
					iif(a.material # 0,WordApp.Selection.TypeText(Chr(11)),"")
					if a.install - g.install < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Installation"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Installation"+Chr(9)+Chr(9))
					endif
				case ninstalldisc > 0
					* add information
					WordApp.Selection.EndKey
					iif(a.material # 0,WordApp.Selection.TypeText(Chr(11)),"")
					if a.install - g.install < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Net Installation"+Chr(9)+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Net Installation"+Chr(9)+Chr(9))
					endif
				endcase
				WordApp.Selection.EndKey
				store round((a.install - g.install) * ninstalldisc,0) to ndisinstall
				do numberformatopt with "(a.install - g.install) - ndisinstall"
			endif    && a.install - g.install = 0
			* add rtotal if necessary
			store 0 to ndiscr
			if a.rtotal - g.rtotal # 0    && don't add the line for rtotal
				* there are four cases here to determine if a chr(11) is used
				* 1 material=0 and installation=0  not needed
				* all others needed
				WordApp.Selection.EndKey
				if (a.material - g.material) = 0 .and. (a.install - g.install) = 0
				else
					WordApp.Selection.Typetext(Chr(11))
				endif
				do case
				case lforavanti = .t. .or. ninstalldisc = 0
					* add information
					if a.rtotal - g.rtotal < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Relocation, Rework & Removal"+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Relocation, Rework & Removal"+Chr(9))
					endif
				case ninstalldisc > 0
					* add information
					if a.rtotal - g.rtotal < 0
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Net Relocation, Rework & Removal"+Chr(9))
					else
						WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Net Relocation, Rework & Removal"+Chr(9))
					endif
				endcase
				WordApp.Selection.EndKey
				store round((a.rtotal - g.rtotal) * ninstalldisc,0) to ndiscr
		    	do numberformatopt with "(a.rtotal - g.rtotal) - ndiscr"
			endif    && a.rtotal - g.rtotal # 0
*!*				* add training if necessary
*!*				if a.trnprice - g.trnprice # 0    && don't add the line for training
*!*		  			WordApp.Selection.TypeText(Chr(9)+Chr(9))
*!*					if a.trnprice - g.trnprice < 0
*!*						WordApp.Selection.TypeText("Deduct Training"+Chr(9)+Chr(9))
*!*					else
*!*						WordApp.Selection.TypeText("Add Training"+Chr(9)+Chr(9))
*!*					endif
*!*					do numberformatopt with "(a.trnprice - g.trnprice)"
*!*					WordApp.Selection.MoveRight
*!*					WordApp.Selection.TypeText(Chr(11))
*!*				endif
			* add pc_tech if necessary
			if a.pc_tech - g.pc_tech # 0    && don't add the line for pc_tech
				* there are eight cases here to determine if a chr(11) is used
				* 1 material=0 and installation=0 and rtotal=0  not needed
				* all others needed
				WordApp.Selection.EndKey
				if a.material - g.material = 0 .and. a.install - g.install = 0 .and. a.rtotal - g.rtotal = 0
				else
					WordApp.Selection.TypeText(Chr(11))
				endif
				* add information
				if a.pc_tech - g.pc_tech < 0
					WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct PLC Start Up"+Chr(9)+Chr(9))
				else
					WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add PLC Start Up"+Chr(9)+Chr(9))
				endif
				WordApp.Selection.EndKey
				do numberformatopt with "a.pc_tech - g.pc_tech"
			endif    && a.pc_tech - g.pc_tech # 0
			* AVANTI QUOTES ARE THE ONLY ONES THAT USE THIS FOR NOW
			* at this point the cursor is at the beginning of the ADD TOTAL line
			* add freight if necessary
			if a.freight - g.freight # 0 .and. lforavanti = .t.   && don't add the line for freight
				* there are eight cases here to determine if a chr(11) is used
				* 1 material=0 and pc_tech=0 and installation=0 and rtotal=0  not needed
				* all others needed
				WordApp.Selection.EndKey
				if a.material - g.material = 0 .and. a.pc_tech - g.pc_tech = 0 .and. ;
				   a.install - g.install = 0 .and. a.rtotal - g.rtotal = 0
				else
					WordApp.Selection.TypeText(Chr(11))
				endif
				* add information
				if a.freight - g.freight < 0
					WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Deduct Estimated Freight"+Chr(9)+Chr(9))
				else
					WordApp.Selection.TypeText(Chr(9)+Chr(9)+"Add Estimated Freight"+Chr(9)+Chr(9))
				endif
				WordApp.Selection.EndKey
				do numberformatopt with "a.freight - g.freight"
			endif    && a.freight - g.freight # 0 .and. lforavanti = .t.
			* required to ensure placement of totals
			WordApp.Selection.EndKey
			WordApp.Selection.TypeText(Chr(11)+Chr(9)+Chr(9))
			* add total if necessary
			* a.total - g.total = a.material - g.material .or. a.total - g.total = a.freight - g.freight .or. ;
			* a.total - g.total = a.install - g.install .or. a.total - g.total = a.rtotal - g.rtotal .or. ;
			* a.total - g.total = a.pc_tech - g.pc_tech
			if a.total - g.total = a.material - g.material .or. a.total - g.total = a.freight - g.freight .or. ;
			   a.total - g.total = a.install - g.install .or. a.total - g.total = a.rtotal - g.rtotal .or. ;
			   a.total - g.total = a.pc_tech - g.pc_tech
				* don't print this line
				* at this time the cursor is at the beginning of the line
				WordApp.Selection.TypeBackspace
				WordApp.Selection.TypeBackspace
				WordApp.Selection.TypeBackspace
				* at this time the cursor is at the beginning of the ADD TOTAL line
			else
				do case
				case ncompnumber = 546
					if a.total - g.total < 0
						if a.material - g.material # 0
							WordApp.Selection.InsertAfter("DEDUCT NET TOTAL")
						endif
					else
						if a.material - g.material # 0 .or. a.install - g.install # 0 .or. a.rtotal - g.rtotal # 0
							WordApp.Selection.InsertAfter("ADD NET TOTAL")
						else
							WordApp.Selection.InsertAfter("ADD TOTAL")
						endif
					endif
				case nmaterialdisc > 0 .or. ninstalldisc > 0
					if a.total - g.total < 0
						if a.material - g.material # 0
							WordApp.Selection.InsertAfter("DEDUCT NET TOTAL")
						endif
					else
						if a.material - g.material # 0 .or. a.install - g.install # 0 .or. a.rtotal - g.rtotal # 0
							WordApp.Selection.InsertAfter("ADD NET TOTAL")
						else
							WordApp.Selection.InsertAfter("ADD TOTAL")
						endif
					endif
				otherwise
					if a.total - g.total < 0
						WordApp.Selection.InsertAfter("DEDUCT TOTAL")
					else
						WordApp.Selection.InsertAfter("ADD TOTAL")
					endif
				endcase
				* selection must be formatted to bold
				WordApp.Selection.Font.Bold = .t.
				WordApp.Selection.EndKey
				WordApp.Selection.TypeText(Chr(9)+Chr(9))
				do numberformatopt with "a.total - g.total - ndismaterial - ndisinstall - ndiscr"
				* do numberformatopt with "a.total - g.total"
				* selection must be formatted to add double underline
				WordApp.Selection.Font.Bold = .t.
				WordApp.Selection.Font.Underline = wdUnderlineDouble
				WordApp.Selection.EndKey
			endif
		endif    && (a.material - g.material) = 0 .and. (a.freight - g.freight) = 0 .and. ;
		         && (a.install - g.install) = 0 .and. (a.rtotal - g.rtotal) = 0 .and. ;
		         && (a.pc_tech - g.pc_tech) = 0
	endif    && option reference
	* do again
	* add a paragraph, then go back to the previous one
	WordApp.Selection.TypeParagraph
	WordApp.Selection.MoveUp(wdline,1)
	pcount = pcount + 2
	* the EndKey and TypeParagraph are at the beginning
endif

*-- EOP WRITEOPTADD