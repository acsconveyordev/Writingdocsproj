* Name......... WRITE / edit the GENERAL NOTES program
* Date......... 02/04/2014
* Caller....... writeppdoc.prg
* Notes........ Using the tables
*               determine which general notes will be edited or removed
*               Added a new note 5
*               Prices quoted do not include fork truck for use during off loading and installation.
*      12/08/04 Fork truck note was added to the template thus + 1 to Paragraphs()
*      01/12/09 Paragraph count reduced by 2 when SOO page removed
*      01/27/09 Added the modification of Note 6 for Rock Tenn
*      04/22/10 Paragraph count increased by 6 when training added
*      02/04/14 Revised Note number 8 for Menasha

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

* the group table is in work area a
go top in a
* variable to count the number of groups
PUBLIC nnumofgroups
store 0 to nnumofgroups
do while a.group # '$'
	store nnumofgroups + 1 to nnumofgroups
	skip 1 in a
enddo
* ? nnumofgroups

* variable for keeping track of the paragraph number store 129 to ngnpcount
* training will be plus 9
store 129 to ngnpcount

* As of 01/27/2009 the Rock Tenn Specifications require us to edit note 6
* note 6
*   Prices quoted do not include Fork Truck(s), Cherry Picker(s), Man Lift(s), or Scissors Lift(s) 
* for use during off loading and/or installation. "ACS" can provide any of the above equipment upon
* request and invoice at cost plus 15%.
if ncompnumber = 879
* was 899
	WordApp.ActiveDocument.Paragraphs(ngnpcount-10).Range.Select
	* find the word this and insert the variable ctitle thereafter in paragraph 22
	WordApp.Selection.Find.ClearFormatting
    With WordApp.Selection.Find
        .Text = "can provide any of the above equipment upon request and invoice at cost plus 15%"
        .Replacement.Text = "will provide any of the above equipment as needed and invoice at cost"
        .Forward = .t.
        .Wrap = wdFindContinue
        .Format = .f.
        .MatchCase = .f.
        .MatchWholeWord = .t.
        .MatchWildcards = .f.
        .MatchSoundsLike = .f.
        .MatchAllWordForms = .f.
        .Execute(,,,,,,,,,.replacement.text,.t.)
    EndWith
endif

* As of 02/04/2014 the Menasha Specifications require us to edit note 8
* note 8

if ncompnumber = 297
	WordApp.ActiveDocument.Paragraphs(ngnpcount-8).Range.Select
	* find the word to Buyer and insert the new text
	WordApp.Selection.Find.ClearFormatting
    With WordApp.Selection.Find
        .Text = "to Buyer or any other party"
        .Replacement.Text = "under this contract to Buyer or any other party"
        .Forward = .t.
        .Wrap = wdFindContinue
        .Format = .f.
        .MatchCase = .f.
        .MatchWholeWord = .t.
        .MatchWildcards = .f.
        .MatchSoundsLike = .f.
        .MatchAllWordForms = .f.
        .Execute(,,,,,,,,,.replacement.text,.t.)
    EndWith	
	WordApp.ActiveDocument.Paragraphs(ngnpcount-8).Range.Select
	* find the word to Buyer or any other party and insert the new text
	WordApp.Selection.Find.ClearFormatting
    With WordApp.Selection.Find
        .Text = ", from any cause whatsoever, whether in contract or in tort, including negligence."
        .Replacement.Text = "."
        .Forward = .t.
        .Wrap = wdFindContinue
        .Format = .f.
        .MatchCase = .f.
        .MatchWholeWord = .t.
        .MatchWildcards = .f.
        .MatchSoundsLike = .f.
        .MatchAllWordForms = .f.
        .Execute(,,,,,,,,,.replacement.text,.t.)
    EndWith	
endif

* at this point the table is on the total record
* go back to (General Notes)
* remove note 14 if there is no installation and no r's prices
*  Automated Conveyor Systems, Inc. will provide the mechanical and electrical
*  installation of all items including installation materials and labor (lodging and
*  transportation). This proposal does not include the removal, handling or modifications
*  of any existing equipment or the installation of any equipment other than that
*  specifically stated in this proposal. For further clarification of this installation
*  proposal, please refer to "ACS" Turnkey Installation specifications.
* note 14 is paragragh 129
if a.install = 0 .and. a.rtotal = 0
	WordApp.ActiveDocument.Paragraphs(ngnpcount).Range.Select
	WordApp.Selection.Delete
	* remove one from the count
	ngnpcount = ngnpcount - 1
endif
* remove note 15 if there is no installation and no r's prices
*  If all groups are not purchased and installed at the same time, prices quoted
*  must be re-evaluated.
* note 15 is paragragh 130
* if there is only 1 group
if nnumofgroups > 1
	if a.install = 0 .and. a.rtotal = 0
		* thus remove note 13
		WordApp.ActiveDocument.Paragraphs(ngnpcount+ 1).Range.Select
		WordApp.Selection.Delete
		* remove one from the count
		ngnpcount = ngnpcount - 1
	else
		* leave note in
	endif
else
	* thus remove note
	WordApp.ActiveDocument.Paragraphs(ngnpcount+ 1).Range.Select
	WordApp.Selection.Delete
	* remove one from the count
	ngnpcount = ngnpcount - 1
endif
* remove note 16 if there is no rework prices
*  Rework does not include refurbishment of existing unless noted otherwise on the
*  equipment list. It is assumed that all existing conveyor and components are in workable
*  condition. If it is determined at the time of Automated Conveyor Systems, Inc.
*  installation that existing equipment is in disrepair, any repairs necessary can be
*  performed on a time and material basis.
* note 16 is paragragh 131
if a.rework > 0
	* leave note in
else
	* thus remove note
	WordApp.ActiveDocument.Paragraphs(ngnpcount+2).Range.Select
	WordApp.Selection.Delete
	* remove one from the count
	ngnpcount = ngnpcount - 1
endif

* note 17
*  Installation, relocation, rework, and removal are not included in this proposal.
*  case statement to determine how this should read.
* note 17 is paragragh 132
do case
	* fours
	case a.install = 0 .and. a.relocate = 0 .and. a.rework = 0 .and. a.remove = 0
		* leave as id
	case a.install <> 0 .and. a.relocate <> 0 .and. a.rework <> 0 .and. a.remove <> 0
		* take out
		* remove note if there is install and all the r's prices
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		* remove one from the count
		ngnpcount = ngnpcount - 1
	* ones
	case a.install = 0 .and. a.relocate <> 0 .and. a.rework <> 0 .and. a.remove <> 0
		* Installation is not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Installation is not included in this proposal.")
	case a.install <> 0 .and. a.relocate = 0 .and. a.rework <> 0 .and. a.remove <> 0
		* Relocation is not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Relocation is not included in this proposal.")
	case a.install <> 0 .and. a.relocate <> 0 .and. a.rework = 0 .and. a.remove <> 0
		* Rework is not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Rework is not included in this proposal.")
	case a.install <> 0 .and. a.relocate <> 0 .and. a.rework <> 0 .and. a.remove = 0
		* Removal is not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Removal is not included in this proposal.")
	* twos
	case a.install = 0 .and. a.relocate = 0 .and. a.rework <> 0 .and. a.remove <> 0
		* Installation and relocation are not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Installation and relocation are not included in this proposal.")
	case a.install = 0 .and. a.relocate <> 0 .and. a.rework = 0 .and. a.remove <> 0
		* Installation and rework are not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Installation and rework are not included in this proposal.")
	case a.install = 0 .and. a.relocate <> 0 .and. a.rework <> 0 .and. a.remove = 0
		* Installation and removal are not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Installation and removal are not included in this proposal.")
	case a.install <> 0 .and. a.relocate = 0 .and. a.rework = 0 .and. a.remove <> 0
		* Relocation and rework are not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Relocation and rework are not included in this proposal.")
	case a.install <> 0 .and. a.relocate = 0 .and. a.rework <> 0 .and. a.remove = 0
		* Relocation and removal are not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Relocation and removal are not included in this proposal.")
	case a.install <> 0 .and. a.relocate <> 0 .and. a.rework = 0 .and. a.remove = 0
		* Rework and removal are not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Rework and removal are not included in this proposal.")
	* threes
	case a.install = 0 .and. a.relocate = 0 .and. a.rework = 0 .and. a.remove <> 0
		* Installation, relocation, and rework are not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Installation, relocation, and rework are not included in this proposal.")
	case a.install = 0 .and. a.relocate = 0 .and. a.rework <> 0 .and. a.remove = 0
		* Installation, relocation, and removal are not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Installation, relocation, and removal are not included in this proposal.")
	case a.install = 0 .and. a.relocate <> 0 .and. a.rework = 0 .and. a.remove = 0
		* Installation, rework, and removal are not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Installation, rework, and removal are not included in this proposal.")
	case a.install <> 0 .and. a.relocate = 0 .and. a.rework = 0 .and. a.remove = 0
		* Relocation, rework, and removal are not included in this proposal.
		* edit the Note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+3).Range.Select
		WordApp.Selection.Delete
		WordApp.Selection.TypeParagraph
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.InsertBefore("Relocation, rework, and removal are not included in this proposal.")
endcase

* note 18
*  An Automated Conveyor Systems, Inc. "Installation Services Sheet" can be provided
*  upon request for a supervised installation. If this is selected, installation materials
*  are not included. However, a recommended list will be provided for your use.
* note 18 is paragragh 133
if a.install = 0 .and. a.relocate = 0 .and. a.rework = 0 .and. a.remove = 0
	* leave in
else
	* take out
	WordApp.ActiveDocument.Paragraphs(ngnpcount+4).Range.Select
	WordApp.Selection.Delete
	* remove one from the count
	ngnpcount = ngnpcount - 1
endif
* note 19 and 21
*  All new conveyor drive cabinets to have OSHA Zero Energy Safety Lockout Packages.
*  Bundle Conveyor will be painted "ACSI" Standard Vista Green unless otherwise instructed. An upcharge of ___________ and additional lead time will apply for any color other than standard.
* note 19 is paragragh 134
* note 21 is paragragh 136
if file(ctabledir + cdrawing + 'P.DBF') = .t.
	use ctabledir + cdrawing + 'P.DBF' in h
	select h
	* must be on the C_PLAN__ or C_PLAN_C layers
	* manual entry items as new panels
	* drive panels = substr(g.part,5,5) = "DR-CP"
	* oem strapper panels = substr(g.part,5,5) = "OEMCP"
	* heavy duty centering devices have their own CP thus substr(g.part,5,3) = "LCD"
	* right angle ejectors have their own CP thus substr(g.part,5,5) = "RTANG"
	set filter to (h.layer = "C_PLAN__" .or. h.layer = "C_PLAN_C") .and. ;
				  (h.part = "S-MANUAL_CP--00" .or. ;
	               substr(h.part,5,5) = "DR-CP" .or. ;
	               substr(h.part,5,5) = "OEMCP" .or. ;
	               substr(h.part,5,3) = "LCD" .or. ;
	               substr(h.part,5,5) = "RTANG")
	go top
	if eof("h") = .t.
		store .t. to lremove19
	else
		store .f. to lremove19
	endif
	if lremove19 = .t.
		* thus remove note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+5).Range.Select
		WordApp.Selection.Delete
		* remove one from the count
		ngnpcount = ngnpcount - 1
	else
		* leave in
	endif
	* reset filter before checking for note 21
	set filter to
	go top
	set filter to (h.layer = "C_PLAN__" .or. h.layer = "C_PLAN_C") .and. ;
				  (h.part = "S-MANUAL_CP--00" .or. ;
	               substr(h.part,5,2) = "BC")
	go top
	if eof("h") = .t.
		store .t. to lremove21
	else
		store .f. to lremove21
	endif
	if lremove21 = .f.
		* thus remove note
		WordApp.ActiveDocument.Paragraphs(ngnpcount+7).Range.Select
		WordApp.Selection.Delete
		* remove one from the count
		ngnpcount = ngnpcount - 1
	else
		* leave in
	endif

	* clean up
	use in h
else
	* leave in
	set message to 'The pricing table drawing number + P.DBF is missing. Thus please check notes 17 and 19 on the General Notes page.'
	wait 'The pricing table drawing number + P.DBF is missing. Thus please check notes 17 and 19 on the General Notes page.' window at 5,10 timeout 5
endif

* note 22 and 23
*  Buyer to provide Automated Conveyor Systems, Inc. with tariff or remission number
*  in order to get shipments across border without custom duty charge.
*  Prices quoted are in U.S. Dollars. U.S. Dollars to be paid in par.
* note 22 is paragragh 137
* note 23 is paragragh 138
if ldeletedollars = .t.
	WordApp.ActiveDocument.Paragraphs(ngnpcount+8).Range.Select
	WordApp.Selection.Delete
	* remove one from the count
	ngnpcount = ngnpcount - 1
	WordApp.ActiveDocument.Paragraphs(ngnpcount+9).Range.Select
	WordApp.Selection.Delete
	* remove one from the count
	ngnpcount = ngnpcount - 1
else
	* leave in
endif

* clean up
set message to
* return to a
select a
go top

*-- EOP WRITEGENERAL NOTES

*!*	GENERAL NOTES as of 09/28/07
*!*	1. 	3All invoices not paid within the specified terms of this contract will be subject to a maximum of 1 1/2% per month finance charge. In the event of payment default buyer will be responsible for any and all collection costs including any attorney fees.
*!*	2.	11This proposal refers to the cost of itemized equipment only. Any other existing equipment not itemized can be installed, rebuilt, relocated, removed or reworked by Automated Conveyor Systems, Inc., with prior approval from "ACS" Customer Service Manager, strictly on a time and material basis only.
*!*	3.	  Any Order held at customer's request after notification of readiness to ship will incur a storage fee of 1% for every month held.
*!*	4.	12Removal of all surplus existing equipment is to be performed by others prior to "ACS" installation.
*!*	5.	13All required concrete work by others.
*!*	6.	  Prices quoted do not include Fork Truck(s), Cherry Picker(s), Man Lift(s), or Scissors Lift(s) for use during off loading and/or installation. "ACS" can provide any of the above equipment upon request and invoice at cost plus 15%.
*!*	7.	15If ceiling hung or infloor conduit is desired, it can be quoted as a separate item upon request.
*!*	8.	16Except for the remedy set forth in the Limited Warranty, Automated Conveyor Systems, Inc., shall have no liability to Buyer or any other party for any loss, damage or injury to person or property, from any cause whatsoever, whether in contract or in tort, including negligence. Automated Conveyor Systems, Inc., will not be liable to Buyer or to any other party for consequential or incidental damages from whatever cause, including but not limited to, extra costs or lost profits from loss of productivity or downtime, even if Automated Conveyor Systems, Inc., has been advised of the possibility of such damages.
*!*	9.	17The prices quoted are on today's market conditions. The prices quoted herein are good for 30 days, except that if any unforeseen increases in material or labor cost occur prior to Buyer's Acceptance, these prices are subject to revision.
*!*	10.	18Acceptance of this proposal must be limited to the terms and conditions set forth herein. Any additional or different terms, written or oral, proposed by Buyer are automatically excluded unless specifically accepted by Automated Conveyor Systems, Inc., in writing.
*!*	11.	21Acceptance of the equipment supplied by Automated Conveyor Systems, Inc., shall automatically be deemed to occur as follows:
*!*	 (a)	If Automated Conveyor Systems, Inc., is providing Turnkey Installation, then Buyer shall automatically be deemed to have irrevocably accepted the goods upon the expiration of 30 days following equipment start up, unless Automated Conveyor Systems, Inc., shall have received written notification to the contrary with particularized details of any alleged nonconformity.
*!*	 (b)	In all other cases Buyer shall automatically be deemed to have irrevocably accepted the goods upon the expiration of 10 days after delivery of equipment, unless Automated Conveyor Systems, Inc., shall have received written notification to the contrary with particularized details of any alleged non-conformity.
*!*	12.	1Machine templates are a generic representation of O.E.M. equipment and Automated Conveyor Systems, Inc. can not assume responsibility for their accuracy.
*!*	13.	2Equipment specification sheets can be supplied upon request.
*!*	14.	10Automated Conveyor Systems, Inc. will provide the mechanical and electrical installation of all items including installation materials and labor (lodging and transportation). This proposal does not include the removal, handling or modifications of any existing equipment or the installation of any equipment other than that specifically stated in this proposal. For further clarification of this installation proposal, please refer to "ACS" Turnkey Installation specifications.
*!*	15.	5If all groups are not purchased and installed at the same time, prices quoted must be re-evaluated.
*!*	16.	9Rework does not include refurbishment of existing unless noted otherwise on the equipment list. It is assumed that all existing conveyor and components are in workable condition. If it is determined at the time of Automated Conveyor Systems, Inc. installation that existing equipment is in disrepair, any repairs necessary can be performed on a time and material basis.
*!*	17.	6Installation, relocation, rework, and removal are not included in this proposal.
*!*	18.	4An Automated Conveyor Systems, Inc. "Installation Services Sheet" can be provided upon request for a supervised installation. If this is selected, installation materials are not included. However, a recommended list will be provided for your use.
*!*	19.	19All new conveyor drive cabinets to have OSHA Zero Energy Safety Lockout Packages.
*!*	20.	20All Control Panels will be built to UL specifications. All components used are UL approved.  
*!*	21.	14Bundle Conveyor will be painted "ACSI" Standard Vista Green unless otherwise instructed. An upcharge of ___________ and additional lead time will apply for any color other than standard.
*!*	22.	7Buyer to provide Automated Conveyor Systems, Inc. with tariff or remission number in order to get shipments across border without custom duty charge.
*!*	23.	8Prices quoted are in U.S. Dollars. U.S. Dollars to be paid in par.
