* Name......... WRITE the ADDRESS LABELS from foxpro into word program
* Date......... 07/31/2003
* Caller....... writeppdoc.prg
* Notes........ This prints the Address Labels if required.

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

* Print the Address Label
******
* Create Application & Document
WordApp = CreateObject("Word.application.8")
*!*	 * Open document template
*!*	WordDoc = WordApp.Documents.open('c:\my documents\AddressLabelspp.dot')
* Open edits the template and does not allow saving as a .doc file thus
* Add a new document using the AddressLabelspp template
*                          .Add(Template:                   , NewTemplate)
* WordDoc = WordApp.Documents.Add('c:\my documents\Dot2.dot', .f.)
* When the network template is used the above statement will change to
WordDoc = WordApp.Documents.Add('G:\Word Templates\Quote Templates\AddressLabelswd.dot', .f.)

* New caption
WordApp.Caption = "ACS Address Labels"
* Make Visible
WordApp.Application.Visible = .t.
* start at paragraph 5 and insert information
WordApp.ActiveDocument.Paragraphs(5).Range.Select
WordApp.Selection.StartOf
WordApp.Selection.MoveRight(wdCharacter,5)
WordApp.Selection.TypeText(cppsal+" "+cppnamef+" "+cppnamel+(chr(11)))
WordApp.Selection.TypeText(" "+cppcustomer+(chr(11)))
WordApp.Selection.TypeText(" "+cppaddress+(chr(11)))
WordApp.Selection.TypeText(" "+cpplocation + " " + cppzip)

* a copy goes to
* insert the copies for names
if lgotcopypernumber = .t.
	if len(tempselectarray[1,1]) > 0
		* Select and insert name(s) that will get a copy
		* print their names - limit set at 10
		* this will created one whole page of labels of 6
		* and one page of 4
		for f = 1 to 5
			* variable npc is numerical paragraph count
			do case
			case f = 1
				npc = 7
			case f = 2
				npc = 13
			case f = 3
				npc = 15
			case f = 4
				npc = 21
			case f = 5
				npc = 23
			endcase
			* if there is a number in tempselectarray[f,1] then print
			* else delete the TO:
			if len(tempselectarray[f,1]) > 0
				* select paragraph and position
				WordApp.ActiveDocument.Paragraphs(npc).Range.Select
				WordApp.Selection.StartOf
				WordApp.Selection.MoveRight(wdCharacter,5)
				* type name
				WordApp.Selection.TypeText(tempselectarray[f,2]+" "+tempselectarray[f,3]+" "+tempselectarray[f,4]+(chr(11)))
				* type company
				WordApp.Selection.TypeText(" "+tempselectarray[f,6]+(chr(11)))
				* type street address
				WordApp.Selection.TypeText(" "+tempselectarray[f,7]+(chr(11)))
				* type street address
				WordApp.Selection.TypeText(" "+tempselectarray[f,8])
			else
				* remove the TO: from the template if not applicable
				WordApp.ActiveDocument.Paragraphs(npc).Range.Select
				WordApp.Selection.Delete
			endif
		endfor
		if len(tempselectarray[6,1]) > 0
			for f = 6 to 11
				* if there are more than five you need a new page
				if f = 6
					* Create Application & Document
					WordApp = CreateObject("Word.application.8")
					* Add a new document using the AddressLabelspp template
					*                          .Add(Template:                   , NewTemplate)
					* WordDoc = WordApp.Documents.Add('c:\my documents\Dot2.dot', .f.)
					* When the network template is used the above statement will change to
					WordDoc = WordApp.Documents.Add('G:\Word Templates\Quote Templates\AddressLabelswd.dot', .f.)
					* New caption
					WordApp.Caption = "ACS Address Labels Page 2"
					* Make Visible
					WordApp.Application.Visible = .t.
				endif
				* variable npc is numerical paragraph count
				do case
				case f = 6
					npc = 5
				case f = 7
					npc = 7
				case f = 8
					npc = 13
				case f = 9
					npc = 15
				case f = 10
					npc = 21
				case f = 11
					npc = 23
				endcase
				* if there is a number in tempselectarray[f,1] then print
				* else delete the TO:
				if len(tempselectarray[f,1]) > 0
					* select paragraph and position
					WordApp.ActiveDocument.Paragraphs(npc).Range.Select
					WordApp.Selection.StartOf
					WordApp.Selection.MoveRight(wdCharacter,5)
					* type name
					WordApp.Selection.TypeText(tempselectarray[f,2]+" "+tempselectarray[f,3]+" "+tempselectarray[f,4]+(chr(11)))
					* type company
					WordApp.Selection.TypeText(" "+tempselectarray[f,6]+(chr(11)))
					* type street address
					WordApp.Selection.TypeText(" "+tempselectarray[f,7]+(chr(11)))
					* type street address
					WordApp.Selection.TypeText(" "+tempselectarray[f,8])
				else
					* remove the TO: from the template if not applicable
					WordApp.ActiveDocument.Paragraphs(npc).Range.Select
					WordApp.Selection.Delete
				endif
			endfor
		endif
	endif
else
	* remove the TO: from the template if not applicable
	WordApp.ActiveDocument.Paragraphs(23).Range.Select
	WordApp.Selection.Delete
	WordApp.ActiveDocument.Paragraphs(21).Range.Select
	WordApp.Selection.Delete
	WordApp.ActiveDocument.Paragraphs(15).Range.Select
	WordApp.Selection.Delete
	WordApp.ActiveDocument.Paragraphs(13).Range.Select
	WordApp.Selection.Delete
	WordApp.ActiveDocument.Paragraphs(7).Range.Select
	WordApp.Selection.Delete
endif
WordApp.ActiveDocument.Paragraphs(1).Range.Select
WordApp.Selection.Collapse

*-- EOP WRITEADDRESSLABELS