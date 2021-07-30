* Name......... WRITE the Price Page DOCument from foxpro into word program
* Date......... 05/29/2012
* Caller....... writingdocs_app.prg
* Notes........ This is the second attempt at using Visual Basic language to create a WORD document.
*               chr(11) = Shift Enter  chr(9) = Tab
*               WordApp.Selection.TypeParagraph = enter
*               Added the System Engineer to the Price Page
*               The mail to information corrected.
*               Changed System Engineer to PLC Start Up
*               Added if statement for sold at price page ('S.DBF')
*               Remove the Print the Address Label
*               Remove the Print the Quote Labels
*               Remove the Price Page Programming from this program.
*      12/08/04 Fork truck note was added to the template thus + 1 to Paragraphs() after 107
*      04/29/05 Added a ten per cent discount for Southern Container
*      05/23/07 Changed Smurfit Price Page to show no discount
*      06/15/07 Changed Smurfit Price Page to show no discount like IP
*      01/12/09 Remove the Sequence Of Operations page
*               Changed IP less IP Ireland pricepagemodnetonly to do pricepagemodshown
*      11/05/10 Changed 1089 to pricepagemodshown 
*               Add the Training Price

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

if lcontinuewriting = .t.
	* the form cformspath + 'selectcopytoname' has a cancel button on it
	* if it is pressed lstoprunning will be set to true (.t.)
	if lstoprunning = .f.
		******
		* Create Application & Document
		WordApp = CreateObject("Word.Application.8")
		*!*	 * Open document template
		*!*	WordDoc = WordApp.Documents.Open('c:\my documents\quote2.dot')
		* Open edits the template and does not allow saving as a .doc file thus
		* Add a new document using the quote2 template
		*                          .Add(Template:                   , NewTemplate)
		* WordDoc = WordApp.Documents.Add('c:\my documents\quote2.dot', .f.)
		if lonnetwork = .f.
			* When on the network this template will be used
			WordDoc = WordApp.Documents.Add('G:\Word Templates\Quote Templates\quote31.dot', .f.)
			* WordDoc = WordApp.Documents.Add('C:\Program Files\Microsoft Visual Studio\Writingdocsproj\INCLUDE\quote31.dot', .f.)
		else
			* When not on the network this template will be used
			WordDoc = WordApp.Documents.Add(cincludepath + 'quote31.dot', .f.)
		endif

		* New caption
		WordApp.Caption = "Quotation Price Page Documents"
		* Make Visible
		WordApp.Application.Visible = .t.

		******** Page 1   The Fax Cover Sheet
		* lforavanti = Avanti Conveyors ldofaxsheet is for Fax Sheet Y or N
		if lforavanti = .t. .or. ldofaxsheet = .f.
			* the actual page is deleted at the end so as to not corrupt the paragraph count
			* no fax page required
		else
			* Insert the customers complete name in paragraph 5
			if len(cppsal+cppnamef+cppnamel) > 0    && name selected
				WordApp.ActiveDocument.Paragraphs(5).Range.Select
				WordApp.Selection.StartOf
				WordApp.Selection.MoveRight(wdCharacter,18)
				WordApp.Selection.Delete(wdCharacter,1)
				WordApp.Selection.InsertAfter(cppsal+" "+cppnamef+" "+cppnamel)
			else
				* do nothing
			endif
			* Insert mdy(date()) in paragraph 8
			WordApp.ActiveDocument.Paragraphs(8).Range.Select
			WordApp.Selection.StartOf
			WordApp.Selection.Delete(wdCharacter,5)
			WordApp.Selection.InsertBefore(mdy(date()))
			WordApp.Selection.Font.Bold = .f.
			WordApp.Selection.Font.Size = 11
			* Insert selected name in paragraph 9 if one was selected
			if len(cppsal+ cppnamef+ cppnamel) > 0    && name selected
				WordApp.ActiveDocument.Paragraphs(9).Range.Select
				WordApp.Selection.StartOf
				if cppsal = "Mr."
					WordApp.Selection.MoveRight(wdCharacter,8)
					WordApp.Selection.InsertAfter(cppnamef+" "+cppnamel)
				else
					WordApp.Selection.MoveRight(wdCharacter,4)
					WordApp.Selection.Delete(wdCharacter,4)
					WordApp.Selection.InsertAfter(cppsal+" "+cppnamef+" "+cppnamel)
				endif
			else
				* do nothing
			endif
			******
			* Insert variable cssname in paragraph 10
			WordApp.ActiveDocument.Paragraphs(10).Range.Select
			WordApp.Selection.StartOf
			WordApp.Selection.MoveRight(wdCharacter,6)
			WordApp.Selection.InsertBefore(cssname)
			******
			******
			* Insert customer plant name and location
			* go to Paragraph 12 and insert/edit information
			WordApp.ActiveDocument.Paragraphs(12).Range.Select
			WordApp.Selection.StartOf
			WordApp.Selection.MoveRight(wdCharacter,1)
			WordApp.Selection.InsertAfter(ccustomer)
			WordApp.Selection.MoveDown(wdLine,1)
			WordApp.Selection.InsertAfter(clocation)
			******
			******
			* Insert customer plant phone number
			* go to Paragraph 15 and insert/edit information
			WordApp.ActiveDocument.Paragraphs(15).Range.Select
			WordApp.Selection.StartOf
			WordApp.Selection.MoveRight(wdCharacter,11)
			WordApp.Selection.InsertAfter(cppphone)
			******
			******
			* Insert customer plant fax number
			* go to Paragraph 17 and insert/edit information
			WordApp.ActiveDocument.Paragraphs(17).Range.Select
			WordApp.Selection.StartOf
			WordApp.Selection.MoveRight(wdCharacter,11)
			WordApp.Selection.InsertAfter(cppfax)
			******
			******
			* Insert selected name here if one was selected in paragraph 20
			if len(cppsal+cppnamef+cppnamel) > 0    && name selected
				WordApp.ActiveDocument.Paragraphs(20).Range.Select
				WordApp.Selection.StartOf
				if cppsal = "Mr."
					WordApp.Selection.MoveRight(wdCharacter,11)
					WordApp.Selection.InsertAfter(cppnamel)
				else
					WordApp.Selection.MoveRight(wdCharacter,7)
					WordApp.Selection.Delete(wdCharacter,4)
					WordApp.Selection.InsertAfter(cppsal+" "+cppnamel)
				endif
			else
				* do nothing
			endif
			******
			******
			* Enter the quotation title in the sentence
			* We are pleased to submit this "quotation title" proposal for your review.
			* find the word this and insert the variable ctitle thereafter in paragraph 22
			WordApp.ActiveDocument.Paragraphs(22).Range.Select
			WordApp.Selection.Find.ClearFormatting
			With WordApp.Selection.Find
				.Text = "this "
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
			* then enter the layout title
			WordApp.Selection.InsertAfter(ctitle + " ")
			* Then change the font of ctitle
			WordApp.ActiveDocument.Paragraphs(22).Range.Select
			WordApp.Selection.Find.ClearFormatting
			With WordApp.Selection.Find
				.Text = ctitle
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
			WordApp.Selection.Font.Bold = .t.
			WordApp.Selection.Font.Italic = .t.
			* add the city state plus the word plant to this line
			if nmailto # ncompleteacsno
				WordApp.ActiveDocument.Paragraphs(22).Range.Select
				WordApp.Selection.Find.ClearFormatting
				With WordApp.Selection.Find
					.Text = "for "
					.Forward = .t.
					.Wrap = wdFindContinue
					.Format = .f.
					.MatchCase = .t.
					.MatchWholeWord = .t.
					.MatchWildcards = .f.
					.MatchSoundsLike = .f.
					.MatchAllWordForms = .f.
				EndWith
				store clocation + " Plant" to cdiffloc
				WordApp.Selection.Find.Execute
				* then enter the different plants location variable (cdiffloc)
				WordApp.Selection.InsertAfter("the " + cdiffloc + " for ")
				* Then change the font of ctitle
				WordApp.ActiveDocument.Paragraphs(22).Range.Select
				WordApp.Selection.Find.ClearFormatting
				With WordApp.Selection.Find
					.Text = cdiffloc
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
				WordApp.Selection.Font.Bold = .t.
				WordApp.Selection.Font.Italic = .t.
			endif

			WordApp.ActiveDocument.Paragraphs(23).Range.Select
			WordApp.Selection.Find.ClearFormatting
			With WordApp.Selection.Find
				.Text = ", ,"
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
			WordApp.Selection.MoveLeft(wdCharacter,1)
			WordApp.Selection.MoveRight(wdCharacter,2)
			* Then enter the salesmans' name.
			WordApp.Selection.InsertBefore(alltrim(cstname))
			WordApp.Selection.Font.Italic = .t.

			******
*!*				* Place sales support name in paragraph 29 above Sales Support Engineer
*!*				* As of 09/25/2003 only the Sales Support Manager will be used
*!*				* unless it is in Mexico then it will be Juan Ramos.
*!*				WordApp.ActiveDocument.Paragraphs(29).Range.Select
*!*				WordApp.Selection.StartOf
*!*				if linmex = .t.
*!*					* insert Juan Ramos
*!*					WordApp.Selection.InsertBefore("Juan Ramos")
*!*					WordApp.Selection.MoveDown(wdLine,3)
*!*				else
*!*					* insert Eddie Stanley
*!*					WordApp.Selection.InsertBefore("Eddie Stanley")
*!*					WordApp.Selection.MoveDown(wdLine,1)
*!*					WordApp.Selection.EndKey(wdLine,1)
*!*					WordApp.Selection.InsertBefore(" Manager")
*!*					WordApp.Selection.MoveDown(wdLine,2)
*!*				endif

			WordApp.ActiveDocument.Paragraphs(29).Range.Select
			WordApp.Selection.StartOf
			* insert Regional Sales Manager
			WordApp.Selection.InsertBefore(cstname)
			* reselect to replace Sales Support with Regional Sales Manager
			WordApp.ActiveDocument.Paragraphs(29).Range.Select
		    WordApp.Selection.Find.ClearFormatting
		    WordApp.Selection.Find.Replacement.ClearFormatting
		    With WordApp.Selection.Find
		        .Text = "Sales Support"
		        .Replacement.Text = "Regional Sales Manager"
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
			WordApp.Selection.MoveDown(wdLine,2)

			if lonnetwork = .t.
				* if cuser = "sf"    && replace mm with cuser
					WordApp.Selection.TypeBackspace
					WordApp.Selection.TypeBackspace
					WordApp.Selection.InsertAfter(" " + cuser)
					WordApp.Selection.MoveRight(wdCharacter,1)
				* endif
			else
				* replace mm with csti
				WordApp.Selection.TypeBackspace
				WordApp.Selection.TypeBackspace
				WordApp.Selection.TypeBackspace
				* WordApp.Selection.InsertAfter(cuser)
				WordApp.Selection.InsertAfter(" " + csti)
				WordApp.Selection.MoveRight(wdCharacter,1)
			endif
			* Add a space by Mary Wade
			* WordApp.Selection.MoveLeft(wdCharacter,3)
			WordApp.Selection.MoveLeft(wdCharacter,4)
			* Place sales support initials
			WordApp.Selection.InsertBefore(upper(cssi))
			* Select paragraph and add quotation number
			WordApp.ActiveDocument.Paragraphs(29).Range.Select
			WordApp.Selection.Find.ClearFormatting
			With WordApp.Selection.Find
				.Text = "03-"
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
			if len(alltrim(cquote)) > 0
				* erases the template information 03-
				WordApp.Selection.Delete
				* insert quotation number
				WordApp.Selection.InsertAfter(cquote)
			endif
		endif   &&    if lforavanti = .t.
		******
		******** This completes the first page

*!*			Removed 04/18/2003
*!*			******** Old Page 2
*!*			* was the Cover Sheet
*!*			* Go to Paragraph 33 and insert/edit information
*!*			WordApp.ActiveDocument.Paragraphs(33).Range.Select
*!*			WordApp.Selection.StartOf
*!*			if len(alltrim(ccustomer)) > 0
*!*				WordApp.Selection.InsertAfter(ccustomer)
*!*			endif
*!*			WordApp.Selection.MoveDown(wdLine,1)
*!*			if len(alltrim(clocation)) > 0
*!*				WordApp.Selection.InsertAfter(clocation)
*!*			endif
*!*			* Select paragraph and add quotation number
*!*			WordApp.ActiveDocument.Paragraphs(33).Range.Select
*!*			WordApp.Selection.Find.ClearFormatting
*!*			With WordApp.Selection.Find
*!*				.Text = "03-"
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.Selection.Find.Execute
*!*			if len(alltrim(cquote)) > 0
*!*				* erases the template information 03-
*!*				WordApp.Selection.Delete
*!*				* insert quotation number
*!*				WordApp.Selection.InsertAfter(cquote)
*!*			endif
*!*			WordApp.Selection.MoveDown(wdLine,1)
*!*			WordApp.Selection.TypeBackspace
*!*			WordApp.Selection.TypeBackspace
*!*			WordApp.Selection.InsertAfter(ctitle + '"' + chr(11))
*!*			WordApp.Selection.MoveDown(wdLine,1)
*!*			* inserts month, day and year
*!*			WordApp.Selection.InsertBefore(mdy(date()))
*!*			******
*!*			******** This completes the second page

		******** Page 2
		******
		* Insert mdy(date()) in paragraph(30)
		WordApp.ActiveDocument.Paragraphs(30).Range.Select
		WordApp.Selection.Collapse(wdCollapseStart)
		*WordApp.Selection.Collapse(1)
		WordApp.Selection.MoveDown(wdLine,15)
		WordApp.Selection.Delete(wdCharacter,6)
		WordApp.Selection.InsertAfter(mdy(date()))
		******
		******
		* Move down 3 lines then insert customer name
		WordApp.Selection.MoveDown(wdLine,3)
		if len(cppsal+cppnamef+cppnamel) > 0    && name selected
			WordApp.Selection.InsertAfter(cppsal+" "+cppnamef+" "+cppnamel)
		else
			* do nothing
		endif
		******
		******
		* Insert company name
		* Move down 1 line then insert company name
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.InsertAfter(cppcustomer)
		* Move down 1 line and check for an address
		WordApp.Selection.MoveDown(wdLine,1)
		if len(alltrim(cppaddress)) > 0    && got an address
			* Insert Address
			WordApp.Selection.InsertBefore(cppaddress)
		endif
		* Check for customer location
		WordApp.Selection.MoveDown(wdLine,1)
		if len(alltrim(cpplocation)) > 0
			WordApp.Selection.StartOf
			WordApp.Selection.InsertBefore(cpplocation + " " + cppzip)
		endif
		******
		******
		* Dear Mr. : ?
		* Last name in paragraph(31)
		WordApp.ActiveDocument.Paragraphs(31).Range.Select
		WordApp.Selection.EndKey
		WordApp.Selection.MoveLeft(wdCharacter,1)
		if len(cppsal+cppnamef+cppnamel) > 0    && name selected
			if cppsal = "Mr."
				WordApp.Selection.InsertAfter(cppnamel)
			else
				WordApp.Selection.MoveLeft(wdCharacter,4)
				WordApp.Selection.Delete(wdCharacter,4)
				WordApp.Selection.InsertAfter(cppsal+" "+cppnamel)
			endif
		else
			* do nothing
		endif
		******
		******
		* Enter the quotation title in the sentence
		* We are pleased to submit this "quotation title" proposal for your review.
		* find the word this and insert the variable ctitle thereafter paragraph 32
		WordApp.ActiveDocument.Paragraphs(32).Range.Select
		WordApp.Selection.Find.ClearFormatting
		With WordApp.Selection.Find
			.Text = "this "
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
		* then enter the layout title
		if len(cquote) > 9    && this is a revision
			WordApp.Selection.InsertAfter("revised " + ctitle + " ")
		else
			WordApp.Selection.InsertAfter(ctitle + " ")
		endif
		* then change font of ctitle
		WordApp.ActiveDocument.Paragraphs(32).Range.Select
		WordApp.Selection.Find.ClearFormatting
		With WordApp.Selection.Find
			.Text = ctitle
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
		WordApp.Selection.Font.Bold = .t.
		WordApp.Selection.Font.Italic = .t.
		* add the city state plus the word plant to this line
		if nmailto # ncompleteacsno
			WordApp.ActiveDocument.Paragraphs(22).Range.Select
			WordApp.Selection.Find.ClearFormatting
			With WordApp.Selection.Find
				.Text = "for "
				.Forward = .t.
				.Wrap = wdFindContinue
				.Format = .f.
				.MatchCase = .t.
				.MatchWholeWord = .t.
				.MatchWildcards = .f.
				.MatchSoundsLike = .f.
				.MatchAllWordForms = .f.
			EndWith
			store clocation + " Plant" to cdiffloc
			WordApp.Selection.Find.Execute
			* then enter the different plants location variable (cdiffloc)
			WordApp.Selection.InsertAfter("the " + cdiffloc + " for ")
			* Then change the font of ctitle
			WordApp.ActiveDocument.Paragraphs(22).Range.Select
			WordApp.Selection.Find.ClearFormatting
			With WordApp.Selection.Find
				.Text = cdiffloc
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
			WordApp.Selection.Font.Bold = .t.
			WordApp.Selection.Font.Italic = .t.
		endif
		******
		******
		* Enter the sales territory in the sentence
		* Your "ACS" representative, "name", look forward to ...
		* find the word this and insert the variable cstname thereafter paragraph 33
		WordApp.ActiveDocument.Paragraphs(33).Range.Select
		WordApp.Selection.Find.ClearFormatting
		With WordApp.Selection.Find
			.Text = ", ,"
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
		WordApp.Selection.MoveLeft(wdCharacter,1)
		WordApp.Selection.MoveRight(wdCharacter,2)
		* Then enter the salesmans' name.
		WordApp.Selection.InsertBefore(alltrim(cstname))
		WordApp.Selection.Font.Italic = .t.

		* As of 08/15/2006 only the Regional Sales Manager will be used
		WordApp.ActiveDocument.Paragraphs(38).Range.Select
		WordApp.Selection.StartOf
		* insert Regional Sales Manager
		WordApp.Selection.InsertBefore(cstname)
		* reselect to replace Sales Support with Regional Sales Manager
		WordApp.ActiveDocument.Paragraphs(38).Range.Select
	    WordApp.Selection.Find.ClearFormatting
	    WordApp.Selection.Find.Replacement.ClearFormatting
	    With WordApp.Selection.Find
	        .Text = "Sales Support"
	        .Replacement.Text = "Regional Sales Manager"
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
		WordApp.Selection.MoveDown(wdLine,2)

		if lonnetwork = .t.
			* if cuser = "sf"    && replace mm with sf
				WordApp.Selection.TypeBackspace
				WordApp.Selection.TypeBackspace
				WordApp.Selection.InsertAfter(" " + cuser)
				WordApp.Selection.MoveRight(wdCharacter,1)
			* endif
		else
			* replace mm with csti    && replace mm with csti
			WordApp.Selection.TypeBackspace
			WordApp.Selection.TypeBackspace
			WordApp.Selection.TypeBackspace
			WordApp.Selection.InsertAfter(" " + csti)
			WordApp.Selection.MoveRight(wdCharacter,1)
		endif
		* Add a space by Mary Wade
		* WordApp.Selection.MoveLeft(wdCharacter,3)
		WordApp.Selection.MoveLeft(wdCharacter,4)
		* Place sales support initials
		WordApp.Selection.InsertBefore(upper(cssi))
		* Select paragraph and add quotation number
		WordApp.ActiveDocument.Paragraphs(38).Range.Select
		WordApp.Selection.Find.ClearFormatting
		With WordApp.Selection.Find
			.Text = "03-"
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
		if len(alltrim(cquote)) > 0
			* erases the template information 03-
			WordApp.Selection.Delete
			* insert quotation number
			WordApp.Selection.InsertAfter(cquote)
		endif
		* insert the copies for names
		if lgotcopypernumber = .t.
			if len(tempselectarray[1,1]) > 0
				* Select paragraph 39 and name(s) that will get a copy
				WordApp.ActiveDocument.Paragraphs(39).Range.Select
				WordApp.Selection.StartOf
				WordApp.Selection.MoveRight(wdCharacter,4)
				* print their names
				for f = 1 to 10
					if len(tempselectarray[f,1]) > 0
						if f = 1
							WordApp.Selection.InsertAfter(tempselectarray[f,2]+" "+tempselectarray[f,3]+" "+tempselectarray[f,4])
						endif
						if f > 1
							* a shift enter plus a tab then the name
							WordApp.Selection.InsertAfter(chr(11)+chr(9))
							WordApp.Selection.InsertAfter(tempselectarray[f,2]+" "+tempselectarray[f,3]+" "+tempselectarray[f,4])
						endif
					else
						* do nothing
					endif
				endfor
			else
				* do nothing
			endif
		endif
		******
		******** This completes the third page

		******** Page 4
		******
		* Paragraph 41 is the Price Page heading
		* go to Paragraph 41 and insert/edit information
		WordApp.ActiveDocument.Paragraphs(41).Range.Select
		WordApp.Selection.StartOf
		WordApp.Selection.MoveDown
		if len(alltrim(ccustomer)) > 0
			WordApp.Selection.InsertAfter(ccustomer)
		endif
		WordApp.Selection.MoveDown(wdLine,1)
		if len(alltrim(clocation)) > 0
			WordApp.Selection.InsertAfter(clocation)
		endif
		if len(alltrim(cquote)) > 0
			WordApp.ActiveDocument.Paragraphs(41).Range.Select
			WordApp.Selection.Find.ClearFormatting
			With WordApp.Selection.Find
				.Text = "03-"
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
			* erases the template information 03-
			WordApp.Selection.Delete
			* insert quotation number
			WordApp.Selection.InsertAfter(cquote)
		endif
		* remove the template information , 2003
		WordApp.ActiveDocument.Paragraphs(41).Range.Select
		WordApp.Selection.Find.ClearFormatting
		With WordApp.Selection.Find
			.Text = ", 2003"
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
		* erases the template information , 2003
		WordApp.Selection.Delete
		WordApp.Selection.InsertBefore(mdy(date()))
		******
		******
		* go to next paragraph
		WordApp.Selection.MoveDown(wdLine,1)
		* if lforavanti = .t.   && Avanti Conveyors
		if lforavanti = .t.
			* find the words and electrical and delete them
			WordApp.Selection.Find.ClearFormatting
			With WordApp.Selection.Find
				.Text = " and electrical"
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
			WordApp.Selection.Delete
		endif
		* find the number symbol(#) and erase
		WordApp.Selection.Find.ClearFormatting
		With WordApp.Selection.Find
			.Text = "#"
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
		WordApp.Selection.TypeBackspace
		* then enter number symbol plus the layout number
		WordApp.Selection.InsertAfter("#" + clayout)
		* then enter the quote title in paragraph 43
		WordApp.ActiveDocument.Paragraphs(43).Range.Select
		WordApp.Selection.InsertBefore(ctitle)
		******
		******** This completes the fourth page for now

		********
		* As of 12/08/04 a paragraph for the Fork Truck was added thus
		* Paragraph must have 1 added

*!*			******** Old Page 2
*!*			* was the Cover Sheet
*!*			Removed 01/12/2009
*!*			******** Old Page 5
*!*			******
*!*			* Paragraph 113 is the Sequence Of Operations Sheet heading
*!*			* go to Paragraph 113 and insert/edit information
*!*			WordApp.ActiveDocument.Paragraphs(113).Range.Select
*!*			WordApp.Selection.StartOf
*!*			WordApp.Selection.MoveDown(wdLine,1)
*!*			if len(alltrim(ccustomer)) > 0
*!*				WordApp.Selection.InsertAfter(ccustomer)
*!*			endif
*!*			WordApp.Selection.MoveDown(wdLine,1)
*!*			if len(alltrim(clocation)) > 0
*!*				WordApp.Selection.InsertAfter(clocation)
*!*			endif
*!*			* if len(alltrim(cquote)) > 0
*!*				WordApp.ActiveDocument.Paragraphs(113).Range.Select
*!*				WordApp.Selection.Find.ClearFormatting
*!*				With WordApp.Selection.Find
*!*					.Text = "03-"
*!*					.Forward = .t.
*!*					.Wrap = wdFindContinue
*!*					.Format = .f.
*!*					.MatchCase = .t.
*!*					.MatchWholeWord = .t.
*!*					.MatchWildcards = .f.
*!*					.MatchSoundsLike = .f.
*!*					.MatchAllWordForms = .f.
*!*				EndWith
*!*				WordApp.Selection.Find.Execute
*!*				* erases the template information 03-
*!*				WordApp.Selection.Delete
*!*				* insert quotation number
*!*				WordApp.Selection.InsertAfter(cquote+chr(11))
*!*			* endif
*!*			WordApp.Selection.MoveDown(wdLine,1)
*!*			* inserts month, day and year
*!*			WordApp.Selection.InsertBefore(mdy(date()))
*!*			******
*!*			******** This completes the fifth page
*!*			* Paragraph must have 2 removed


*!*			******** Old Page 6
*!*			Removed 04/18/2003
*!*			******
*!*			* Do the General Equipment Specifications next
*!*			* Paragraph 117 is the first paragraph on General Equipment Specifications page
*!*			* go to Paragraph 117 and insert/edit information
*!*			* top of roller
*!*			WordApp.ActiveDocument.Paragraphs(117).Range.Select
*!*			WordApp.Selection.EndKey
*!*			WordApp.Selection.MoveLeft(wdCharacter,1)
*!*			WordApp.Selection.InsertAfter(alltrim(str(b.tor)) + '"')
*!*			* roll centers
*!*			WordApp.ActiveDocument.Paragraphs(118).Range.Select
*!*			WordApp.Selection.EndKey
*!*			WordApp.Selection.MoveLeft(wdCharacter,1)
*!*			WordApp.Selection.InsertAfter(alltrim(str(b.rtr)) + '"')
*!*			* conveyor speed
*!*			WordApp.ActiveDocument.Paragraphs(119).Range.Select
*!*			WordApp.Selection.EndKey
*!*			WordApp.Selection.MoveLeft(wdCharacter,1)
*!*			WordApp.Selection.InsertAfter(alltrim(str(b.speed)) + ' FPM')
*!*			* floor conveyor color
*!*			WordApp.ActiveDocument.Paragraphs(120).Range.Select
*!*			WordApp.Selection.EndKey
*!*			WordApp.Selection.MoveLeft(wdCharacter,1)
*!*			WordApp.Selection.InsertAfter("To Be Determined")
*!*			* floor conveyor paint spec. number
*!*			WordApp.ActiveDocument.Paragraphs(121).Range.Select
*!*			WordApp.Selection.EndKey
*!*			WordApp.Selection.MoveLeft(wdCharacter,1)
*!*			WordApp.Selection.InsertAfter("To Be Determined")
*!*			* bundle conveyor color
*!*			WordApp.ActiveDocument.Paragraphs(122).Range.Select
*!*			WordApp.Selection.EndKey
*!*			WordApp.Selection.MoveLeft(wdCharacter,1)
*!*			WordApp.Selection.InsertAfter("To Be Determined" + " (See Note   & Option    )")
*!*			* voltage
*!*			WordApp.ActiveDocument.Paragraphs(123).Range.Select
*!*			WordApp.Selection.EndKey
*!*			WordApp.Selection.MoveLeft(wdCharacter,1)
*!*			WordApp.Selection.InsertAfter(alltrim(b.voltage))
*!*			******
*!*			******** This completes the sixth page

		******** Page 7
		******
		* Do the General Notes heading next because you don't know how many
		*  paragraphs will be added on the price page
		* Paragraph 113 is the General Notes heading
		* go to Paragraph 113 and insert/edit information
		* training will add 9 to this
		WordApp.ActiveDocument.Paragraphs(113).Range.Select
		WordApp.Selection.StartOf
		WordApp.Selection.MoveDown(wdLine,1)
		if len(alltrim(ccustomer)) > 0
			WordApp.Selection.InsertAfter(ccustomer)
		endif
		WordApp.Selection.MoveDown(wdLine,1)
		if len(alltrim(clocation)) > 0
			WordApp.Selection.InsertAfter(clocation)
		endif
		*if len(alltrim(cquote)) > 0
		* go to Paragraph 113 and insert/edit information
		* training will add 9 to this
			WordApp.ActiveDocument.Paragraphs(113).Range.Select
			WordApp.Selection.Find.ClearFormatting
			With WordApp.Selection.Find
				.Text = "03-"
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
			* erases the template information 03-
			WordApp.Selection.Delete
			WordApp.Selection.InsertAfter(cquote)
		*endif
		* inserts month, day and year
		WordApp.Selection.MoveDown(wdLine,1)
		WordApp.Selection.InsertBefore(mdy(date()))
		******
		* editing of the general notes will be done before the Price Page program has run
		******** This temporary completes the seventh page

		******** Page 8
		******
		* This is the Installation Responsibilities page
		* There is nothing on this page to edit or add.
		******
		******** This completes the eighth page

		******** Page 9
		******
		* Paragraph 298 is the Date at the end of the
		* Limited Warranty & Miscellaneous Terms Page
		* before the price page has been run edit the Limited Warranty & Miscellaneous Terms Page
		do writegandt
		* inserts month, day and year
		*WordApp.ActiveDocument.Paragraphs(308+ngtpcount).Range.Select
		WordApp.ActiveDocument.Paragraphs(299+ngtpcount).Range.Select
		WordApp.Selection.EndKey
		WordApp.Selection.InsertAfter(mdy(date()))
		******
		******** This completes the ninth page

		******** Return to the seventh page
		* before the price page has been run edit the General Notes
		do writegeneralnotes
		******** This completes the seventh page

		**************************************
		PUBLIC loptionnewpage
		store .f. to loptionnewpage
		* The price page will now be done to show/don't show discounts thus there will be
		*  a no discount price page
		*  material and installation discount price page
		*  an Avanti Conveyor price page
		*  and a material only not shown discount price page 
		*  and a material only shown discount price page

		do case
			case ncompnumber = 198     && As of 01/01/2007 overall discount will be shown for Weyerhaeuser
				do pricepageweyshow
				? "pricepageweyshow"
			* case ncompnumber = 198 .and. lshowdiscount = .t.    && overall discount shown for Weyerhaeuser
				* do pricepagewey
				* ? "pricepagewey"
			* case ncompnumber = 198 .and. lshowdiscount = .f.    && overall discount for Weyerhaeuser
				* do pricepageweyover
				* ? "pricepageweyover"
			case lforavanti = .t.                            && material only discount for AVANTI
				do pricepageavanti
				? "pricepageavanti"
			case nmaterialdisc = 0 .and. ninstalldisc = 0    && no discounts
				do pricepagend
				? "pricepagend"
			case nmaterialdisc > 0 .and. ninstalldisc > 0    && material and installation discount
				do pricepagemaid
				? "pricepagemaid"
			case nmaterialdisc > 0 .and. ncompnumber = 297   && material only discount for Menasha
				do pricepagemodshown
				? "pricepagemodshown"
			case nmaterialdisc > 0 .and. ncompnumber = 546   && material only discount for Smurfit
				* do pricepagesmurfitnd
				* ? "pricepagesmurfitnd"
				do pricepagemodshown
				? "pricepagemodshown"
*!*				case nmaterialdisc > 0 .and. ncompnumber = 546   && no discounts shown for Smurfit
*!*					do pricepagemodnetonly
*!*					? "pricepagemodnetonly"
			case nmaterialdisc > 0 .and. ncompnumber = 630   && material only discount for Georgia-Pacific
				do pricepagemodshown
				? "pricepagemodshown"
			case nmaterialdisc > 0 .and. ncompnumber = 682   && material only discount for Inland
				do pricepagemodshown
				? "pricepagemodshown"
			case nmaterialdisc > 0 .and. ncompnumber = 755   && material only discount for Longview Fibre
				do pricepagemodshown
				? "pricepagemodshown"
			case nmaterialdisc > 0 .and. ncompnumber = 841   && material only discount for PCA
				do pricepagemodshown
				? "pricepagemodshown"
			case nmaterialdisc > 0 .and. ncompnumber = 879   && material only discount for Rock Tenn *old Southern Container
				* was 899
				do pricepagemodshown
				? "pricepagemodshown"
			case nmaterialdisc > 0 .and. ncompnumber = 1089  && material only discount for Rock-Tenn *Alliance Group
				* do pricepagemodnetonly
				* ? "pricepagemodnetonly"
				* Changed 11/05/2010
				do pricepagemodshown
				? "pricepagemodshown"
			case nmaterialdisc > 0 .and. ncompnumber = 1098 ;
			                 .and. ncompleteacsno # 1098068  && material only discount for IP less IP Ireland
				* do pricepagemodnetonly
				*? "pricepagemodnetonly"
				do pricepagemodshown
				? "pricepagemodshown"
			case nmaterialdisc > 0                           && material only discount
				do pricepagemodshown
				? "pricepagemodshown"
			otherwise                                        && default to no discount for now
				do pricepagend
				? "otherwise pricepagend"
		endcase
		**************************************

		* remove the fax sheet if not required
		if lforavanti = .t. .or. ldofaxsheet = .f.
			WordApp.ActiveDocument.Paragraphs(9).Range.Select
			WordApp.Selection.Tables(1).Select
			WordApp.Selection.Tables(1).Delete
			WordApp.ActiveDocument.Range(WordApp.ActiveDocument.Paragraphs(1).Range.Start,WordApp.ActiveDocument.Paragraphs(17).Range.End).Delete
		endif

		* release form
		* progressreport.release

		* Make Visible
		WordApp.Application.Visible = .t.

		** Save document
		** document directory
		* cdocdir = "\\ACS-FS1\SALESSRV\WORK IN PROGRESS\DOCUMENTS\WORD\"
		* ctabledir = "\\ACS-FS1\SALESSRV\WORK IN PROGRESS\DRAWINGS\QUOTE\"
		* ccopytodir = "\\ACS-FS1\ADM\WORK IN PROGRESS\DOCUMENTS\WORD\QUOTE\"
		* save in tabledir for now
		* WordDoc.Saveas(ctabledir + cdrawing + 'PP')
		*!*	WordDoc.Saveas(ccopytodir + cdrawing + 'PP')

		* Quit Application, Set to nil
		* WordApp.quit
		* WordApp = "nil"

		* Print the Address Label
		if lonnetwork = .t.
			** if both of these variables are true there are no names to print
			* if lnoname = .t. .and. lcopynoname = .t.
				* * don't print Address Label
			* else
				* do writeaddresslabels
			* endif
			* Print the Quote Folder Labels
			do writequotelabels
		endif
	else
		* cancel selected in the form cformspath + 'selectcopytoname'
		* do nothing
		MESSAGEBOX('NO Printout will be done', 32, 'CANCEL PRESSED')
	endif   && if lstoprunning = .f.
endif   && lcontinuewriting = .t.

release progressreport
*-- EOP WRITEPPDOC

Procedure numberformat
Parameters cellpassed,fieldpassed
do case
	* seven digit
	case &fieldpassed. >  999999 .and. &fieldpassed. < 10000000
		.cell(tcount,cellpassed).Range.InsertAfter("$" + ;
		substr(str(&fieldpassed.,8),1,2) + "," + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3))
	* six digit
	case &fieldpassed. >   99999 .and. &fieldpassed. < 1000000
		.cell(tcount,cellpassed).Range.InsertAfter("$  " + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3))
	* five digit
	case &fieldpassed. >    9999 .and. &fieldpassed. < 100000
		.cell(tcount,cellpassed).Range.InsertAfter("$   " + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3))
	* four digit
	case &fieldpassed. >     999 .and. &fieldpassed. < 10000
		.cell(tcount,cellpassed).Range.InsertAfter("$    " + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3))
	* three digit
	case &fieldpassed. >      99 .and. &fieldpassed. < 1000
		.cell(tcount,cellpassed).Range.InsertAfter("$         " + ;
		substr(str(&fieldpassed.,8),6,3))
	* two digit
	case &fieldpassed. >       9 .and. &fieldpassed. < 100
		.cell(tcount,cellpassed).Range.InsertAfter("$          " + ;
		substr(str(&fieldpassed.,8),6,3))
	* one digit
	case &fieldpassed. =>      0 .and. &fieldpassed. < 10
		.cell(tcount,cellpassed).Range.InsertAfter("$           " + ;
		substr(str(&fieldpassed.,8),6,3))
	* minus one digit
	case &fieldpassed. <       0 .and. &fieldpassed. > -10
		.cell(tcount,cellpassed).Range.InsertAfter("$           -(" + ;
		substr(str(&fieldpassed.,8),8,1) + ")")
	* minus two digit
	case &fieldpassed. <      -9 .and. &fieldpassed. > -100
		.cell(tcount,cellpassed).Range.InsertAfter("$       -(" + ;
		substr(str(&fieldpassed.,8),7,2) + ")")
	* minus three digit
	case &fieldpassed. <     -99 .and. &fieldpassed. > -1000
		.cell(tcount,cellpassed).Range.InsertAfter("$     -(" + ;
		substr(str(&fieldpassed.,8),6,3) + ")")
	* minus four digit
	case &fieldpassed. <    -999 .and. &fieldpassed. > -10000
		.cell(tcount,cellpassed).Range.InsertAfter("$  -(" + ;
		substr(str(&fieldpassed.,8),5,1) + "," + ;
		substr(str(&fieldpassed.,8),6,3) + ")")
	* minus five digit
	case &fieldpassed. <   -9999 .and. &fieldpassed. > -100000
		.cell(tcount,cellpassed).Range.InsertAfter("$ -(" + ;
		substr(str(&fieldpassed.,8),4,2) + "," + ;
		substr(str(&fieldpassed.,8),6,3) + ")")
	* minus five digit
	case &fieldpassed. <  -99999 .and. &fieldpassed. > -1000000
		.cell(tcount,cellpassed).Range.InsertAfter("$ -(" + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3) + ")")
	* minus six digit
	case &fieldpassed. < -999999 .and. &fieldpassed. > -10000000
		.cell(tcount,cellpassed).Range.InsertAfter("$ -(" + ;
		substr(str(&fieldpassed.,8),2,1) + "," + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3) + ")")
	otherwise
		.cell(tcount,cellpassed).Range.InsertAfter("$  " + alltrim(str(&fieldpassed.)))
endcase
endproc

Procedure numberformatopt
Parameters fieldpassedo
do case
	* eight digit
	case &fieldpassedo. >  9999999 .and. &fieldpassedo. < 100000000
		WordApp.Selection.InsertAfter("$ " + ;
		substr(str(&fieldpassedo.,8),1,2) + "," + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3))
	* seven digit
	case &fieldpassedo. >  999999 .and. &fieldpassedo. < 10000000
		WordApp.Selection.InsertAfter("$  " + ;
		substr(str(&fieldpassedo.,8),1,2) + "," + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3))
	* six digit
	case &fieldpassedo. >   99999 .and. &fieldpassedo. < 1000000
		WordApp.Selection.InsertAfter("$      " + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3))
	* five digit
	case &fieldpassedo. >    9999 .and. &fieldpassedo. < 100000
		WordApp.Selection.InsertAfter("$       " + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3))
	* four digit
	case &fieldpassedo. >     999 .and. &fieldpassedo. < 10000
		WordApp.Selection.InsertAfter("$        " + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3))
	* three digit
	case &fieldpassedo. >      99 .and. &fieldpassedo. < 1000
		WordApp.Selection.InsertAfter("$             " + ;
		substr(str(&fieldpassedo.,8),6,3))
	* two digit
	case &fieldpassedo. >       9 .and. &fieldpassedo. < 100
		WordApp.Selection.InsertAfter("$              " + ;
		substr(str(&fieldpassedo.,8),6,3))
	* one digit
	case &fieldpassedo. =>      0 .and. &fieldpassedo. < 10
		WordApp.Selection.InsertAfter("$               " + ;
		substr(str(&fieldpassedo.,8),6,3))
	* minus one digit
	case &fieldpassedo. <       0 .and. &fieldpassedo. > -10
		WordApp.Selection.InsertAfter("$               -(" + ;
		substr(str(&fieldpassedo.,8),8,1) + ")")
	* minus two digit
	case &fieldpassedo. <      -9 .and. &fieldpassedo. > -100
		WordApp.Selection.InsertAfter("$             -(" + ;
		substr(str(&fieldpassedo.,8),7,2) + ")")
	* minus three digit
	case &fieldpassedo. <     -99 .and. &fieldpassedo. > -1000
		WordApp.Selection.InsertAfter("$           -(" + ;
		substr(str(&fieldpassedo.,8),6,3) + ")")
	* minus four digit
	case &fieldpassedo. <    -999 .and. &fieldpassedo. > -10000
		WordApp.Selection.InsertAfter("$        -(" + ;
		substr(str(&fieldpassedo.,8),5,1) + "," + ;
		substr(str(&fieldpassedo.,8),6,3) + ")")
	* minus five digit
	case &fieldpassedo. <   -9999 .and. &fieldpassedo. > -100000
		WordApp.Selection.InsertAfter("$      -(" + ;
		substr(str(&fieldpassedo.,8),4,2) + "," + ;
		substr(str(&fieldpassedo.,8),6,3) + ")")
	* minus five digit
	case &fieldpassedo. <  -99999 .and. &fieldpassedo. > -1000000
		WordApp.Selection.InsertAfter("$    -(" + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3) + ")")
	* minus six digit
	case &fieldpassedo. < -999999 .and. &fieldpassedo. > -10000000
		WordApp.Selection.InsertAfter("$ -(" + ;
		substr(str(&fieldpassedo.,8),2,1) + "," + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3) + ")")
	otherwise
		WordApp.Selection.InsertAfter("$  " + alltrim(str(&fieldpassedo.)))
endcase
endproc