* Name......... WRITE the Price Page DOCument for AVANTI conveyors from foxpro into word program
* Date......... 04/08/2003
* Caller....... writingdocs_app.prg
* Notes........ This is the second attempt at using Visual Basic language to create a WORD document.
*               chr(11) = Shift Enter  chr(9) = Tab
*               WordApp.selection.TypeParagraph = enter
*               Added the System Engineer to the Price Page
*               The mail to information corrected.
*               Changed System Engineer to PLC Start Up
*               Added if statement for sold at price page ( 'S.DBF' )
*               Print the Address Label
*               Print the Quote Labels


*!*		if upper(alias("C")) = 'QUOTETRK' .and. upper(alias("D")) = 'QUOTHIST' .and. ;
*!*		   upper(alias("E")) = 'SALESSER' .and. upper(alias("F")) = 'SALETERR'

*!*			* determine if a reference is needed
*!*			select a
*!*			go bottom
*!*			* if group not equal to '$' (total) or group is a number
*!*			if alltrim(group) # '$' .and. asc(alltrim(group)) < 58
*!*				* since the report can not LOCATE from within the program
*!*				* create a new table with the option reference information in it
*!*				select a
*!*				copy structure to ctemppath + cdrawing + 'REF'
*!*				select g
*!*				use ctemppath + cdrawing + 'REF' in g
*!*				index on group to ctemppath + cdrawing + 'group'
*!*			endif
*!*			* find references
*!*			select a
*!*			go top
*!*			PUBLIC lreffound    && reference found
*!*			store .f. to lreffound
*!*			locate for isblank(option_ref) = .f.
*!*			do while found() = .t.
*!*				store .t. to lreffound
*!*				select g
*!*				append blank
*!*				replace g.GROUP with A.GROUP
*!*				select a
*!*				store a.OPTION_REF to cfind
*!*				store recno("a") to nreturnto
*!*				go top
*!*				do while alltrim(a.GROUP) # alltrim(cfind)
*!*					skip 1 in a
*!*				enddo
*!*				replace g.MATERIAL with a.MATERIAL
*!*				replace g.INSTALL  with a.INSTALL
*!*				replace g.RELOCATE with a.RELOCATE
*!*				replace g.REMOVE   with a.REMOVE
*!*				replace g.REWORK   with a.REWORK
*!*				replace g.RTOTAL   with a.RTOTAL
*!*				replace g.PC_TECH  with a.PC_TECH
*!*				replace g.TOTAL    with a.TOTAL
*!*				goto nreturnto
*!*				continue
*!*			enddo

*!*			* reset table in a to the top
*!*			select a
*!*			go top

*!*			* set relation here
*!*			if lreffound = .t.
*!*				set relation to group into g
*!*			endif

*!*			* get the sales service and sales territory initials
*!*			PUBLIC cssi,csti
*!*			select a
*!*			if seek(b.num_quote,"quotetrk") = .t.
*!*				store quotetrk.ssi to cssi
*!*				store quotetrk.st to csti
*!*			else
*!*				if seek(b.num_quote,"quothist") = .t.
*!*					store quothist.ssi to cssi
*!*					store quothist.st to csti
*!*				endif
*!*			endif

*!*			* get the sales service first and last name
*!*			PUBLIC cssname,csssname
*!*			if seek(cssi,"salesser") = .t.
*!*				store e.sal + rtrim(e.name_f) + " " + rtrim(e.name_l) to cssname
*!*				store rtrim(e.name_f) + " " + rtrim(e.name_l) to csssname
*!*			else
*!*				store "no name" to cssname,csssname
*!*			endif
*!*			* get the sales service full name for those that want to use it
*!*			if seek(cssi,"salesser") = .t.
*!*				if alltrim(cssi) = "MTM"
*!*					store e.sal + rtrim(e.name_f) + " " + substr(e.name_m,1,1) + ". " + rtrim(e.name_l) to cssname
*!*					store rtrim(e.name_f) + " " + substr(e.name_m,1,1) + ". " +  + rtrim(e.name_l) to csssname
*!*				endif
*!*			endif
*!*			* get the sales territory first and last name
*!*			PUBLIC cstname
*!*			if seek(csti,"saleterr") = .t.
*!*				store rtrim(f.name_f) + " " + rtrim(f.name_l) to cstname
*!*				* set Mike Lucado's to Mike Lucado
*!*				if csti = "019"
*!*					store "Mike Lucado" to cstname
*!*				endif
*!*				* get the sales territory full name for those that want to use it
*!*				if seek(cssi,"salesser") = .t.
*!*					* Michael P. Shenigo and Terry D. Davis
*!*					if csti = "013" .or. csti = "021"
*!*						store rtrim(f.name_f) + " " + substr(f.name_m,1,1) + ". " + rtrim(f.name_l) to cstname
*!*					endif
*!*				endif
*!*			else
*!*				store "no name" to cstname
*!*			endif

*!*			* Progress report form variables
*!*			PUBLIC nreccount, nrecnum
*!*			store reccount("a") + 1 to nreccount
*!*			store recno("a") to nrecnum

*!*			do form cformspath + "progressreport"
*!*			* 

*!*			******
*!*			* Create Application & Document
*!*			WordApp = CreateObject("Word.application.8")
*!*			*!*	 * Open document template
*!*			*!*	WordDoc = WordApp.Documents.open('c:\my documents\quote2.dot')
*!*			* Open edits the template and does not allow saving as a .doc file thus
*!*			* Add a new document using the quote2 template
*!*			*                          .Add(Template:                   , NewTemplate)
*!*			WordDoc = WordApp.Documents.Add('c:\my documents\quote2.dot', .f.)
*!*			* When the network template is used the above statement will change to
*!*			* WordDoc = WordApp.Documents.Add('G:\Word Templates\Quote Templates\quote2.dot', .f.)
*!*			* New caption
*!*			WordApp.Caption = "Quotation Price Page Documents"
*!*			* Make Visible
*!*			WordApp.Application.Visible = .t.

*!*			******** Page 1

*!*			* Insert the customers complete name in paragraph 5
*!*			if len(cppsal+cppnamef+cppnamel) > 0    && name selected
*!*				WordApp.activedocument.paragraphs(5).range.select
*!*				WordApp.selection.StartOf
*!*				WordApp.selection.MoveRight(wdCharacter,18)
*!*				WordApp.selection.Delete(wdCharacter,1)
*!*				WordApp.selection.insertafter(cppsal+" "+cppnamef+" "+cppnamel)
*!*			else
*!*				* do nothing
*!*			endif

*!*			* Insert mdy(date()) in paragraph 7
*!*			WordApp.activedocument.paragraphs(7).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.MoveDown(wdLine,1)
*!*			WordApp.selection.MoveLeft(wdCharacter,20)
*!*			WordApp.selection.Delete(wdCharacter,5)
*!*			WordApp.selection.insertafter(mdy(date()))
*!*			WordApp.selection.Font.Bold = .f.
*!*			WordApp.selection.Font.Size = 11

*!*			* Insert selected name in paragraph 9 if one was selected
*!*			if len(cppsal+ cppnamef+ cppnamel) > 0    && name selected
*!*				WordApp.activedocument.paragraphs(9).range.select
*!*				WordApp.selection.StartOf
*!*				if cppsal = "Mr."
*!*					WordApp.selection.MoveRight(wdCharacter,8)
*!*					WordApp.selection.insertafter(cppnamef+" "+cppnamel)
*!*				else
*!*					WordApp.selection.MoveRight(wdCharacter,4)
*!*					WordApp.selection.Delete(wdCharacter,4)
*!*					WordApp.selection.insertafter(cppsal+" "+cppnamef+" "+cppnamel)
*!*				endif
*!*			else
*!*				* do nothing
*!*			endif

*!*			******
*!*			* Insert variable cssname in paragraph 10
*!*			WordApp.activedocument.paragraphs(10).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.MoveRight(wdCharacter,6)
*!*			WordApp.selection.insertbefore(cssname)
*!*			******

*!*			******
*!*			* Insert customer plant name and location
*!*			* go to Paragraph 12 and insert/edit information
*!*			WordApp.activedocument.paragraphs(12).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.MoveRight(wdCharacter,1)
*!*			WordApp.selection.insertafter(ccustomer)
*!*			WordApp.selection.MoveDown(wdLine,1)
*!*			WordApp.selection.insertafter(clocation)
*!*			******

*!*			******
*!*			* Insert customer plant phone number
*!*			* go to Paragraph 15 and insert/edit information
*!*			WordApp.activedocument.paragraphs(15).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.MoveRight(wdCharacter,11)
*!*			WordApp.selection.insertafter(cppphone)
*!*			******

*!*			******
*!*			* Insert customer plant fax number
*!*			* go to Paragraph 17 and insert/edit information
*!*			WordApp.activedocument.paragraphs(17).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.MoveRight(wdCharacter,11)
*!*			WordApp.selection.insertafter(cppfax)
*!*			******

*!*			******
*!*			* Insert selected name here if one was selected in paragraph 20
*!*			if len(cppsal+cppnamef+cppnamel) > 0    && name selected
*!*				WordApp.activedocument.paragraphs(20).range.select
*!*				WordApp.selection.StartOf
*!*				if cppsal = "Mr."
*!*					WordApp.selection.MoveRight(wdCharacter,11)
*!*					WordApp.selection.insertafter(cppnamel)
*!*				else
*!*					WordApp.selection.MoveRight(wdCharacter,7)
*!*					WordApp.selection.Delete(wdCharacter,4)
*!*					WordApp.selection.insertafter(cppsal+" "+cppnamel)
*!*				endif
*!*			else
*!*				* do nothing
*!*			endif
*!*			******

*!*			******
*!*			* Enter the quotation title in the sentence
*!*			* We are pleased to submit this "quotation title" proposal for your review.
*!*			* find the word this and insert the variable ctitle thereafter in paragraph 22
*!*			WordApp.activedocument.paragraphs(22).range.select
*!*			WordApp.selection.Find.ClearFormatting
*!*			With WordApp.selection.Find
*!*				.Text = "this"
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.selection.Find.Execute
*!*			WordApp.selection.MoveEnd
*!*			* then enter the layout title
*!*			WordApp.selection.insertafter(ctitle + " ")
*!*			* Then change the font of ctitle
*!*			WordApp.activedocument.paragraphs(22).range.select
*!*			WordApp.selection.Find.ClearFormatting
*!*			With WordApp.selection.Find
*!*				.Text = ctitle
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.selection.Find.Execute
*!*			WordApp.selection.Font.Bold = .t.
*!*			WordApp.selection.Font.Italic = .t.
*!*			******

*!*			******
*!*			* Enter the sales territory in the sentence
*!*			* Your "ACS" representative, "name", will be contacting you ...
*!*			* find the word this and insert the variable cstname thereafter in paragraph 23
*!*			WordApp.activedocument.paragraphs(23).range.select
*!*			WordApp.selection.Find.ClearFormatting
*!*			With WordApp.selection.Find
*!*				.Text = ", ,"
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.selection.Find.Execute
*!*			WordApp.selection.MoveLeft(wdCharacter,1)
*!*			WordApp.selection.MoveRight(wdCharacter,2)
*!*			* Then enter number symbol plus the layout number
*!*			WordApp.selection.insertbefore(alltrim(cstname))
*!*			WordApp.selection.Font.Italic = .t.
*!*			******

*!*			******
*!*			* Place sales service name in paragraph 29 above Sales Service Engineer
*!*			WordApp.activedocument.paragraphs(29).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.insertbefore(csssname)
*!*			WordApp.selection.MoveDown(wdline,3)
*!*			WordApp.selection.MoveLeft(wdCharacter,3)
*!*			* Place sales service initials
*!*			WordApp.selection.insertbefore(upper(cssi))
*!*			* Select paragraph and add quotation number
*!*			WordApp.activedocument.paragraphs(29).range.select
*!*			WordApp.selection.Find.ClearFormatting
*!*			With WordApp.selection.Find
*!*				.Text = "02-"
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.selection.Find.Execute
*!*			if len(alltrim(cquote)) > 0
*!*				* erases the template information 02-
*!*				WordApp.selection.Delete
*!*				* insert quotation number
*!*				WordApp.selection.insertafter(cquote)
*!*			endif
*!*			******

*!*			******** This completes the first page

*!*			******** New Page 2
*!*			******** Old Page 6

*!*			******
*!*			* Paragraph 32 is the Cover Sheet heading
*!*			* Go to Paragraph 32 and insert/edit information
*!*			WordApp.activedocument.paragraphs(32).range.select
*!*			WordApp.selection.StartOf
*!*			if len(alltrim(ccustomer)) > 0
*!*				WordApp.selection.insertafter(ccustomer)
*!*			endif
*!*			WordApp.selection.MoveDown(wdLine,1)
*!*			if len(alltrim(clocation)) > 0
*!*				WordApp.selection.insertafter(clocation)
*!*			endif
*!*			* Select paragraph and add quotation number
*!*			WordApp.activedocument.paragraphs(32).range.select
*!*			WordApp.selection.Find.ClearFormatting
*!*			With WordApp.selection.Find
*!*				.Text = "02-"
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.selection.Find.Execute
*!*			if len(alltrim(cquote)) > 0
*!*				* erases the template information 02-
*!*				WordApp.selection.Delete
*!*				* insert quotation number
*!*				WordApp.selection.insertafter(cquote)
*!*			endif
*!*		WordApp.selection.MoveDown(wdLine,1)
*!*		WordApp.selection.TypeBackspace
*!*		WordApp.selection.TypeBackspace
*!*		WordApp.selection.insertafter(ctitle + '"' + chr(11))
*!*		WordApp.selection.MoveDown(wdLine,1)
*!*		* inserts month, day and year
*!*		WordApp.selection.insertbefore(mdy(date()))
*!*		******

*!*		******** This completes the second page

*!*			******** New Page 3
*!*			******** Old Page 2

*!*			******
*!*			* Enter mdy(date())
*!*			* Insert mdy(date()) in paragraph(33)
*!*			WordApp.activedocument.paragraphs(33).range.select
*!*			WordApp.selection.Collapse(wdCollapseStart)
*!*			WordApp.selection.MoveDown(wdline,15)
*!*			WordApp.selection.Delete(wdCharacter,5)
*!*			WordApp.selection.insertafter(mdy(date()))
*!*			******

*!*			******
*!*			* Move down 4 lines then insert customer name
*!*			WordApp.selection.MoveDown(wdline,4)
*!*			if len(cppsal+cppnamef+cppnamel) > 0    && name selected
*!*				WordApp.selection.insertafter(cppsal+" "+cppnamef+" "+cppnamel)
*!*			else
*!*				* do nothing
*!*			endif

*!*			******
*!*			* Insert company name
*!*			* Move down 1 lines then insert company name
*!*			WordApp.selection.MoveDown(wdline,1)
*!*			WordApp.selection.insertafter(cppcustomer)
*!*			* Move down 1 line and check for an address
*!*			WordApp.selection.MoveDown(wdline,1)
*!*			if len(alltrim(cppaddress)) > 0    && got an address
*!*				* Insert Address
*!*				WordApp.selection.insertbefore(cppaddress)
*!*			endif
*!*			* Check for customer location
*!*			WordApp.selection.MoveDown(wdline,1)
*!*			if len(alltrim(cpplocation)) > 0
*!*				WordApp.selection.StartOf
*!*				WordApp.selection.insertbefore(cpplocation + " " + cppzip)
*!*			endif
*!*			******

*!*			******
*!*			* Dear Mr. : ?
*!*			* Last name in paragraph(34)
*!*			WordApp.activedocument.paragraphs(34).range.select
*!*			WordApp.selection.EndKey
*!*			WordApp.selection.MoveLeft(wdCharacter,1)
*!*			if len(cppsal+cppnamef+cppnamel) > 0    && name selected
*!*				if cppsal = "Mr."
*!*					WordApp.selection.insertafter(cppnamel)
*!*				else
*!*					WordApp.selection.MoveLeft(wdCharacter,4)
*!*					WordApp.selection.Delete(wdCharacter,4)
*!*					WordApp.selection.insertafter(cppsal+" "+cppnamel)
*!*				endif
*!*			else
*!*				* do nothing
*!*			endif
*!*			******

*!*			******
*!*			* Enter the quotation title in the sentence
*!*			* We are pleased to submit this "quotation title" proposal for your review.
*!*			* find the word this and insert the variable ctitle thereafter paragraph 35
*!*			WordApp.activedocument.paragraphs(35).range.select
*!*			WordApp.selection.Find.ClearFormatting
*!*			With WordApp.selection.Find
*!*				.Text = "this"
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.selection.Find.Execute
*!*			WordApp.selection.MoveEnd
*!*			* then enter the layout title
*!*			WordApp.selection.insertafter(ctitle + " ")
*!*			* then change font of ctitle
*!*			WordApp.activedocument.paragraphs(35).range.select
*!*			WordApp.selection.Find.ClearFormatting
*!*			With WordApp.selection.Find
*!*				.Text = ctitle
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.selection.Find.Execute
*!*			WordApp.selection.Font.Bold = .t.
*!*			WordApp.selection.Font.Italic = .t.
*!*			******

*!*			******
*!*			* Enter the sales territory in the sentence
*!*			* Your "ACS" representative, "name", will be contacting you ...
*!*			* find the word this and insert the variable cstname thereafter paragraph 36
*!*			WordApp.activedocument.paragraphs(36).range.select
*!*			WordApp.selection.Find.ClearFormatting
*!*			With WordApp.selection.Find
*!*				.Text = ", ,"
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.selection.Find.Execute
*!*			WordApp.selection.MoveLeft(wdCharacter,1)
*!*			WordApp.selection.MoveRight(wdCharacter,2)
*!*			* then enter the sales territory name
*!*			WordApp.selection.insertbefore(alltrim(cstname))
*!*			WordApp.selection.Font.Italic = .t.
*!*			******

*!*			******
*!*			* Place sales service name in paragraph 41 above Sales Service Engineer
*!*			WordApp.activedocument.paragraphs(41).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.insertbefore(csssname)
*!*			WordApp.selection.MoveDown(wdline,3)
*!*			WordApp.selection.MoveLeft(wdCharacter,3)
*!*			* Place sales service initials
*!*			WordApp.selection.insertbefore(upper(cssi))
*!*			* Select paragraph and add quotation number
*!*			WordApp.activedocument.paragraphs(41).range.select
*!*			WordApp.selection.Find.ClearFormatting
*!*			With WordApp.selection.Find
*!*				.Text = "02-"
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.selection.Find.Execute
*!*			if len(alltrim(cquote)) > 0
*!*				* erases the template information 02-
*!*				WordApp.selection.Delete
*!*				* insert quotation number
*!*				WordApp.selection.insertafter(cquote)
*!*			endif
*!*			* insert the copies for names
*!*			if lgotcopypernumber = .t.
*!*				if len(tempselectarray[1,1]) > 0
*!*					* Select paragraph 42 and name(s) that will get a copy
*!*					WordApp.activedocument.paragraphs(42).range.select
*!*					WordApp.selection.StartOf
*!*					WordApp.selection.MoveRight(wdCharacter,4)
*!*					* print their names
*!*					for f = 1 to 10
*!*						if len(tempselectarray[f,1]) > 0
*!*							if f = 1
*!*								WordApp.selection.insertafter(tempselectarray[f,2]+" "+tempselectarray[f,3]+" "+tempselectarray[f,4])
*!*							endif
*!*							if f > 1
*!*								* a shift enter plus a tab then the name
*!*								WordApp.selection.insertafter(chr(11)+chr(9))
*!*								WordApp.selection.insertafter(tempselectarray[f,2]+" "+tempselectarray[f,3]+" "+tempselectarray[f,4])
*!*							endif
*!*						else
*!*							* do nothing
*!*						endif
*!*					endfor
*!*				else
*!*					* do nothing
*!*				endif
*!*			endif
*!*			******

*!*			******** This completes the third page

*!*			******** New Page 4
*!*			******** Old Page 3

*!*			******
*!*			* Paragraph 44 is the Price Page heading
*!*			* go to Paragraph 44 and insert/edit information
*!*			WordApp.activedocument.paragraphs(44).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.MoveDown
*!*			if len(alltrim(ccustomer)) > 0
*!*				WordApp.selection.insertafter(ccustomer)
*!*			endif
*!*			WordApp.selection.MoveDown(wdLine,1)
*!*			if len(alltrim(clocation)) > 0
*!*				WordApp.selection.insertafter(clocation)
*!*			endif
*!*			if len(alltrim(cquote)) > 0
*!*				WordApp.activedocument.paragraphs(44).range.select
*!*				WordApp.selection.Find.ClearFormatting
*!*				With WordApp.selection.Find
*!*					.Text = "02-"
*!*					.Forward = .t.
*!*					.Wrap = wdFindContinue
*!*					.Format = .f.
*!*					.MatchCase = .t.
*!*					.MatchWholeWord = .t.
*!*					.MatchWildcards = .f.
*!*					.MatchSoundsLike = .f.
*!*					.MatchAllWordForms = .f.
*!*				EndWith
*!*				WordApp.selection.Find.Execute
*!*				* erases the template information 02-
*!*				WordApp.selection.Delete
*!*				* insert quotation number
*!*				WordApp.selection.insertafter(cquote)
*!*			endif
*!*			* remove the template information , 2001
*!*			WordApp.activedocument.paragraphs(44).range.select
*!*			WordApp.selection.Find.ClearFormatting
*!*			With WordApp.selection.Find
*!*				.Text = ", 2002"
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.selection.Find.Execute
*!*			* erases the template information , 2002
*!*			WordApp.selection.Delete
*!*			WordApp.selection.insertbefore(mdy(date()))
*!*			******

*!*			******
*!*			* go to next paragraph
*!*			WordApp.selection.MoveDown(wdLine,1)
*!*			* find the number symbol(#) and erase
*!*			WordApp.selection.Find.ClearFormatting
*!*			With WordApp.selection.Find
*!*				.Text = "#"
*!*				.Forward = .t.
*!*				.Wrap = wdFindContinue
*!*				.Format = .f.
*!*				.MatchCase = .t.
*!*				.MatchWholeWord = .t.
*!*				.MatchWildcards = .f.
*!*				.MatchSoundsLike = .f.
*!*				.MatchAllWordForms = .f.
*!*			EndWith
*!*			WordApp.selection.Find.Execute
*!*			WordApp.selection.TypeBackspace
*!*			* then enter number symbol plus the layout number
*!*			WordApp.selection.insertafter("#" + clayout)
*!*			* then enter the quote title in paragraph 46
*!*			WordApp.activedocument.paragraphs(46).range.select
*!*			WordApp.selection.insertbefore(ctitle)
*!*			******

*!*			******** This completes the fourth page for now

*!*			******** New Page 5
*!*			******** Old Page 7

*!*			******
*!*			* Paragraph 115 is the Sequence Of Operations Sheet heading
*!*			* go to Paragraph 115 and insert/edit information
*!*			WordApp.activedocument.paragraphs(115).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.MoveDown(wdLine,1)
*!*			if len(alltrim(ccustomer)) > 0
*!*				WordApp.selection.insertafter(ccustomer)
*!*			endif
*!*			WordApp.selection.MoveDown(wdLine,1)
*!*			if len(alltrim(clocation)) > 0
*!*				WordApp.selection.insertafter(clocation)
*!*			endif
*!*			if len(alltrim(cquote)) > 0
*!*				WordApp.activedocument.paragraphs(115).range.select
*!*				WordApp.selection.Find.ClearFormatting
*!*				With WordApp.selection.Find
*!*					.Text = "02-"
*!*					.Forward = .t.
*!*					.Wrap = wdFindContinue
*!*					.Format = .f.
*!*					.MatchCase = .t.
*!*					.MatchWholeWord = .t.
*!*					.MatchWildcards = .f.
*!*					.MatchSoundsLike = .f.
*!*					.MatchAllWordForms = .f.
*!*				EndWith
*!*				WordApp.selection.Find.Execute
*!*				* erases the template information 02-
*!*				WordApp.selection.Delete
*!*				* insert quotation number
*!*				WordApp.selection.insertafter(cquote+chr(11))
*!*			endif
*!*			WordApp.selection.MoveDown(wdLine,1)
*!*			* inserts month, day and year
*!*			WordApp.selection.insertbefore(mdy(date()))
*!*			******

*!*			******** This completes the fifth page

*!*			******** New Page 6
*!*			******** Old Page 5

*!*			*!*	******
*!*			* Do the General Equipment Specifications next
*!*			* Paragraph 119 is the first paragraph on General Equipment Specifications page
*!*			* go to Paragraph 119 and insert/edit information
*!*			* top of roller
*!*			WordApp.activedocument.paragraphs(119).range.select
*!*			WordApp.selection.EndKey
*!*			WordApp.selection.MoveLeft(wdCharacter,1)
*!*			WordApp.selection.insertafter(alltrim(str(b.tor)) + '"')
*!*			* roll centers
*!*			WordApp.activedocument.paragraphs(120).range.select
*!*			WordApp.selection.EndKey
*!*			WordApp.selection.MoveLeft(wdCharacter,1)
*!*			WordApp.selection.insertafter(alltrim(str(b.rtr)) + '"')
*!*			* conveyor speed
*!*			WordApp.activedocument.paragraphs(121).range.select
*!*			WordApp.selection.EndKey
*!*			WordApp.selection.MoveLeft(wdCharacter,1)
*!*			WordApp.selection.insertafter(alltrim(str(b.speed)) + ' FPM')
*!*			* floor conveyor color
*!*			WordApp.activedocument.paragraphs(122).range.select
*!*			WordApp.selection.EndKey
*!*			WordApp.selection.MoveLeft(wdCharacter,1)
*!*			WordApp.selection.insertafter("To Be Determined")
*!*			* floor conveyor paint spec. number
*!*			WordApp.activedocument.paragraphs(123).range.select
*!*			WordApp.selection.EndKey
*!*			WordApp.selection.MoveLeft(wdCharacter,1)
*!*			WordApp.selection.insertafter("To Be Determined")
*!*			* bundle conveyor color
*!*			WordApp.activedocument.paragraphs(124).range.select
*!*			WordApp.selection.EndKey
*!*			WordApp.selection.MoveLeft(wdCharacter,1)
*!*			WordApp.selection.insertafter("To Be Determined" + " (See Note   & Option    )")
*!*			* voltage
*!*			WordApp.activedocument.paragraphs(125).range.select
*!*			WordApp.selection.EndKey
*!*			WordApp.selection.MoveLeft(wdCharacter,1)
*!*			WordApp.selection.insertafter(alltrim(b.voltage))
*!*			******

*!*			******** This completes the sixth page

*!*			******** New Page 7
*!*			******** Old Page 4

*!*			******
*!*			* Do the General Notes heading next because you don't know how many
*!*			*  paragraphs will be added on the price page
*!*			* Paragraph 128 is the General Notes heading
*!*			* go to Paragraph 128 and insert/edit information
*!*			WordApp.activedocument.paragraphs(128).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.MoveDown(wdLine,2)
*!*			if len(alltrim(ccustomer)) > 0
*!*				WordApp.selection.insertafter(ccustomer)
*!*			endif
*!*			WordApp.selection.MoveDown(wdLine,1)
*!*			if len(alltrim(clocation)) > 0
*!*				WordApp.selection.insertafter(clocation)
*!*			endif
*!*			if len(alltrim(cquote)) > 0
*!*				WordApp.activedocument.paragraphs(128).range.select
*!*				WordApp.selection.Find.ClearFormatting
*!*				With WordApp.selection.Find
*!*					.Text = "02-"
*!*					.Forward = .t.
*!*					.Wrap = wdFindContinue
*!*					.Format = .f.
*!*					.MatchCase = .t.
*!*					.MatchWholeWord = .t.
*!*					.MatchWildcards = .f.
*!*					.MatchSoundsLike = .f.
*!*					.MatchAllWordForms = .f.
*!*				EndWith
*!*				WordApp.selection.Find.Execute
*!*				* erases the template information 02-
*!*				WordApp.selection.Delete
*!*				WordApp.selection.insertafter(cquote)
*!*			endif
*!*			* inserts month, day and year
*!*			WordApp.selection.MoveDown(wdLine,1)
*!*			WordApp.selection.insertbefore(mdy(date()))
*!*			******

*!*			******** This completes the seventh page

*!*			******** New Page 8
*!*			******
*!*			* This is the installation responsibilities page
*!*			* There is nothing on this page to edit or add.
*!*			******

*!*			******** This completes the eighth page

*!*			******** New Page 9
*!*			******** Old Page 8

*!*			******
*!*			* Paragraph 308 is the Date at the end of the
*!*			* Limited Warranty & Miscellaneous Terms Page
*!*			* go to Paragraph 308 and insert information
*!*			WordApp.activedocument.paragraphs(308).range.select
*!*			WordApp.selection.EndKey
*!*			* inserts month, day and year
*!*			WordApp.selection.insertafter(mdy(date()))
*!*			******

*!*			******** This completes the ninth page

*!*			*********************************
*!*			* Enter information on Price Page
*!*			*********************************
*!*			* Continue with New Page 4
*!*			* Continue with Old Page 3

*!*			******
*!*			* fill in table with group prices
*!*			MyGroupTable = WordApp.ActiveDocument.tables(2)
*!*			tcount = 2
*!*			* paragraph count needed to find Option - Paragraph
*!*			pcount = 0
*!*			* clastgroup must have a value of at least "A"
*!*			store "A" to clastgroup
*!*			do while alltrim(a.group) # '$'
*!*				if tcount => 6 .and. alltrim(a.group) # '$'
*!*					MyGroupTable.rows.add
*!*					* each row adds 8 paragraphs
*!*					pcount = pcount + 8
*!*				endif
*!*				With MyGroupTable
*!*					.cell(tcount,1).range.insertafter(a.group)
*!*					.cell(tcount,2).range.insertafter(alltrim(a.gname1) + " " + alltrim(a.gname2))
*!*					do numberformat with 3,"a.material"
*!*					do numberformat with 4,"a.install"
*!*					do numberformat with 5,"a.rtotal"
*!*					do numberformat with 6,"a.pc_tech"
*!*					do numberformat with 7,"a.total"
*!*				EndWith
*!*				tcount = tcount + 1
*!*				skip 1 in a
*!*				* check next group
*!*				if alltrim(a.group) # "$"
*!*					store alltrim(a.group) to clastgroup
*!*				endif

*!*				* Get recno() for updating progress report
*!*				store recno("a") to nrecnum
*!*				progressreport.refresh

*!*			enddo
*!*			******

*!*			******
*!*			* Paragraph 88 is the "Total For Groups :" line
*!*			*  before any group lines are added which increase the value of pcount
*!*			*  go to Paragraph 88 plus pcount and insert information
*!*			WordApp.activedocument.paragraphs(88+pcount).range.select
*!*			WordApp.selection.Endkey
*!*			if clastgroup = "A"
*!*				WordApp.selection.TypeBackspace
*!*				WordApp.selection.insertafter("A:")
*!*			else
*!*				WordApp.selection.TypeBackspace
*!*				WordApp.selection.insertafter("A - " + clastgroup + ":")
*!*			endif
*!*			******

*!*			******
*!*			* totals changed to a table
*!*			* fill in totals
*!*			* Tables(3) and insert/edit information
*!*			tcount = 1
*!*			MyTotalTable = WordApp.ActiveDocument.tables(3)
*!*			With MyTotalTable
*!*				do numberformat with 3,"a.material"
*!*				tcount = tcount + 1
*!*				do numberformat with 3,"a.install"
*!*				tcount = tcount + 1
*!*				do numberformat with 3,"a.rtotal"
*!*				tcount = tcount + 1
*!*				do numberformat with 3,"a.pc_tech"
*!*				tcount = tcount + 1
*!*				do numberformat with 3,"a.total"
*!*			EndWith
*!*			******

*!*			if a.install = 0 .and. a.rtotal = 0
*!*				* revise note if there is no installation and no r's prices
*!*				* go to Paragraph 110+pcount
*!*				WordApp.activedocument.paragraphs(110+pcount).range.select
*!*				WordApp.selection.Find.ClearFormatting
*!*				With WordApp.selection.Find
*!*					.Text = "include f"
*!*					.Forward = .t.
*!*					.Wrap = wdFindContinue
*!*					.Format = .f.
*!*					.MatchCase = .t.
*!*					.MatchWholeWord = .t.
*!*					.MatchWildcards = .f.
*!*					.MatchSoundsLike = .f.
*!*					.MatchAllWordForms = .f.
*!*				EndWith
*!*				WordApp.selection.Find.Execute

*!*				* erases the template information include
*!*				WordApp.selection.Delete
*!*				WordApp.selection.insertafter("include installation, installation materials, f")

*!*				* add two notes
*!*				WordApp.activedocument.paragraphs(110+pcount).range.select
*!*				WordApp.selection.EndKey
*!*				WordApp.selection.TypeParagraph
*!*				WordApp.selection.insertafter('An Automated Conveyor Systems, Inc. supervisor can be provided on a time and expenses basis to assist plant personnel with installation. Please see the enclosed "Installation Services Sheet" for applicable rates and expenses.')
*!*				WordApp.selection.EndKey
*!*				WordApp.selection.TypeParagraph
*!*				WordApp.selection.insertafter('A recommended list of installation materials will be provided for your use.')

*!*				* add 2 to paragraph count
*!*				pcount = pcount + 2

*!*				* go back to page 7 (General Notes) and remove notes not needed.
*!*				* 5 and 7 thru 12
*!*				WordApp.activedocument.paragraphs(128+pcount+5).range.select
*!*				WordApp.selection.Delete
*!*				* this statement deletes the next five notes without selecting them
*!*				WordApp.activedocument.range(WordApp.activedocument.paragraphs(128+pcount+6).range.start,WordApp.activedocument.paragraphs(133+pcount+6).range.end).delete
*!*			endif

*!*			******
*!*			* Need a variable to find the paragraph to start adding returns until the options,
*!*			*  if there is more than 1, are moved to a new page.
*!*			* at this point the pcount variable will be used to determine the correct paragraph.
*!*			ptircount = 112 + pcount - 2
*!*			******

*!*			******
*!*			* fill in option 01 information if required
*!*			skip 1 in a
*!*			loptionnewpage = .f.
*!*			if eof() = .f.
*!*				* variable to move the options to a new page
*!*				loptionnewpage = .t.
*!*				* always add the group name
*!*				* go to Paragraph 112+pcount and insert/edit information won't work paragraph will change
*!*				WordApp.activedocument.paragraphs(112+pcount).range.select
*!*				Wordapp.selection.EndKey
*!*				WordApp.selection.insertafter(alltrim(a.gname1) + " " + alltrim(a.gname2))
*!*				if isblank(option_ref) = .t.    && no reference
*!*					if a.material = 0 and a.install = 0 and a.rtotal = 0 and a.pc_tech = 0
*!*						* template information must be removed then add a paragraph
*!*						WordApp.selection.Movedown(wdLine,1)
*!*						WordApp.selection.HomeKey(wdline,1) &&    Selection.HomeKey Unit:=wdLine
*!*						WordApp.selection.Delete(wdWord,21)
*!*						* all template information cleared at this point
*!*					else
*!*						* move cursor to the beginning of the Add Material statement
*!*						WordApp.selection.MoveRight(wdCharacter,4)
*!*						if a.material = 0    && remove the line for material
*!*							* at this time the cursor is somewhere on the Add Material line thus delete
*!*							WordApp.selection.StartOf
*!*							WordApp.selection.TypeBackspace
*!*							WordApp.selection.TypeBackspace
*!*							WordApp.selection.Delete(wdCharacter,15)
*!*							* cursor is at the beginning
*!*						else
*!*							if a.material < 0
*!*							    WordApp.selection.Delete(wdCharacter,3)
*!*								WordApp.selection.insertbefore("Deduct")
*!*							endif
*!*							WordApp.selection.EndKey
*!*							do numberformatopt with "a.material"
*!*							WordApp.selection.MoveRight(wdCharacter,2)
*!*							* cursor is at the beginning
*!*						endif
*!*						* at his point the cursor is at the beginning of the line
*!*						if a.install = 0    && remove the line for install
*!*							* at this time the cursor is at the beginning of the line
*!*							WordApp.selection.Delete(wdCharacter,21)
*!*							* at this time the cursor is at the beginning of the ADD TOTAL line
*!*						else
*!*							if a.install < 0
*!*							    WordApp.selection.MoveRight(wdCharacter,2)
*!*							    WordApp.selection.Delete(wdCharacter,3)
*!*								WordApp.selection.insertbefore("Deduct")
*!*							else
*!*							endif
*!*							WordApp.selection.EndKey
*!*							do numberformatopt with "a.install"
*!*							WordApp.selection.MoveRight(wdCharacter,2)
*!*							* cursor is at the beginning of the ADD TOTAL line
*!*						endif
*!*						* at his point the cursor is at the beginning of the ADD TOTAL line
*!*						* add rtotal if necessary
*!*						if a.rtotal # 0    && don't add the line for rtotal
*!*					    	WordApp.selection.TypeText(Chr(9)+Chr(9))
*!*							if a.rtotal < 0
*!*								WordApp.selection.TypeText("Deduct Relocation, Rework & Removal"+Chr(9))
*!*							else
*!*								WordApp.selection.TypeText("Add Relocation, Rework & Removal"+Chr(9))
*!*							endif
*!*					    	do numberformatopt with "a.rtotal"
*!*							WordApp.selection.MoveRight
*!*					    	WordApp.selection.TypeText(Chr(11))
*!*						endif
*!*						* if rtotal not used then the cursor is at the beginning of the ADD TOTAL line
*!*						* if rtotal is used then the cursor is at the beginning of the ADD TOTAL line

*!*						* at his point the cursor is at the beginning of the ADD TOTAL line
*!*						* add pc_tech if necessary
*!*						if a.pc_tech # 0    && don't add the line for pc_tech
*!*					    	WordApp.selection.TypeText(Chr(9)+Chr(9))
*!*							if a.pc_tech < 0
*!*								WordApp.selection.TypeText("Deduct PLC Start Up"+Chr(9)+Chr(9))
*!*							else
*!*								WordApp.selection.TypeText("Add PLC Start Up"+Chr(9)+Chr(9))
*!*							endif
*!*					    	do numberformatopt with "a.pc_tech"
*!*							WordApp.selection.MoveRight
*!*					    	WordApp.selection.TypeText(Chr(11))
*!*						endif

*!*						* if pc_tech not used then the cursor is at the beginning of the
*!*						* ADD TOTAL line
*!*						* if pc_tech is used then the cursor is at the beginning of the
*!*						* ADD TOTAL line

*!*						if a.total < 0
*!*						    WordApp.selection.MoveRight(wdCharacter,5)
*!*							WordApp.selection.insertbefore("DEDUCT")
*!*							* Deduct must be written within the paragraph
*!*							*  thus you must go back to the start of and delete the word add
*!*							WordApp.selection.StartOf
*!*						    WordApp.selection.Delete(wdCharacter,3)
*!*							WordApp.selection.EndKey
*!*							do numberformatopt with "a.total"
*!*							* a.total was formatted and must be written within the paragraph
*!*							* thus you must go back to the start of and delete the '$'
*!*							WordApp.selection.StartOf
*!*						    WordApp.selection.Delete(wdCharacter,1)
*!*						else
*!*							WordApp.selection.EndKey
*!*							do numberformatopt with "a.total"
*!*							WordApp.selection.StartOf
*!*						    WordApp.selection.Delete(wdCharacter,1)
*!*						endif
*!*					endif   && material, install, rtotal, and pc_tech all equal zero if
*!*				else    && reference used
*!*					* add to group name
*!*					WordApp.selection.insertafter(" IN LIEU OF GROUP " + alltrim(option_ref))
*!*					if (a.material - g.material) = 0 and (a.install - g.install) = 0 and ;
*!*					   (a.rtotal - g.rtotal) = 0 and (a.pc_tech - g.pc_tech) = 0
*!*						* template information must be removed then add a paragraph
*!*						WordApp.selection.Movedown(wdLine,1)
*!*						WordApp.selection.HomeKey(wdline,1) &&    Selection.HomeKey Unit:=wdLine
*!*						WordApp.selection.Delete(wdWord,21)
*!*						* all template information cleared at this point
*!*					else
*!*						* move cursor to the beginning of the Add Material statement
*!*						WordApp.selection.MoveRight(wdCharacter,4)
*!*						if g.material = 0    && remove the line for material
*!*							* at this time the cursor is somewhere on the Add Material line thus delete
*!*							WordApp.selection.StartOf
*!*							WordApp.selection.TypeBackspace
*!*							WordApp.selection.TypeBackspace
*!*							WordApp.selection.Delete(wdCharacter,15)
*!*							* cursor is at the beginning
*!*						else
*!*							if a.material - g.material < 0
*!*							    WordApp.selection.Delete(wdCharacter,3)
*!*								WordApp.selection.insertbefore("Deduct")
*!*							endif
*!*							WordApp.selection.EndKey
*!*							do numberformatopt with "a.material - g.material"
*!*							WordApp.selection.MoveRight(wdCharacter,2)
*!*							* cursor is at the beginning
*!*						endif
*!*						* at his point the cursor is at the beginning of the line
*!*						if a.install - g.install = 0    && remove the line for install
*!*							* at this time the cursor is at the beginning of the line
*!*							WordApp.selection.Delete(wdCharacter,21)
*!*							* at this time the cursor is at the beginning of the ADD TOTAL line
*!*						else
*!*							if a.install - g.install < 0
*!*							    WordApp.selection.MoveRight(wdCharacter,2)
*!*							    WordApp.selection.Delete(wdCharacter,3)
*!*								WordApp.selection.insertbefore("Deduct")
*!*							else
*!*							endif
*!*							WordApp.selection.EndKey
*!*							do numberformatopt with "a.install-g.install"
*!*							WordApp.selection.MoveRight(wdCharacter,2)
*!*							* cursor is at the beginning of the ADD TOTAL line
*!*						endif
*!*						* at his point the cursor is at the beginning of the ADD TOTAL line
*!*						* add rtotal if necessary
*!*						if a.rtotal - g.rtotal # 0    && don't add the line for rtotal
*!*					    	WordApp.selection.TypeText(Chr(9)+Chr(9))
*!*							if a.rtotal - g.rtotal < 0
*!*								WordApp.selection.TypeText("Deduct Relocation, Rework & Removal"+Chr(9))
*!*							else
*!*								WordApp.selection.TypeText("Add Relocation, Rework & Removal"+Chr(9))
*!*							endif
*!*					    	do numberformatopt with "a.rtotal - g.rtotal"
*!*							WordApp.selection.MoveRight
*!*					    	WordApp.selection.TypeText(Chr(11))
*!*						endif
*!*						* if rtotal not used then the cursor is at the beginning of the ADD TOTAL line
*!*						* if rtotal is used then the cursor is at the beginning of the ADD TOTAL line

*!*						* at his point the cursor is at the beginning of the ADD TOTAL line
*!*						* add pc_tech if necessary
*!*						if a.pc_tech - g.pc_tech # 0    && don't add the line for pc_tech
*!*					    	WordApp.selection.TypeText(Chr(9)+Chr(9))
*!*							if a.pc_tech - g.pc_tech < 0
*!*								WordApp.selection.TypeText("Deduct PLC Start Up"+Chr(9)+Chr(9))
*!*							else
*!*								WordApp.selection.TypeText("Add PLC Start Up"+Chr(9)+Chr(9))
*!*							endif
*!*					    	do numberformatopt with "a.pc_tech - g.pc_tech"
*!*							WordApp.selection.MoveRight
*!*					    	WordApp.selection.TypeText(Chr(11))
*!*						endif

*!*						* if pc_tech not used then the cursor is at the beginning of the
*!*						* ADD TOTAL line
*!*						* if pc_tech is used then the cursor is at the beginning of the
*!*						* ADD TOTAL line
*!*						if a.total - g.total < 0
*!*						    WordApp.selection.MoveRight(wdCharacter,5)
*!*							WordApp.selection.insertbefore("DEDUCT")
*!*							* Deduct must be written within the paragraph
*!*							*  thus you must go back to the start of and delete the word add
*!*							WordApp.selection.StartOf
*!*						    WordApp.selection.Delete(wdCharacter,3)
*!*							WordApp.selection.EndKey
*!*							do numberformatopt with "a.total - g.total"
*!*							* g.total was formatted and must be written within the paragraph
*!*							* thus you must go back to the start of and delete the '$'
*!*							WordApp.selection.StartOf
*!*						    WordApp.selection.Delete(wdCharacter,1)
*!*						else
*!*							WordApp.selection.EndKey
*!*							do numberformatopt with "a.total - g.total"
*!*							WordApp.selection.StartOf
*!*						    WordApp.selection.Delete(wdCharacter,1)
*!*						endif
*!*					endif   && material, install, rtotal, and pc_tech all equal zero if
*!*				endif    && option reference
*!*			endif
*!*			if recno("a") <= reccount("a")
*!*				skip 1 in a
*!*			endif

*!*			* Get recno() for updating progress report
*!*			store recno("a") to nrecnum
*!*			progressreport.refresh

*!*			******

*!*			******
*!*			* Add additional options if required
*!*			do while eof() = .f.
*!*				if eof() = .f.
*!*					* go to the end then add a paragraph
*!*					* always add the group name
*!*			 	    WordApp.selection.EndKey
*!*					WordApp.selection.TypeParagraph
*!*					* Enter information
*!*					if isblank(option_ref) = .t.    && no reference
*!*						WordApp.selection.TypeText("Option " + a.group + " - " + alltrim(a.gname1) + " " + alltrim(a.gname2))
*!*					else
*!*						* add to group name
*!*						WordApp.selection.TypeText("Option " + a.group + " - " + alltrim(a.gname1) + " " + alltrim(a.gname2) + " IN LIEU OF GROUP " + alltrim(option_ref))
*!*					endif
*!*					* go to end then add a paragraph
*!*			 	    WordApp.selection.EndKey
*!*			 	    WordApp.selection.TypeParagraph
*!*					* now go back and format previous paragraph
*!*					wordapp.activedocument.paragraphs(114+pcount).range.select
*!*					With WordApp.selection
*!*						With .Shading
*!*							.Texture = wdTexture5Percent
*!*							.ForegroundPatternColorIndex = wdAuto
*!*							.BackgroundPatternColorIndex = wdWhite
*!*						EndWith
*!*						With .Font
*!*							.Underline = wdUnderlineSingle
*!*							.Bold = .t.
*!*						EndWith
*!*					EndWith
*!*					* move down the the paragraph that was added and insert information
*!*					WordApp.selection.MoveDown(wdLine,1)
*!*					* determine if a reference has been used
*!*					if isblank(option_ref) = .t.    && no reference
*!*						if a.material = 0 .and. a.install = 0 .and. a.rtotal = 0 .and. a.pc_tech = 0
*!*							* just add a line between titles
*!*						else
*!*							if a.material = 0
*!*								* don't add the line for material
*!*							else
*!*								* add information
*!*								if a.material < 0
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Deduct Material" + Chr(9) + Chr(9))
*!*									WordApp.selection.EndKey
*!*									do numberformatopt with "a.material"
*!*								else
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Add Material" + Chr(9) + Chr(9))
*!*									WordApp.selection.EndKey
*!*									do numberformatopt with "a.material"
*!*								endif
*!*							endif
*!*							if a.install = 0    && don't add the line for installation
*!*								* don't add the line for installation
*!*							else
*!*								* add information
*!*								if a.install < 0
*!*									WordApp.selection.EndKey
*!*									iif(a.material # 0,WordApp.selection.Typetext(Chr(11)),"")
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Deduct Installation" + Chr(9) + Chr(9))
*!*									WordApp.selection.EndKey
*!*									do numberformatopt with "a.install"
*!*								else
*!*									WordApp.selection.EndKey
*!*									iif(a.material # 0,WordApp.selection.Typetext(Chr(11)),"")
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Add Installation" + Chr(9) + Chr(9))
*!*									WordApp.selection.EndKey
*!*									do numberformatopt with "a.install"
*!*								endif
*!*							endif
*!*							if a.rtotal # 0    && don't add the line for rtotal
*!*								* there are four cases here to determine if a chr(11) is used
*!*								* 1 material=0 and installation=0  not needed
*!*								* 2 material=0 and installation#0  needed
*!*								* 3 material#0 and installation=0  needed
*!*								* 4 material#0 and installation#0  needed
*!*								WordApp.selection.EndKey
*!*								if a.material = 0 .and. a.install = 0
*!*								else
*!*									WordApp.selection.Typetext(Chr(11))
*!*								endif
*!*								* add information
*!*								if a.rtotal < 0
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Deduct Relocation, Rework & Removal" + Chr(9))
*!*								else
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Add Relocation, Rework & Removal" + Chr(9))
*!*								endif
*!*								WordApp.selection.EndKey
*!*								do numberformatopt with "a.rtotal"
*!*							endif
*!*							if a.pc_tech # 0    && don't add the line for pc_tech
*!*								* there are eight cases here to determine if a chr(11) is used
*!*								* 1 material=0 and installation=0 and rtotal=0  not needed
*!*								* 2 material=0 and installation=0 and rtotal#0  needed
*!*								* 3 material=0 and installation#0 and rtotal=0  needed
*!*								* 4 material=0 and installation#0 and rtotal#0  needed
*!*								* 5 material#0 and installation=0 and rtotal=0  needed
*!*								* 6 material#0 and installation=0 and rtotal#0  needed
*!*								* 7 material#0 and installation#0 and rtotal=0  needed
*!*								* 8 material#0 and installation#0 and rtotal#0  needed
*!*								WordApp.selection.EndKey
*!*								if a.material = 0 .and. a.install = 0 .and. a.rtotal = 0
*!*								else
*!*									WordApp.selection.Typetext(Chr(11))
*!*								endif
*!*								* add information
*!*								if a.pc_tech < 0
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Deduct PLC Start Up" + Chr(9) + Chr(9))
*!*								else
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"PLC Start Up" + Chr(9) + Chr(9))
*!*								endif
*!*								WordApp.selection.EndKey
*!*								do numberformatopt with "a.pc_tech"
*!*							endif
*!*							WordApp.selection.EndKey
*!*							WordApp.selection.TypeText(Chr(11) + Chr(9) + Chr(9))
*!*							if a.total < 0
*!*								WordApp.selection.insertafter("DEDUCT TOTAL")
*!*							else
*!*								WordApp.selection.insertafter("ADD TOTAL")
*!*							endif
*!*							* selection must be formatted to bold
*!*							WordApp.selection.Font.Bold = .t.
*!*							WordApp.selection.EndKey
*!*							WordApp.selection.TypeText(Chr(9) + Chr(9))
*!*							do numberformatopt with "a.total"
*!*							* selection must be formatted to add double underline
*!*							WordApp.selection.Font.Bold = .t.
*!*							WordApp.selection.Font.Underline = wdUnderlineDouble
*!*							WordApp.selection.EndKey
*!*						endif
*!*					else
*!*						if (a.material - g.material) = 0 .and. (a.install - g.install) = 0 .and. ;
*!*						   (a.rtotal -g.rtotal) = 0 .and. (a.pc_tech - g.pc_tech) = 0
*!*							* just add a line between titles
*!*						else
*!*							if a.material - g.material = 0
*!*								* don't add the line for material
*!*							else
*!*								* add information
*!*								if a.material - g.material < 0
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Deduct Material" + Chr(9) + Chr(9))
*!*									WordApp.selection.EndKey
*!*									do numberformatopt with "a.material - g.material"
*!*								else
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Add Material" + Chr(9) + Chr(9))
*!*									WordApp.selection.EndKey
*!*									do numberformatopt with "a.material - g.material"
*!*								endif
*!*							endif
*!*							if a.install - g.install = 0    && don't add the line for installation
*!*								* don't add the line for installation
*!*							else
*!*								* add information
*!*								if a.install - g.install < 0
*!*									WordApp.selection.EndKey
*!*									iif(a.material # 0,WordApp.selection.Typetext(Chr(11)),"")
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Deduct Installation" + Chr(9) + Chr(9))
*!*									WordApp.selection.EndKey
*!*									do numberformatopt with "a.install - g.install"
*!*								else
*!*									WordApp.selection.EndKey
*!*									iif(a.material # 0,WordApp.selection.Typetext(Chr(11)),"")
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Add Installation" + Chr(9) + Chr(9))
*!*									WordApp.selection.EndKey
*!*									do numberformatopt with "a.install - g.install"
*!*								endif
*!*							endif
*!*							if a.rtotal - g.rtotal # 0    && don't add the line for rtotal
*!*								* there are four cases here to determine if a chr(11) is used
*!*								* 1 material=0 and installation=0  not needed
*!*								* 2 material=0 and installation#0  needed
*!*								* 3 material#0 and installation=0  needed
*!*								* 4 material#0 and installation#0  needed
*!*								WordApp.selection.EndKey
*!*								if (a.material - g.material) = 0 .and. (a.install - g.install) = 0
*!*								else
*!*									WordApp.selection.Typetext(Chr(11))
*!*								endif
*!*								* add information
*!*								if a.rtotal - g.rtotal < 0
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Deduct Relocation, Rework & Removal" + Chr(9))
*!*								else
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Add Relocation, Rework & Removal" + Chr(9))
*!*								endif
*!*								WordApp.selection.EndKey
*!*								do numberformatopt with "a.rtotal - g.rtotal"
*!*							endif
*!*							if a.pc_tech - g.pc_tech # 0    && don't add the line for pc_tech
*!*								* there are eight cases here to determine if a chr(11) is used
*!*								* 1 material=0 and installation=0 and rtotal=0  not needed
*!*								* 2 material=0 and installation=0 and rtotal#0  needed
*!*								* 3 material=0 and installation#0 and rtotal=0  needed
*!*								* 4 material=0 and installation#0 and rtotal#0  needed
*!*								* 5 material#0 and installation=0 and rtotal=0  needed
*!*								* 6 material#0 and installation=0 and rtotal#0  needed
*!*								* 7 material#0 and installation#0 and rtotal=0  needed
*!*								* 8 material#0 and installation#0 and rtotal#0  needed
*!*								WordApp.selection.EndKey
*!*								if a.material - g.material = 0 .and. a.install - g.install = 0 .and. a.rtotal - g.rtotal = 0
*!*								else
*!*									WordApp.selection.Typetext(Chr(11))
*!*								endif
*!*								* add information
*!*								if a.pc_tech - g.pc_tech < 0
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Deduct PLC Start Up" + Chr(9) + Chr(9))
*!*								else
*!*									WordApp.selection.TypeText(Chr(9)+Chr(9)+"Add PLC Start Up" + Chr(9) + Chr(9))
*!*								endif
*!*								WordApp.selection.EndKey
*!*								do numberformatopt with "a.pc_tech - g.pc_tech"
*!*							endif
*!*							WordApp.selection.EndKey
*!*							WordApp.selection.TypeText(Chr(11) + Chr(9) + Chr(9))
*!*							if a.total - g.total < 0
*!*								WordApp.selection.insertafter("DEDUCT TOTAL")
*!*							else
*!*								WordApp.selection.insertafter("ADD TOTAL")
*!*							endif
*!*							* selection must be formatted to bold
*!*							WordApp.selection.Font.Bold = .t.
*!*							WordApp.selection.EndKey
*!*							WordApp.selection.TypeText(Chr(9) + Chr(9))
*!*							do numberformatopt with "a.total - g.total"
*!*							* selection must be formatted to add double underline
*!*							WordApp.selection.Font.Bold = .t.
*!*							WordApp.selection.Font.Underline = wdUnderlineDouble
*!*							WordApp.selection.EndKey
*!*						endif
*!*					endif
*!*					* do again
*!*					* identify new paragraph before continuing
*!*					pcount = pcount + 2
*!*				endif
*!*				skip 1 in a

*!*				* Get recno() for updating progress report
*!*				store recno("a") to nrecnum
*!*				progressreport.refresh

*!*			enddo
*!*			******

*!*			******
*!*			* there was more than one option group
*!*			if loptionnewpage = .t.
*!*				* this selects the paragraph to add returns to force the option to their own page
*!*				WordApp.activedocument.paragraphs(ptircount).range.select
*!*				WordApp.selection.EndKey
*!*				WordApp.selection.TypeParagraph

*!*				* add 1 to ptircount
*!*				ptircount = ptircount + 1

*!*				* at this point you are at the top of the next page
*!*				NewRange = WordApp.activeDocument.paragraphs(ptircount).Range
*!*				WordApp.activeDocument.Tables.Add(NewRange,4,2)

*!*				* set style, font, paragraphformat and border settings
*!*				WordApp.activeDocument.Tables(4).range.select
*!*				With WordApp.selection
*!*					.Style = WordApp.activeDocument.Styles("Body Text")
*!*					With .Font
*!*						.Bold = .t.
*!*						.Italic = .f.
*!*						.Name = "Arial"
*!*						.Size = 11
*!*						.Smallcaps = .t.
*!*					EndWith
*!*					With .ParagraphFormat
*!*				        .SpaceBefore = 0
*!*				        .SpaceAfter = 0
*!*			        EndWith
*!*			        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
*!*			        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
*!*			        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
*!*			        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
*!*			        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
*!*			        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
*!*				EndWith
*!*				* fill in table with information
*!*				MyHeaderTable = WordApp.ActiveDocument.tables(4)
*!*				With MyHeaderTable
*!*					.cell(1,1).range.insertafter("Price Proposal")
*!*					.cell(2,1).range.insertafter(ccustomer)
*!*					.cell(3,1).range.insertafter(clocation)
*!*					.cell(3,2).range.insertafter(mdy(date()))
*!*					.cell(4,1).range.insertafter("Quotation #" + cquote)
*!*					.cell(4,2).range.insertafter("PAGE 2")
*!*				EndWith

*!*				* change paragraphs to right alignment
*!*				WordApp.activedocument.paragraphs(ptircount+7).range.select
*!*				With WordApp.selection.ParagraphFormat
*!*					.Alignment = wdAlignParagraphRight
*!*				EndWith
*!*				WordApp.activedocument.paragraphs(ptircount+10).range.select
*!*				With WordApp.selection.ParagraphFormat
*!*					.Alignment = wdAlignParagraphRight
*!*				EndWith

*!*				* change style of paragraph
*!*				WordApp.activedocument.paragraphs(ptircount+11).range.select
*!*				With WordApp.selection
*!*					.Style = WordApp.activeDocument.Styles("Body Text")
*!*				EndWith

*!*				* move to another page
*!*				WordApp.activedocument.Paragraphs(ptircount).PageBreakBefore = .t.
*!*			endif
*!*			******

*!*			* release form
*!*			progressreport.release

*!*			* Make Visible
*!*			WordApp.Application.Visible = .t.

*!*			* Save document
*!*			* document directory
*!*			cdocdir = "\\ACS-NT4-FS1\SALESSRV\WORK IN PROGRESS\DOCUMENTS\WORD\"
*!*			ctabledir = "\\ACS-NT4-FS1\SALESSRV\WORK IN PROGRESS\DRAWINGS\QUOTE\"
*!*			ccopytodir = "\\ACS-NT4-FS1\ADM\WORK IN PROGRESS\DOCUMENTS\WORD\QUOTE\"
*!*			* save in tabledir for now
*!*			* WordDoc.Saveas(ctabledir + cdrawing + 'PP')
*!*			*!*	WordDoc.Saveas(ccopytodir + cdrawing + 'PP')

*!*			* Quit Application, Set to nil
*!*			* WordApp.quit
*!*			* WordApp = "nil"

*!*			* Print the Address Label
*!*			******
*!*			* Create Application & Document
*!*			WordApp = CreateObject("Word.application.8")
*!*			*!*	 * Open document template
*!*			*!*	WordDoc = WordApp.Documents.open('c:\my documents\AddressLabelspp.dot')
*!*			* Open edits the template and does not allow saving as a .doc file thus
*!*			* Add a new document using the AddressLabelspp template
*!*			*                          .Add(Template:                   , NewTemplate)
*!*			* WordDoc = WordApp.Documents.Add('c:\my documents\AddressLabelspp.dot', .f.)
*!*			* When the network template is used the above statement will change to
*!*			WordDoc = WordApp.Documents.Add('G:\Word Templates\Quote Templates\AddressLabelspp.dot', .f.)
*!*			* New caption
*!*			WordApp.Caption = "ACS Address Labels"
*!*			* Make Visible
*!*			WordApp.Application.Visible = .t.

*!*			WordApp.activedocument.paragraphs(8).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.TypeText("TO:")
*!*			WordApp.selection.Delete(wdCharacter,1)
*!*			WordApp.selection.MoveRight(wdCharacter,1)
*!*			WordApp.selection.TypeText(cppsal+" "+cppnamef+" "+cppnamel)
*!*			WordApp.selection.Delete(wdCharacter,1)
*!*			WordApp.activedocument.paragraphs(9).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.MoveRight(wdCharacter,1)
*!*			WordApp.selection.TypeText(cppcustomer)
*!*			WordApp.selection.Delete(wdCharacter,1)
*!*			WordApp.activedocument.paragraphs(10).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.MoveRight(wdCharacter,1)
*!*			WordApp.selection.TypeText(cppaddress)
*!*			WordApp.selection.Delete(wdCharacter,1)
*!*			WordApp.activedocument.paragraphs(11).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.MoveRight(wdCharacter,1)
*!*			WordApp.selection.TypeText(cpplocation + " " + cppzip)
*!*			WordApp.selection.Delete(wdCharacter,1)

*!*			* a copy goes to
*!*			* insert the copies for names
*!*			if lgotcopypernumber = .t.
*!*				if len(tempselectarray[1,1]) > 0
*!*					* Select and insert name(s) that will get a copy
*!*					* WordApp.selection.MoveRight(wdCharacter,31)

*!*					* print their names - limit set at 5 for now
*!*					* this will created one whole page of labels
*!*					for f = 1 to 5
*!*						if len(tempselectarray[f,1]) > 0
*!*							if f > 0
*!*								* variable npc is numerical paragraph count
*!*								do case
*!*								case f = 1
*!*									npc = 28
*!*								case f = 2
*!*									npc = 39
*!*								case f = 3
*!*									npc = 57
*!*								case f = 4
*!*									npc = 67
*!*								case f = 5
*!*									npc = 77
*!*								endcase

*!*								WordApp.activedocument.paragraphs(npc).range.select
*!*								WordApp.selection.StartOf
*!*								if f = 1 .or. f = 3 .or. f = 5
*!*									WordApp.selection.MoveRight(wdCharacter,1)
*!*								endif
*!*								WordApp.selection.TypeText("TO:")
*!*								WordApp.selection.Delete(wdCharacter,1)
*!*								WordApp.selection.MoveRight(wdCharacter,1)
*!*								WordApp.selection.TypeText(tempselectarray[f,2]+" "+tempselectarray[f,3]+" "+tempselectarray[f,4])
*!*								WordApp.selection.Delete(wdCharacter,1)
*!*								* company
*!*								WordApp.activedocument.paragraphs(npc+1).range.select
*!*								WordApp.selection.StartOf
*!*								WordApp.selection.MoveRight(wdCharacter,1)
*!*								if f = 1 .or. f = 3 .or. f = 5
*!*									WordApp.selection.MoveRight(wdCharacter,1)
*!*								endif
*!*								WordApp.selection.TypeText(tempselectarray[f,6])
*!*								WordApp.selection.Delete(wdcharacter,1)
*!*								* street address
*!*								WordApp.activedocument.paragraphs(npc+2).range.select
*!*								WordApp.selection.StartOf
*!*								WordApp.selection.MoveRight(wdCharacter,1)
*!*								if f = 1 .or. f = 3 .or. f = 5
*!*									WordApp.selection.MoveRight(wdCharacter,1)
*!*								endif
*!*								WordApp.selection.TypeText(tempselectarray[f,7])
*!*								WordApp.selection.Delete(wdcharacter,1)
*!*								* street address
*!*								WordApp.activedocument.paragraphs(npc+3).range.select
*!*								WordApp.selection.StartOf
*!*								WordApp.selection.MoveRight(wdCharacter,1)
*!*								if f = 1 .or. f = 3 .or. f = 5
*!*									WordApp.selection.MoveRight(wdCharacter,1)
*!*								endif
*!*								WordApp.selection.TypeText(tempselectarray[f,8])
*!*								WordApp.selection.Delete(wdcharacter,1)
*!*							endif
*!*						else
*!*							* do nothing
*!*						endif
*!*					endfor
*!*				else
*!*					* do nothing
*!*				endif
*!*			endif

*!*			* Print the Quote Folder Labels
*!*			******
*!*			* Create Application & Document
*!*			WordApp = CreateObject("Word.application.8")
*!*			*!*	 * Open document template
*!*			*!*	WordDoc = WordApp.Documents.open('c:\my documents\ACSQtLabelspp.dot')
*!*			* Open edits the template and does not allow saving as a .doc file thus
*!*			* Add a new document using the AddressLabelspp template
*!*			*                          .Add(Template:                   , NewTemplate)
*!*			* WordDoc = WordApp.Documents.Add('c:\my documents\ACSQtLabelspp.dot', .f.)
*!*			* When the network template is used the above statement will change to
*!*			WordDoc = WordApp.Documents.Add('G:\Word Templates\Quote Templates\ACSQtLabelspp.dot', .f.)
*!*			* New caption
*!*			WordApp.Caption = "ACS Quote Labels"
*!*			* Make Visible
*!*			WordApp.Application.Visible = .t.

*!*			* WordApp.selection.TypeParagraph = enter
*!*			WordApp.activedocument.paragraphs(1).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.TypeText("Quotation #"+ cquote)
*!*			WordApp.selection.TypeParagraph
*!*			WordApp.selection.TypeText(ccustomer + ", " + clocation)
*!*			WordApp.selection.TypeParagraph
*!*			WordApp.selection.TypeText(ctitle)
*!*			WordApp.selection.TypeParagraph
*!*			WordApp.selection.TypeText(dtoc(date()))
*!*			* second label
*!*			WordApp.activedocument.paragraphs(6).range.select
*!*			WordApp.selection.StartOf
*!*			WordApp.selection.TypeText("Quotation #"+ cquote)
*!*			WordApp.selection.TypeParagraph
*!*			WordApp.selection.TypeText(ccustomer + ", " + clocation)
*!*			WordApp.selection.TypeParagraph
*!*			WordApp.selection.TypeText(ctitle)
*!*			WordApp.selection.TypeParagraph
*!*			WordApp.selection.TypeText(dtoc(date()))
*!*		else
*!*			set message to "All the tables were not opened."
*!*			wait "All the tables were not opened." window at 6,10 timeout 3
*!*			set message to "The Quote Tracking or Quote History table was not opened. - Please try again."
*!*			wait "The Quote Tracking or Quote History table was not opened. - Please try again." window at 7,10 timeout 3
*!*		endif   && if upper(aliases = 'QUOTETRK', 'QUOTHIST', 'SALESSER', 'SALETERR'
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
*!*	* done in program
release lnoname
release cppsal, cppnamef, cppnamel, cpptitle
release cppcustomer, cpplocation, cppaddress, cppphone, cppfax, cppzip
release lgotcopypernumber, ncopyfromlist, lcopynoname
release lreffound
release cssi,csti
release cssname,csssname
release cstname
release nreccount, nrecnum

* erase the temporary tables
erase ctemppath + cdrawing + 'REF.DBF'
erase ctemppath + cdrawing + 'group.IDX'

* release the array
release tempselectarray

*-- EOP WRITEPPDOC

Procedure createltable
* create the table
use cmoldpath + 'moldmls' in a shared noupdate
copy structure to ctabledir+cdrawing+'L.DBF'
use ctabledir+cdrawing+'L.DBF' in a exclusive
* only the A and B work areas are being used
use ctabledir+cdrawing+'V.DBF' in c
select a
* do not create table if b.num_acs = 0
if c.NUM_ACS > 0
	append from mailtore for NUM_ACS = c.NUM_ACS
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
endproc

Procedure numberformat
Parameters cellpassed,fieldpassed
do case
	* seven digit
	case &fieldpassed. >  999999 .and. &fieldpassed. < 10000000
		.cell(tcount,cellpassed).range.insertafter("$" + ;
		substr(str(&fieldpassed.,8),1,2) + "," + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3))
	* six digit
	case &fieldpassed. >   99999 .and. &fieldpassed. < 1000000
		.cell(tcount,cellpassed).range.insertafter("$  " + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3))
	* five digit
	case &fieldpassed. >    9999 .and. &fieldpassed. < 100000
		.cell(tcount,cellpassed).range.insertafter("$   " + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3))
	* four digit
	case &fieldpassed. >     999 .and. &fieldpassed. < 10000
		.cell(tcount,cellpassed).range.insertafter("$    " + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3))
	* three digit
	case &fieldpassed. >      99 .and. &fieldpassed. < 1000
		.cell(tcount,cellpassed).range.insertafter("$         " + ;
		substr(str(&fieldpassed.,8),6,3))
	* two digit
	case &fieldpassed. >       9 .and. &fieldpassed. < 100
		.cell(tcount,cellpassed).range.insertafter("$          " + ;
		substr(str(&fieldpassed.,8),6,3))
	* one digit
	case &fieldpassed. =>      0 .and. &fieldpassed. < 10
		.cell(tcount,cellpassed).range.insertafter("$           " + ;
		substr(str(&fieldpassed.,8),6,3))
	* minus one digit
	case &fieldpassed. <       0 .and. &fieldpassed. > -10
		.cell(tcount,cellpassed).range.insertafter("$           -(" + ;
		substr(str(&fieldpassed.,8),8,1) + ")")
	* minus two digit
	case &fieldpassed. <      -9 .and. &fieldpassed. > -100
		.cell(tcount,cellpassed).range.insertafter("$       -(" + ;
		substr(str(&fieldpassed.,8),7,2) + ")")
	* minus three digit
	case &fieldpassed. <     -99 .and. &fieldpassed. > -1000
		.cell(tcount,cellpassed).range.insertafter("$     -(" + ;
		substr(str(&fieldpassed.,8),6,3) + ")")
	* minus four digit
	case &fieldpassed. <    -999 .and. &fieldpassed. > -10000
		.cell(tcount,cellpassed).range.insertafter("$  -(" + ;
		substr(str(&fieldpassed.,8),5,1) + "," + ;
		substr(str(&fieldpassed.,8),6,3) + ")")
	* minus five digit
	case &fieldpassed. <   -9999 .and. &fieldpassed. > -100000
		.cell(tcount,cellpassed).range.insertafter("$ -(" + ;
		substr(str(&fieldpassed.,8),4,2) + "," + ;
		substr(str(&fieldpassed.,8),6,3) + ")")
	* minus five digit
	case &fieldpassed. <  -99999 .and. &fieldpassed. > -1000000
		.cell(tcount,cellpassed).range.insertafter("$ -(" + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3) + ")")
	* minus six digit
	case &fieldpassed. < -999999 .and. &fieldpassed. > -10000000
		.cell(tcount,cellpassed).range.insertafter("$ -(" + ;
		substr(str(&fieldpassed.,8),2,1) + "," + ;
		substr(str(&fieldpassed.,8),3,3) + "," + ;
		substr(str(&fieldpassed.,8),6,3) + ")")
	otherwise
		.cell(tcount,cellpassed).range.insertafter("$  " + alltrim(str(&fieldpassed.)))
endcase
endproc

Procedure numberformatopt
Parameters fieldpassedo
do case
	* seven digit
	case &fieldpassedo. >  999999 .and. &fieldpassedo. < 10000000
		WordApp.selection.insertafter("$" + ;
		substr(str(&fieldpassedo.,8),1,2) + "," + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3))
	* six digit
	case &fieldpassedo. >   99999 .and. &fieldpassedo. < 1000000
		WordApp.selection.insertafter("$  " + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3))
	* five digit
	case &fieldpassedo. >    9999 .and. &fieldpassedo. < 100000
		WordApp.selection.insertafter("$   " + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3))
	* four digit
	case &fieldpassedo. >     999 .and. &fieldpassedo. < 10000
		WordApp.selection.insertafter("$    " + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3))
	* three digit
	case &fieldpassedo. >      99 .and. &fieldpassedo. < 1000
		WordApp.selection.insertafter("$         " + ;
		substr(str(&fieldpassedo.,8),6,3))
	* two digit
	case &fieldpassedo. >       9 .and. &fieldpassedo. < 100
		WordApp.selection.insertafter("$          " + ;
		substr(str(&fieldpassedo.,8),6,3))
	* one digit
	case &fieldpassedo. =>      0 .and. &fieldpassedo. < 10
		WordApp.selection.insertafter("$           " + ;
		substr(str(&fieldpassedo.,8),6,3))
	* minus one digit
	case &fieldpassedo. <       0 .and. &fieldpassedo. > -10
		WordApp.selection.insertafter("$           -(" + ;
		substr(str(&fieldpassedo.,8),8,1) + ")")
	* minus two digit
	case &fieldpassedo. <      -9 .and. &fieldpassedo. > -100
		WordApp.selection.insertafter("$         -(" + ;
		substr(str(&fieldpassedo.,8),7,2) + ")")
	* minus three digit
	case &fieldpassedo. <     -99 .and. &fieldpassedo. > -1000
		WordApp.selection.insertafter("$       -(" + ;
		substr(str(&fieldpassedo.,8),6,3) + ")")
	* minus four digit
	case &fieldpassedo. <    -999 .and. &fieldpassedo. > -10000
		WordApp.selection.insertafter("$    -(" + ;
		substr(str(&fieldpassedo.,8),5,1) + "," + ;
		substr(str(&fieldpassedo.,8),6,3) + ")")
	* minus five digit
	case &fieldpassedo. <   -9999 .and. &fieldpassedo. > -100000
		WordApp.selection.insertafter("$   -(" + ;
		substr(str(&fieldpassedo.,8),4,2) + "," + ;
		substr(str(&fieldpassedo.,8),6,3) + ")")
	* minus five digit
	case &fieldpassedo. <  -99999 .and. &fieldpassedo. > -1000000
		WordApp.selection.insertafter("$  -(" + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3) + ")")
	* minus six digit
	case &fieldpassedo. < -999999 .and. &fieldpassedo. > -10000000
		WordApp.selection.insertafter("$ -(" + ;
		substr(str(&fieldpassedo.,8),2,1) + "," + ;
		substr(str(&fieldpassedo.,8),3,3) + "," + ;
		substr(str(&fieldpassedo.,8),6,3) + ")")
	otherwise
		WordApp.selection.insertafter("$  " + alltrim(str(&fieldpassed.)))
endcase
endproc