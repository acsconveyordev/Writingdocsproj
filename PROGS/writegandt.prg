* Name......... WRITE the Guarantee AND Terms into the word program
* Date......... 05/29/2012
* Caller....... writeppdoc.prg
* Notes........ This inserts the guarantee and terms if required.
*      12/08/04 Fork truck note was added to the template thus + 1 to Paragraphs()
*      09/28/07 Storage cost added to the template thus +1 to Paragraphs()
*      01/12/09 Paragraph count reduced by 2 when SOO page removed
*      02/19/09 Paragraph count reduced by 2 for Smurfit when page changed
*      04/22/10 Paragraph count increased by 10 when training added

* This header contains all the Word constants.
#INCLUDE [..\wordvba.h]

PUBLIC ngtpcount
store 0 to ngtpcount

* inserts the guarantee and terms if required
* different for BEACON CONTAINER, BIRDSBORO (481000)
* different terms for LIN PAC (97)
* different terms for WEYERHAEUSER (198)
* different terms for SMURFIT (546)
* different terms for Georgia-Pacific (630)
* different terms for INLAND (682)
* different terms for LONGVIEW FIBRE (755)
* different terms for PCA (841)
* different terms for SOUTHERN (899)
* different terms for IP (1098) Ireland (1098068)
* different terms for Avanti
* SOUTHERN CHANGED TO ROCK TENN
* Rock Tenn cjanged from 899 to 879

* remove the existing information
if lcoocfgt = .t. .or. ncompleteacsno = 481000 .or. ;
   ncompnumber =  97 .or. ncompnumber = 198 .or. ncompnumber = 546 .or. ncompnumber = 630 .or. ;
   ncompnumber = 682 .or. ncompnumber = 755 .or. ncompnumber = 841 .or. ncompnumber = 879 .or. ;
   (ncompnumber = 1098 .and. ncompleteacsno # 1098068)
	* Select the page and delete then replace with new page
*!*		WordApp.ActiveDocument.Range(WordApp.ActiveDocument.Paragraphs(297).Range.Start,WordApp.ActiveDocument.Paragraphs(308).Range.End).Select
	WordApp.ActiveDocument.Range(WordApp.ActiveDocument.Paragraphs(288).Range.Start,WordApp.ActiveDocument.Paragraphs(299).Range.End).Delete

*!*		WordApp.ActiveDocument.Paragraphs(296).Range.Select
	WordApp.ActiveDocument.Paragraphs(287).Range.Select

	WordApp.Selection.MoveLeft(wdCharacter,1)
	WordApp.Selection.MoveRight(wdCharacter,1)
	* this places the cursor at the beginning of the word LIMITED
	WordApp.Selection.Delete(wdCharacter,38)

	do case
	case lcoocfgt = .t.
		WordApp.Selection.InsertFile((cgtdocdir + "NONUSAG&Tpp.DOC"),"",.f.,.f.,.f.)
		store 1 to ngtpcount
	case ncompleteacsno = 481000    && different for BEACON CONTAINER, BIRDSBORO
		WordApp.Selection.InsertFile((cgtdocdir + "BEACONG&Tpp.DOC"),"",.f.,.f.,.f.)
		store 0 to ngtpcount
	case ncompnumber = 97  && different terms for U.S. Corrugated
		WordApp.Selection.InsertFile((cgtdocdir + "USCORRG&Tpp.DOC"),"",.f.,.f.,.f.)
		store -2 to ngtpcount
	* Weyerhaeuser bought by IPCO thus use the IPCO G&T
	case ncompnumber = 198  && different terms for WEYERHAEUSER
		* WordApp.Selection.InsertFile((cgtdocdir + "WEYCOG&Tpp.DOC"),"",.f.,.f.,.f.)
		WordApp.Selection.InsertFile((cgtdocdir + "IPCOG&Tpp.DOC"),"",.f.,.f.,.f.)
		* store -2 to ngtpcount
		store 4 to ngtpcount
	case ncompnumber = 546  && different terms for SMURFIT
		WordApp.Selection.InsertFile((cgtdocdir + "SMURFITG&Tpp.DOC"),"",.f.,.f.,.f.)
*!*			store 3 to ngtpcount
		store 1 to ngtpcount
	case ncompnumber = 630  && different terms for Georgia-Pacific
		WordApp.Selection.InsertFile((cgtdocdir + "GPG&Tpp.DOC"),"",.f.,.f.,.f.)
		store 1 to ngtpcount
	case ncompnumber = 682    && different terms for Temple-Inland
		WordApp.Selection.InsertFile((cgtdocdir + "TEMINLG&Tpp.DOC"),"",.f.,.f.,.f.)
		store 3 to ngtpcount
	case ncompnumber = 755    && different terms for LONGVIEW FIBRE
		WordApp.Selection.InsertFile((cgtdocdir + "LONGVG&Tpp.DOC"),"",.f.,.f.,.f.)
		store 0 to ngtpcount
	case ncompnumber = 841    && different terms for PCA
		WordApp.Selection.InsertFile((cgtdocdir + "PCAG&Tpp.DOC"),"",.f.,.f.,.f.)
		store 0 to ngtpcount
	case ncompnumber = 879    && different terms for ROCK TENN
		WordApp.Selection.InsertFile((cgtdocdir + "ROCKTENNG&Tpp.DOC"),"",.f.,.f.,.f.)
		store 2 to ngtpcount
	case ncompnumber = 1098 .and. ncompleteacsno # 1098068    && different terms for IP
		WordApp.Selection.InsertFile((cgtdocdir + "IPCOG&Tpp.DOC"),"",.f.,.f.,.f.)
		store 4 to ngtpcount
	endcase
	WordApp.Selection.TypeBackspace
	WordApp.Selection.TypeBackspace
else
	* leave as is
endif

* different terms for Avanti
* this is not a insertion but a correction of the standard terms
if lforavanti = .t.
*!*		WordApp.ActiveDocument.Paragraphs(301).Range.Select
	WordApp.ActiveDocument.Paragraphs(296).Range.Select
	WordApp.Selection.Collapse
	WordApp.Selection.Delete(wdCharacter,2)
	WordApp.Selection.InsertAfter("15")
*!*		WordApp.ActiveDocument.Paragraphs(302).Range.Select
	WordApp.ActiveDocument.Paragraphs(297).Range.Select
	WordApp.Selection.Collapse
	WordApp.Selection.Delete(wdCharacter,2)
	WordApp.Selection.InsertAfter("75")
*!*		WordApp.ActiveDocument.Paragraphs(303).Range.Select
	WordApp.ActiveDocument.Paragraphs(298).Range.Select
	WordApp.Selection.Collapse
	WordApp.Selection.Delete(wdCharacter,6)
	WordApp.Selection.InsertAfter("10%")
endif

*-- EOP WRITEGANDT