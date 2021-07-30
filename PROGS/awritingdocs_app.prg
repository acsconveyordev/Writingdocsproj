* AWRITINGDOCSPROJ_APP.PRG
* This file is a generated, framework-enabling component
* created by APPBUILDER
* (c) Microsoft Corporation

* If machine information has not been assigned or the network shell hasn't
* been loaded, SYS(0) returns a character string consisting of 15 spaces,
* a number sign (#) followed by another space, and then 0.
* Consult your network documentation for further information on defining machine information.
*****************************************

* revised 05/23/2011
* ; c:\program files\microsoft visual studio\vfp98\ffc
* C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\databasefiles
* setup program
do writingdocssetup.prg

* added by t. hurt to test use of variables for the paths
* character variable for the data path
*!*	PUBLIC cdatapath
*!*	cdatapath = "C:\Quotedata\Molds\"

* character variable for the forms path
PUBLIC cformspath
*cformspath = "C:\Program Files\Microsoft Visual Studio\writingdocsproj\forms\"
cformspath = SYS(2003) + "\forms\"

* character variable for the program path
PUBLIC cprogpath
*cprogpath = "C:\Program Files\Microsoft Visual Studio\writingdocsproj\progs\"
cprogpath = SYS(2003) + "\progs\"

* character variable for the reports path
PUBLIC creportspath
*creportspath = "C:\Program Files\Microsoft Visual Studio\writingdocsproj\reports\"
creportspath = SYS(2003) +  "\reports\"

* character variable for the include path
* location for the templates
PUBLIC cincludepath
cincludepath = cInstallDir + "\include\"

* character variable for the database tables template directory
PUBLIC cmoldpath
*cmoldpath = "C:\Program Files\Microsoft Visual Studio\writingdocsproj\molds\"
cmoldpath = cInstallDir + "\Molds\"

* character variable for the quote tables directory
PUBLIC ctabledir
ctabledir = "C:\Quotedata\Quotes\"

* character variable for the quarantee and terms documents directory
PUBLIC cgtdocdir
cgtdocdir = cInstallDir + "\guarterms\"

PUBLIC ctemppath
ctemppath = SYS(2003) + "\temp\"
* make directory if it does not exist 
if directory(sys(2003) + "\TEMP") = .f.
	mkdir sys(2003) + "\Temp"
endif

* to be used to determine if the user is on the network
set exact on
PUBLIC lonnetwork, cuser
store .f. to lonnetwork
store "mm" to cuser

* use a table to identify the servers
Public cgserver,cserver
cgserver = rtrim("\\acs-fs1")
cserver = rtrim("\\acs-fs1")

cgserverdir = cgserver + "\general\documents\database"
*!*	cservertabledir = SYS(2003) + "\progs"
*!*	use cservertabledir + "\acsserverinfo" in a
*!*	go 3 in a
*!*	cgserver = rtrim("\\" + a.server)
*!*	*!*	? CGSERVER
*!*	cgserverdir = cgserver + "\general\documents\database"

* Check to make sure they are on the network
if directory(cgserverdir) = .t.
	store .t. to lonnetwork
	store lower(alltrim(substr(id(),at('#',id())+2,2))) to cuser
*!*		use cgserverdir + "\acsserverinfo.dbf" in a
*!*		go top
*!*		cserver = rtrim("\\" + a.server)
*!*		go 3 in a
*!*		cgserver = rtrim("\\" + a.server)
*!*		use in a
else
	* if not on network existing paths variables must be used
	store .f. to lonnetwork
*!*		? "ELSE USED"
endif

set exact off

* files on network
if lonnetwork = .t.
	* character variable for the data path
	PUBLIC cdatapath
		cdatapath = cgserver + "\GENERAL\DOCUMENTS\DATABASE\"

	* character variable for the database tables template directory
	PUBLIC cmoldpath
	cmoldpath = cgserver + "\SALESSRV\PROGRAMS\FOXPRO SUPPORT FILES\TEMPLATES\"
	* character variable for the quote tables directory
	PUBLIC ctabledir
	ctabledir = cgserver + "\SALESSRV\WORK IN PROGRESS\DRAWINGS\QUOTE\"
	* character variable for the quarantee and terms documents directory
	PUBLIC cgtdocdir
	cgtdocdir = cgserver + "\ADMINISTRATIVE\WORK IN PROGRESS\DOCUMENTS\WORD\G&T\"
	* character variable for the temporary path
	PUBLIC coverviewdir
	coverviewdir = cgserver + "\SALESSRV\PROGRAMS\FOXPRO SUPPORT FILES\OVERVIEWS\"
	PUBLIC ctemppath
	ctemppath = SYS(2003) + "\temp\"
	* make directory if it does not exist
	if directory(sys(2003) + "\TEMP") = .f.
		mkdir sys(2003) + "\Temp"
	endif
else
	* not on network
	cMessageTitle = 'WRITING DOCUMENTS PROGRAM'
	cMessageText = '   The NETWORK was not found do you want to quit?' + (chr(13)) ;
	             + '         Please YES to quit or NO to continue.'
	nDialogType = 4 + 16 + 256
	*  4 = YES and NO buttons
	*  16 = Stop sign
	*  256 = Second button is default
	nanswer = MESSAGEBOX(cMessageText, nDialogType, cMessageTitle)
	do case
	case nanswer = 6
		quit
	case nanswer = 7
		* continue
	endcase
endif

* set default and path
* SET DEFAULT TO "c:\program files\microsoft visual studio\writingdocsproj"
set default to sys(2003)
set path to progs; forms; libs; reports; menus; temp; graphics; &cgserverdir; &cwtserverdir, c:\program files\microsoft visual studio\vfp98\ffc, \\acs-fs1\Programs\Prod\Foxpro\Writingdocsproj


* temporary printout of path variables
* character variable for the program path
*!*	if cuser = "tn"
*!*		? "The data path is " + cdatapath
*!*		? "The forms path is " + cformspath
*!*		? "The program path is " + cprogpath
*!*		? "The reports path is " + creportspath
*!*		? "The include path is " + cincludepath
*!*		? "The mold path is " + cmoldpath
*!*		? "The table directory is " + ctabledir
*!*		? "The guarantee and terms document directory is " + cgtdocdir
*!*		? "The temporary path is " + ctemppath
*!*		? "The overview directory is " + coverviewdir
*!*		? "The set default is " + set("default")
*!*		? "The set path is " + set("path")
*!*		? "The cInstallDir is " + cInstallDir
*!*	endif

* defined in SELECT Drawing NUMBER program
PUBLIC cdrawing    && Drawing number
PUBLIC clayout     && Same as Dwg. No. plus dashes 9999-C-C
PUBLIC ctitle      && Quotation title
PUBLIC cquote      && Quote number
PUBLIC ccustomer   && Customer name
PUBLIC clocation   && Customers' location
store "" to ctitle, cquote, cdrawing, ccustomer, clocation

PUBLIC nreccount, nrecnum
store 0 to nreccount, nrecnum

PUBLIC cgetdrawingnum
store "" to cgetdrawingnum

* clear screen before starting
* clear

* Framework-generated application startup program
* for C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\WRITINGDOCSPROJ\WRITINGDOCSPROJ Project

#INCLUDE [..\WRITINGDOCSPROJ_APP.H]

IF TYPE([APP_GLOBAL.Class]) = "C" AND ;
   UPPER(APP_GLOBAL.Class) == UPPER(APP_CLASSNAME)
   MESSAGEBOX(APP_ALREADY_RUNNING_LOC,48, APP_GLOBAL.cCaption )
   IF VARTYPE(APP_GLOBAL.oFrame) = "O"
      APP_GLOBAL.oFrame.Show()
   ENDIF
   RETURN
ENDIF
RELEASE APP_GLOBAL
PUBLIC  APP_GLOBAL
LOCAL lcLastSetTalk, llAppRan, lnSeconds, loSplash
LOCAL ARRAY laCheck[1]
lcLastSetTalk=SET("TALK")
loSplash = .NULL.
SET TALK OFF
#IFDEF APP_SPLASHCLASS
   IF NOT EMPTY(APP_SPLASHCLASS)
      loSplash = NEWOBJECT(APP_SPLASHCLASS, APP_SPLASHCLASSLIB)
      IF VARTYPE(loSplash) = "O"
         lnSeconds = SECONDS()
         loSplash.Show()
      ENDIF
   ENDIF
#ENDIF
APP_GLOBAL = NEWOBJECT(APP_CLASSNAME, APP_CLASSLIB)
IF VARTYPE(APP_GLOBAL) = "O" ;
      AND ACLASS(laCheck,APP_GLOBAL) > 0 AND ;
      ASCAN(laCheck,UPPER(APP_SUPERCLASS)) > 0
   APP_GLOBAL.cReference =[APP_GLOBAL]
   APP_GLOBAL.cFormMediatorName = APP_MEDIATOR_NAME
   #IFDEF APP_CD
      APP_CD
   #ENDIF
   #IFDEF APP_PATH
      APP_PATH
   #ENDIF
   #IFDEF APP_INITIALIZE
       APP_INITIALIZE
   #ENDIF
   IF VARTYPE(loSplash) = "O"
      IF SECONDS() < lnSeconds + APP_SPLASHDELAY
         =INKEY(APP_SPLASHDELAY-(SECONDS()-lnSeconds),"MH")
      ENDIF
      loSplash.Release()
      loSplash = .NULL.
   ENDIF
   RELEASE laCheck, loSplash, lnSeconds
   IF NOT APP_GLOBAL.Show()
      IF TYPE([APP_GLOBAL.Name]) = "C"
         MESSAGEBOX(APP_CANNOT_RUN_LOC,16, APP_GLOBAL.cCaption)
         APP_GLOBAL.Release()
      ELSE
         MESSAGEBOX(APP_CANNOT_RUN_LOC,16)
      ENDIF
   ELSE
      llAppRan = .T.
   ENDIF
   IF TYPE([APP_GLOBAL.lReadEvents]) = "L"
      IF APP_GLOBAL.lReadEvents
         * the Release() method was not used
         * but we've somehow gotten out of READ EVENTS...
         APP_GLOBAL.Release()
      ENDIF
   ELSE
      RELEASE APP_GLOBAL
   ENDIF
ELSE
   MESSAGEBOX(APP_WRONG_SUPERCLASS_LOC,16)
   RELEASE APP_GLOBAL
ENDIF
IF lcLastSetTalk=="ON"
   SET TALK ON
ELSE
   SET TALK OFF
ENDIF
IF TYPE([APP_GLOBAL]) = "O"
   * non-read events app
   RETURN APP_GLOBAL
ELSE
   RETURN llAppRan
ENDIF