* Name......... WRITING the word DOCumentS SETUP program
* Date......... 11/12/2003
* Called from.. writingdocs_app.prg

* environment for this application/executable
* to get this information
* first select the Tools menu
* then select the Options selection
* then from the dialogue press Shift and select the ok button
* this write the information to the command window
* from there copy and paste into this program
* shown here from four stars to four stars
****
SET TALK OFF
SET NOTIFY ON
SET CLOCK STATUS
SET COMPATIBLE OFF
SET PALETTE ON
SET BELL ON
SET BELL TO '', 1
SET SAFETY ON
SET ESCAPE ON
SET LOGERRORS ON
SET KEYCOMP TO WINDOWS
SET CARRY OFF
SET CONFIRM OFF
SET BROWSEIMECONTROL OFF
SET STRICTDATE TO 1
&& TabOrdering = 0 && ResWidth = 1024 && ResHeight = 768 && GridHorz = 6 && GridVert = 6
&& ScaleUnits = 0 && FormSetLib =  && FormSetClass =  && FormsLib =  && FormsClass = 
SET EXACT OFF
SET NEAR OFF
SET ANSI OFF
SET LOCK OFF
SET EXCLUSIVE OFF
SET MULTILOCKS ON
SET HEADINGS ON
SET DELETED OFF
SET OPTIMIZE ON
SET UNIQUE OFF
SET CPDIALOG ON
SET REFRESH TO 0,5
SET ODOMETER TO 100
SET BLOCKSIZE TO 64
SET REPROCESS TO 0
SET COLLATE TO ""
&& SCCProvider = 
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\SCCTEXT.PRG" TO _SCCTEXT
&& ProjectHookLib =  && ProjectHookClass =  && CrsBuffering = 1 && CrsMethodUsed = 1 && CrsWhereClause = 3
&& CrsFetchSize = 100 && CrsMaxRows = -1 && CrsNumBatch = 1 && CrsUseMemoSize = 255 && SQLDispLogin = 1
&& SQLTransactions = 1 && SQLConnectTimeOut = 15 && SQLIdleTimeOut = 0 && SQLQueryTimeOut = 0 && SQLWaitTime = 100
&& TMPFILES = c:\windows\temp 
&& HelpTo = c:\program files\microsoft visual studio\msdn98\98vs\1033\msdnvs98.col
* SET HELP ON
* SET HELP TO "c:\program files\microsoft visual studio\msdn98\98vs\1033\msdnvs98.col"
&& ResourceTo = c:\program files\microsoft visual studio\vfp98\foxuser.dbf
*!*	SET RESOURCE ON
*!*	SET RESOURCE TO "c:\program files\microsoft visual studio\vfp98\foxuser.dbf"
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\WIZARD.APP" TO _WIZARD
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\BUILDER.APP" TO _BUILDER
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\CONVERT.APP" TO _CONVERTER
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\SPELLCHK.APP" TO _SPELLCHK
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\GENMENU.FXP" TO _GENMENU
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\BROWSER.APP" TO _BROWSER
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\GALLERY.APP" TO _GALLERY
*!*	STORE "" TO _INCLUDE
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\BEAUTIFY.APP" TO _BEAUTIFY
*!*	STORE "" TO _GETEXPR
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\GENHTML.FXP" TO _GENHTML
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\RUNACTD.PRG" TO _RUNACTIVEDOC
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\VFP6STRT.APP" TO _STARTUP
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\MSDN98\98VS\1033\SAMPLES\VFP98\" TO _SAMPLES
*!*	STORE "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VFP98\COVERAGE.APP" TO _COVERAGE
SET SYSFORMATS OFF
SET SECONDS ON
SET CENTURY ON
&& CurrSymbol = $
SET CURRENCY LEFT
SET CURRENCY TO "$"
SET HOURS TO 12
SET DATE TO AMERICAN
SET DECIMALS TO 2
SET FDOW TO 1
SET FWEEK TO 1
SET MARK TO ""
SET SEPARATOR TO ","
SET POINT TO "."
&& DebugEnvironment = 0
SET TRBETWEEN OFF
*!*	STORE     0.00 TO _THROTTLE
&& DebugOutputFileName =  && TraceFontName = Courier New && TraceFontSize = 10 && TraceFontStyle = 0
&& WatchFontName = MS Sans Serif && WatchFontSize = 8 && WatchFontStyle = 0 && LocalsFontName = MS Sans Serif
&& LocalsFontSize = 8 && LocalsFontStyle = 0 && OutputFontName = MS Sans Serif && OutputFontSize = 8
&& OutputFontStyle = 0 && CallstackFontName = MS Sans Serif && CallstackFontSize = 8 && CallstackFontStyle = 0
&& TraceNormalColor = RGB(0,0,0,255,255,255), Auto, Auto && TraceExecutingColor = RGB(255,255,0,0,0,0), NoAuto, Auto
&& TraceCallstackColor = RGB(0,0,0,192,192,192), Auto, Auto && TraceBreakpointColor = RGB(255,0,0,0,0,0), NoAuto, Auto
&& TraceSelectedColor = RGB(255,255,255,0,0,0), Auto, Auto && WatchNormalColor = RGB(0,0,0,255,255,255), Auto, Auto
&& WatchSelectedColor = RGB(255,255,255,0,0,0), Auto, Auto && WatchChangedColor = RGB(255,0,0,255,255,255), NoAuto, Auto
&& LocalsNormalColor = RGB(0,0,0,255,255,255), Auto, Auto && LocalsSelectedColor = RGB(255,255,255,0,0,0), Auto, Auto
&& OutputNormalColor = RGB(0,0,0,255,255,255), Auto, Auto && OutputSelectedColor = RGB(255,255,255,0,0,0), Auto, Auto
&& CallstackNormalColor = RGB(0,0,0,255,255,255), Auto, Auto && CallstackSelectedColor = RGB(255,255,255,0,0,0), Auto, Auto

* needed for use with memlines() function
SET MEMOWIDTH TO 65

****
PUBLIC cInstallDir
cInstallDir = SYS(2003)

* set screen to maximize
_screen.windowstate = 2

*-- EOP WRITINGDOCSSETUP