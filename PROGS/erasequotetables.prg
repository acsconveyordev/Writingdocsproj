* Name......... ERASE the selected drawing number QUOTE TABLES from the network program
* Date......... 01/09/2007
* Caller....... writingdocs_app.prg
* Notes........

if cuser = "th"
	? "lgotdwgnfq is " 
	? lgotdwgnfq
	? "lgotdwgnfbd is "
	? lgotdwgnfbd
	? "lgotdwgnfbdd is "
	? lgotdwgnfbdd
	? "lgotdwgnfbdd0 is "
	? lgotdwgnfbdd0
	? "lgotdwgnfcqq is "
	? lgotdwgnfcqq
endif

close tables all

* variables for form and/or program
* four variables for the four possible directorys
*!*	PUBLIC lgotdwgnfq      && cserver + \salessrv\work in progress\drawings\quote
*!*	PUBLIC lgotdwgnfbd     && cserver + \salessrv\backup\drawings\
*!*	PUBLIC lgotdwgnfbdd    && cserver + \salessrv\backup\drawings\" + alltrim(str(int(val(cgetdrawingnum)/100))) + "00\"
*!*	PUBLIC lgotdwgnfbdd0   && cserver + \salessrv\backup\drawings\" + "0" + alltrim(str(int(val(cgetdrawingnum)/100))) + "00\"
if lgotdwgnfq = .t.
	* do nothing sales support has not zipped up the tables yet.
	set message to 'Tables not erased because sales support has not zipped them yet.'
	wait 'Tables not erased because sales support has not zipped them yet.' window at 5,10 timeout 5
endif
* tables have been zipped up and were unzipped by operator
if lgotdwgnfbd = .t. .or. lgotdwgnfbdd = .t. .or. lgotdwgnfbdd0 = .t.
	erase cserver + "\salessrv\work in progress\drawings\quote\" + alltrim(cgetdrawingnum) + "*.dbf"
	erase cserver + "\salessrv\work in progress\drawings\quote\" + alltrim(cgetdrawingnum) + "*.fpt"
	if file(cserver + "\salessrv\work in progress\drawings\quote\" + alltrim(cgetdrawingnum) + "v.dbf") = .t.
		set message to 'Tables erased.'
		wait 'Tables erased.' window at 5,10 timeout 5
	endif
endif

* set message
set message to 'Completed.'
wait 'Completed.' window at 6,10 timeout 5

* clean up
set message to
* if you just erased the tables you must reset cdrawing so 
*  that the operator cannot try to create the documents again
store "" to cdrawing
store "" to cgetdrawing

*-- EOP ERASEQUOTETABLES










