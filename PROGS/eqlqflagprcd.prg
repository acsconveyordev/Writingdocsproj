* Name......... EQuipment List QFLAG PRoCeDure program
* Date......... 03/01/2002
* Caller....... quotation_app.prg
* Caller....... Equiplist
* Notes........ Standard qflag subroutines without pricing

* qflag = 1 Use for parts which are priced per each. Identical parts within
*           the same item are counted and the count is multiplied by the
*           fixed price from the master price file.
PROCEDURE ELQFLAG1
select a
do while PART = cpart .and. ITEM = citem .and. alltrim(LAYER) = clayer
	store nqty + 1 to nqty
	skip 1 in a
enddo
store '' to clength
store '' to cscale
return
ENDPROC
*-- EOP ELQFLAG1

* qflag = 2 For parts, primarily conveyor, which are pricd by there length per inch
*           The XSCALE field from the drawing is used. In the pricing routine, the XSCALE
*           field is converted to a string and added to the end of the description field.
*           If identical pieces of conveyor are found in the same item they are also
*           quantified. The tenth character in the part number is checked for identifing
*           the roll centers to determine hoe the item will be rounded.
PROCEDURE ELQFLAG2
select a
store '' to clength
store 1 to nqty
store PART to cpart
store XSCALE to nscale,ntest
store val(substr(cpart,10,1)) to ncenter
store round(nscale/ncenter,0) * ncenter to nscale
store ltrim(str(int(nscale/12))) + "'-" + ltrim(str(mod(nscale,12))) + '"' to cscale
skip 1 in a
do while ITEM = citem .and. PART = cpart .and. XSCALE = ntest .and. alltrim(LAYER) = clayer
	store nqty + 1 to nqty
	skip 1 in a
enddo
return
ENDPROC
*-- EOP ELQFLAG2

* qflag = 3 For manual entry parts. If identical descriptions are found
*           in the same item the prices are combined.
PROCEDURE ELQFLAG3
select a
do while ITEM = citem .and. trim(LAYER) = clayer .and. ;
 alltrim(DESCR) = cdescr .and. len(alltrim(DESCR)) = len(alltrim(cdescr))
	store nqty + 1 to nqty
	store '' to clength
	store '' to cscale
	skip 1 in a
enddo
return
ENDPROC
*-- EOP ELQFLAG3

* qflag = 4 For zones.  These blocks are presented by AUTOCAD as an increment of the zone
*           length and rounded up to a whole zone. Zone uses the XSCALE from the drawing
*           file as its quantity and multiplies by material, installation,rework, and total
*           prices. If identical entries are found in the same item they are quantified.
PROCEDURE ELQFLAG4
select a
do while PART = cpart .and. ITEM = citem .and. alltrim(LAYER) = clayer
	store nqty + XSCALE to nqty
	skip 1 in a
enddo
store '' to clength
store '' to cscale
return
ENDPROC
*-- EOP ELQFLAG4 

* qflag = 5 For fork truck bumpers which have a length in feet added to the beginning of
*           their description fields. XSCALE is converted to feet and forced up to a whole
*           foot quantity. In BUMPER the actual XSCALE is used in the description field.
PROCEDURE ELQFLAG5
select a
store ' ' + ltrim(str(XSCALE)) + '"' to clength
store XSCALE to ntest
store '' to cscale
store 1 to nqty
skip 1 in a
do while PART = cpart .and. ITEM = citem .and. XSCALE = ntest .and. alltrim(LAYER) = clayer
	store nqty + 1 to nqty
	skip 1 in a
enddo
return
ENDPROC
*-- EOP ELQFLAG5 

* qflag = 6 For combining items with the same part number. When identical part numbers
*           are found the quantity is tallied and the prices are multiplied by the quantity.
PROCEDURE ELQFLAG6
select a
do while PART = cpart .and. ITEM = citem .and. alltrim(LAYER) = clayer
	store nqty + 1 to nqty
	store '' to clength
	store '' to cscale
	skip 1 in a
enddo
return
ENDPROC
*-- EOP ELQFLAG6

* qflag = 7 Identical drives within ter same item are counted first. Then the program searches
*           for a feet per minute symbol to include in the description. if the feet per minute
*           symbol is not found, "???" is used in its place.
PROCEDURE ELQFLAG7
ctlayer = substr(clayer,1,8)
ctlayers = substr(clayer,1,7)
select a
do while PART = cpart .and. ITEM = citem .and. substr(LAYER,1,8) = ctlayer .and. ;
 alltrim(DESCR) = cdescr .and. len(alltrim(DESCR)) = len(alltrim(cdescr))
	store nqty + 1 TO nqty
	skip 1 in a
enddo
store '' to clength
store '' to cscale
if eof() = .t.
	store '???' to cscale
	return
else
	store recno() to nhold
endif

locate next 20 for mod(qflag,100) = 8 .and. substr(layer,1,7) = ctlayers while item = citem
if found() = .t.
	store alltrim(descr) to cscale
else
	store '???' to cscale
endif
goto nhold
return
ENDPROC
*-- EOP ELQFLAG7

* qflag = 8 Feet per minute symbols carry this flag.
*           They are searched for from qflag 7 and are then skipped over.

* qflag = 9 Two stacking zones can be operated by a common control. While they will be shown
*           on the drawing twice, they will be priced only once. This procedure looks for a
*           second stacking zone, if found, the price of each is halved.
PROCEDURE ELQFLAG9
select a
do while PART = cpart .and. ITEM = citem .and. alltrim(LAYER) = clayer
	store nqty + 1 to nqty
	skip 1 in a
enddo
store '' to clength
store '' to cscale
return
*-- EOP ELQFLAG9 
*-- EOP EQLQFLAGPRCD
