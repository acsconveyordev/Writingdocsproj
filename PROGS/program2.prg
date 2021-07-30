*!*	USE "G:\Documents\Database\TIMEMOLD" IN A EXCL
*!*	GO TOP
*!*	DO WHILE EOF() = .F.
*!*		DELETE
*!*		SKIP 1 IN A
*!*	ENDDO

USE "G:\Documents\Database\TIME" IN A EXCL
SET ORDER TO WEEKENDING
N = 100
GO TOP
DO WHILE N > 0
	DELETE
	N = N-1
	SKIP 1 IN A
ENDDO

PACK
REINDEX 
BROWSE

*!*	date()-489