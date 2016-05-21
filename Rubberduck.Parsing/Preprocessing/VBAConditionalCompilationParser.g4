parser grammar VBAConditionalCompilationParser;

options { tokenVocab = VBALexer; }

compilationUnit : ccBlock EOF;

ccBlock : (ccConst | ccIfBlock | logicalLine)*;

ccConst : HASHCONST ccVarLhs WS+ EQ WS+ ccExpression ccEol;

logicalLine :
    ~(HASHCONST | HASHIF | HASHELSEIF | HASHELSE | HASHENDIF)+
    | NEWLINE
;

ccVarLhs : name;

ccExpression :
    LPAREN WS* ccExpression WS* RPAREN
	| ccExpression WS* POW WS* ccExpression
	| MINUS WS* ccExpression
	| ccExpression WS* (MULT | DIV) WS* ccExpression
	| ccExpression WS* INTDIV WS* ccExpression
	| ccExpression WS* MOD WS* ccExpression
	| ccExpression WS* (PLUS | MINUS) WS* ccExpression
	| ccExpression WS* AMPERSAND WS* ccExpression
	| ccExpression WS* (EQ | NEQ | LT | GT | LEQ | GEQ | LIKE | IS) WS* ccExpression
	| NOT WS* ccExpression
	| ccExpression WS* AND WS* ccExpression
	| ccExpression WS* OR WS* ccExpression
	| ccExpression WS* XOR WS* ccExpression
	| ccExpression WS* EQV WS* ccExpression
	| ccExpression WS* IMP WS* ccExpression
    | intrinsicFunction
    | literal
    | name;

ccIfBlock : ccIf ccBlock ccElseIfBlock* ccElseBlock? ccEndIf;

ccIf : HASHIF ccExpression WS+ THEN WS* ccEol;

ccElseIfBlock : ccElseIf ccBlock;

ccElseIf : HASHELSEIF ccExpression WS+ THEN WS* ccEol;

ccElseBlock : ccElse ccBlock;

ccElse : HASHELSE;

ccEndIf : HASHENDIF;

ccEol : comment? (NEWLINE | EOF);

intrinsicFunction : intrinsicFunctionName LPAREN WS* ccExpression WS* RPAREN;

intrinsicFunctionName :
    INT
    | FIX
    | ABS
    | SGN
    | LEN
    | LENB
    | CBOOL
    | CBYTE
    | CCUR
    | CDATE
    | CDBL
    | CINT
    | CLNG
    | CLNGLNG
    | CLNGPTR
    | CSNG
    | CSTR
    | CVAR
;

name : nameValue typeHint?;
nameValue : IDENTIFIER | keyword | foreignName | statementKeyword | markerKeyword;
foreignName : L_SQUARE_BRACKET foreignIdentifier* R_SQUARE_BRACKET;
foreignIdentifier : ~L_SQUARE_BRACKET | foreignName;

typeHint : PERCENT | AMPERSAND | POW | EXCLAMATIONPOINT | HASH | AT | DOLLAR;

literal : DATELITERAL | HEXLITERAL | OCTLITERAL | FLOATLITERAL | INTEGERLITERAL | STRINGLITERAL | TRUE | FALSE | NOTHING | NULL | EMPTY;

comment: SINGLEQUOTE (LINE_CONTINUATION | ~NEWLINE)*;

keyword : 
       ABS
     | ADDRESSOF
     | ALIAS
     | AND
     | ANY
     | ARRAY
     | ATTRIBUTE
     | BEGIN
     | BOOLEAN
     | BYREF
     | BYTE
     | BYVAL
     | CBOOL
     | CBYTE
     | CCUR
     | CDATE
     | CDBL
     | CDEC
     | CINT
     | CLASS
     | CLNG
     | CLNGLNG
     | CLNGPTR
     | COLLECTION
     | CSNG
     | CSTR
     | CURRENCY
     | CVAR
     | CVERR
     | DATABASE
     | DATE
     | DEBUG
     | DELETESETTING
     | DOEVENTS
     | DOUBLE
     | END
     | EQV
     | FALSE
     | FIX
     | IMP
     | IN
     | INPUTB
     | INT
     | INTEGER
     | IS
     | LBOUND
     | LEN
     | LEN
     | LENB
     | LIB
     | LIKE
     | LOAD
     | LONG
     | LONGLONG
     | LONGPTR
     | ME
     | MID
     | MIDB
     | MIDBTYPESUFFIX
     | MIDTYPESUFFIX
     | MOD
     | NEW
     | NOT
     | NOTHING
     | NULL
     | OBJECT
     | OPTIONAL
     | OR
     | PARAMARRAY
     | PRESERVE
     | PSET
     | REM
     | RMDIR
     | SENDKEYS
     | SETATTR
     | SGN
     | SINGLE
     | SPC
     | STRING
     | TAB
     | TEXT
     | THEN
     | TO
     | TRUE
     | TYPEOF
     | UBOUND
     | UNTIL
     | VARIANT
     | VERSION
     | WITHEVENTS
     | XOR
     | STEP
     | ON_ERROR
     | RESUME_NEXT
     | ERROR
     | APPEND
     | BINARY
     | OUTPUT
     | RANDOM
     | ACCESS
     | READ
     | WRITE
     | READ_WRITE
     | SHARED
     | LOCK_READ
     | LOCK_WRITE
     | LOCK_READ_WRITE
     | LINE_INPUT    
     | RESET
     | WIDTH
     | PRINT
     | GET
     | PUT
     | CLOSE
     | INPUT
     | LOCK
     | OPEN
     | SEEK
     | UNLOCK
     | WRITE
;

markerKeyword : AS;

statementKeyword :
    CALL
    | CASE
    | CONST
    | DECLARE
    | DEFBOOL
    | DEFBYTE
    | DEFCUR
    | DEFDATE
    | DEFDBL
    | DEFINT
    | DEFLNG
    | DEFLNGLNG
    | DEFLNGPTR
    | DEFOBJ
    | DEFSNG
    | DEFSTR
    | DEFVAR
    | DIM
    | DO
    | ELSE
    | ELSEIF
    | ENUM
    | ERASE
    | EVENT
    | EXIT
    | EXIT_DO 
    | EXIT_FOR 
    | EXIT_FUNCTION 
    | EXIT_PROPERTY 
    | EXIT_SUB
    | END_SELECT
    | END_WITH
    | FOR
    | FRIEND
    | FUNCTION
    | GLOBAL
    | GOSUB
    | GOTO
    | IF
    | IMPLEMENTS
    | LET
    | LOOP
    | LSET
    | NEXT
    | ON
    | OPTION
    | PRIVATE
    | PUBLIC
    | RAISEEVENT
    | REDIM
    | RESUME
    | RETURN
    | RSET
    | SELECT
    | SET
    | STATIC
    | STOP
    | SUB
    | TYPE
    | WEND
    | WHILE
    | WITH
;