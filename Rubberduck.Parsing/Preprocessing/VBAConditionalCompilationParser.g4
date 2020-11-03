parser grammar VBAConditionalCompilationParser;

options { tokenVocab = VBALexer; }

compilationUnit : ccBlock EOF;

ccBlock :
    // We use the non-greedy operator so that we stop consuming EOFs as soon as possible.
    // EOFs can be used in other places than the start rule since they're "emitted as long as someone takes them".
    (ccConst | ccIfBlock | physicalLine)*?
;

ccConst : whiteSpace* hashConst whiteSpace+ ccVarLhs whiteSpace* EQ whiteSpace* ccExpression ccEol;
ccVarLhs : name;

ccIfBlock : ccIf ccBlock ccElseIfBlock* ccElseBlock? ccEndIf;
ccIf : whiteSpace* hashIf whiteSpace+ ccExpression whiteSpace+ THEN ccEol;
ccElseIfBlock : ccElseIf ccBlock;
ccElseIf : whiteSpace* hashElseIf whiteSpace+ ccExpression whiteSpace+ THEN ccEol;
ccElseBlock : ccElse ccBlock;
ccElse : whiteSpace* hashElse ccEol;
ccEndIf : whiteSpace* hashEndIf ccEol;
ccEol : whiteSpace* comment? (NEWLINE | EOF);
// We use parser rules instead of tokens (such as HASHCONST) because
// marked file numbers have a similar format and cause conflicts.
hashConst : HASH CONST;
hashIf : HASH IF;
hashElseIf : HASH ELSEIF;
hashElse : HASH ELSE;
hashEndIf : HASH END_IF;

physicalLine : ~(NEWLINE | EOF)* (NEWLINE | EOF);

ccExpression :
    LPAREN whiteSpace* ccExpression whiteSpace* RPAREN
    | ccExpression whiteSpace* POW whiteSpace* ccExpression
    | MINUS whiteSpace* ccExpression
    | ccExpression whiteSpace* (MULT | DIV) whiteSpace* ccExpression
    | ccExpression whiteSpace* INTDIV whiteSpace* ccExpression
    | ccExpression whiteSpace* MOD whiteSpace* ccExpression
    | ccExpression whiteSpace* (PLUS | MINUS) whiteSpace* ccExpression
    | ccExpression whiteSpace* AMPERSAND whiteSpace* ccExpression
    | ccExpression whiteSpace* (EQ | NEQ | LT | GT | LEQ | GEQ | LIKE | IS) whiteSpace* ccExpression
    | NOT whiteSpace* ccExpression
    | ccExpression whiteSpace* AND whiteSpace* ccExpression
    | ccExpression whiteSpace* OR whiteSpace* ccExpression
    | ccExpression whiteSpace* XOR whiteSpace* ccExpression
    | ccExpression whiteSpace* EQV whiteSpace* ccExpression
    | ccExpression whiteSpace* IMP whiteSpace* ccExpression
    | intrinsicFunction
    | literal
    | name;

intrinsicFunction : intrinsicFunctionName LPAREN whiteSpace* ccExpression whiteSpace* RPAREN;

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
     | CSNG
     | CSTR
     | CURRENCY
     | CVAR
     | CVERR
     | DATABASE
     | DATE
     | DEBUG
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
     | PTRSAFE
     | REM
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
	 | NAME
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

whiteSpace : (WS | LINE_CONTINUATION)+;