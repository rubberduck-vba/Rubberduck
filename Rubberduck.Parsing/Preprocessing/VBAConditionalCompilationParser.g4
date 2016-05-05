parser grammar VBAConditionalCompilationParser;

options { tokenVocab = VBALexer; }

compilationUnit : ccBlock EOF;

ccBlock : (ccConst | ccIfBlock | logicalLine)*;

ccConst : WS* HASHCONST WS+ ccVarLhs WS+ EQ WS+ ccExpression ccEol;

logicalLine : extendedLine+;

extendedLine : (LINE_CONTINUATION | ~(HASHCONST | HASHIF | HASHELSEIF | HASHELSE | HASHENDIF))+ NEWLINE?;

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

ccIf : HASHIF WS+ ccExpression WS+ THEN ccEol;

ccElseIfBlock : ccElseIf ccBlock;

ccElseIf : HASHELSEIF WS+ ccExpression WS+ THEN ccEol;

ccElseBlock : ccElse ccBlock;

ccElse : HASHELSE ccEol;

ccEndIf : HASHENDIF ccEol;

ccEol : (SINGLEQUOTE ~NEWLINE*)? NEWLINE?;

intrinsicFunction : intrinsicFunctionName LPAREN WS* ccExpression WS* RPAREN;

intrinsicFunctionName :
    INT |
    FIX |
    ABS |
    SGN |
    LEN |
    LENB |
    CBOOL |
    CBYTE |
    CCUR |
    CDATE |
    CDBL |
    CINT |
    CLNG |
    CLNGLNG |
    CLNGPTR |
    CSNG |
    CSTR |
    CVAR
;

name : IDENTIFIER typeHint?;

typeHint : PERCENT | AMPERSAND | POW | EXCLAMATIONPOINT | HASH | AT | DOLLAR;

literal : DATELITERAL | HEXLITERAL | OCTLITERAL | FLOATLITERAL | INTEGERLITERAL | STRINGLITERAL | TRUE | FALSE | NOTHING | NULL | EMPTY;