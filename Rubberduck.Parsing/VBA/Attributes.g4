/*
* Copyright (C) 2014 Ulrich Wolffgang <u.wol@wwu.de>
*
* This program is free software: you can redistribute it and/or modify
* it under the terms of the GNU General Public License as published by
* the Free Software Foundation, either version 3 of the License, or
* (at your option) any later version.
* 
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
* GNU General Public License for more details.
* 
* You should have received a copy of the GNU General Public License
* along with this program. If not, see <http://www.gnu.org/licenses/>.
*/

/* VBA grammar based on Microsoft's [MS-VBAL]: VBA Language Specification. */

grammar Attributes;

// module ----------------------------------

startRule : module EOF;

module : 
	WS?
	endOfStatement
	(moduleHeader endOfStatement)?
	moduleConfig? endOfStatement
	moduleAttributes? endOfStatement
	moduleBody? endOfStatement
	WS?
;

moduleHeader : VERSION WS DOUBLELITERAL WS CLASS;

moduleConfig :
	BEGIN endOfStatement
	moduleConfigElement+
	END
;

moduleConfigElement :
	ambiguousIdentifier WS? EQ WS? literal endOfStatement
;

moduleAttributes : (attributeStmt endOfStatement)+;

moduleBody : 
	moduleBodyElement (endOfStatement moduleBodyElement)* endOfStatement;

moduleBodyElement : 
	functionStmt 
	| propertyGetStmt 
	| propertySetStmt 
	| propertyLetStmt 
	| subStmt 
;

attributeStmt : ATTRIBUTE WS implicitCallStmt_InStmt WS? EQ WS? literal (WS? ',' WS? literal)*;

block : blockStmt (endOfStatement blockStmt)* endOfStatement;

blockStmt : attributeStmt;

functionStmt :
	(visibility WS)? (STATIC WS)? FUNCTION WS? ambiguousIdentifier typeHint? (WS? argList)? (WS? asTypeClause)? endOfStatement
	block?
	END_FUNCTION
;

propertyGetStmt : 
	(visibility WS)? (STATIC WS)? PROPERTY_GET WS ambiguousIdentifier typeHint? (WS? argList)? (WS asTypeClause)? endOfStatement 
	block? 
	END_PROPERTY
;

propertySetStmt : 
	(visibility WS)? (STATIC WS)? PROPERTY_SET WS ambiguousIdentifier (WS? argList)? endOfStatement 
	block? 
	END_PROPERTY
;

propertyLetStmt : 
	(visibility WS)? (STATIC WS)? PROPERTY_LET WS ambiguousIdentifier (WS? argList)? endOfStatement 
	block? 
	END_PROPERTY
;

subStmt : 
	(visibility WS)? (STATIC WS)? SUB WS? ambiguousIdentifier (WS? argList)? endOfStatement
	block? 
	END_SUB
;

midStmt : MID WS? LPAREN WS? argsCall WS? RPAREN;

valueStmt : 
	literal                                                                         # vsLiteral
	| implicitCallStmt_InStmt                                                       # vsICS
	| LPAREN WS? valueStmt (WS? ',' WS? valueStmt)* RPAREN                          # vsStruct
	| NEW WS? valueStmt                                                             # vsNew
	| typeOfStmt                                                                    # vsTypeOf
	| midStmt                                                                       # vsMid
	| ADDRESSOF WS? valueStmt                                                       # vsAddressOf
	| implicitCallStmt_InStmt WS? ASSIGN WS? valueStmt                              # vsAssign
	| valueStmt WS? POW WS? valueStmt                                               # vsPow
	| MINUS WS? valueStmt                                                           # vsNegation
	| valueStmt WS? (MULT | DIV) WS? valueStmt                                      # vsMult
	| valueStmt WS? INTDIV WS? valueStmt                                            # vsIntDiv
	| valueStmt WS? MOD WS? valueStmt                                               # vsMod
	| valueStmt WS? (PLUS | MINUS) WS? valueStmt                                    # vsAdd
	| valueStmt WS? AMPERSAND WS? valueStmt                                         # vsAmp
	| valueStmt WS? (EQ | NEQ | LT | GT | LEQ | GEQ | LIKE | IS) WS? valueStmt      # vsRelational
	| NOT WS? valueStmt                                                             # vsNot
	| valueStmt WS? AND WS? valueStmt                                               # vsAnd
	| valueStmt WS? OR WS? valueStmt                                                # vsOr
	| valueStmt WS? XOR WS? valueStmt                                               # vsXor
	| valueStmt WS? EQV WS? valueStmt                                               # vsEqv
	| valueStmt WS? IMP WS? valueStmt                                               # vsImp
;

implicitCallStmt_InStmt :
	iCS_S_MembersCall
	| iCS_S_VariableOrProcedureCall
	| iCS_S_ProcedureOrArrayCall
	| iCS_S_DictionaryCall
;

iCS_S_VariableOrProcedureCall : ambiguousIdentifier typeHint? (WS? dictionaryCallStmt)? (WS? LPAREN subscripts RPAREN)*;

iCS_S_ProcedureOrArrayCall : (ambiguousIdentifier | baseType) typeHint? WS? LPAREN WS? (argsCall WS?)? RPAREN (WS? dictionaryCallStmt)? (WS? LPAREN subscripts RPAREN)*;

iCS_S_MembersCall : (iCS_S_VariableOrProcedureCall | iCS_S_ProcedureOrArrayCall)? (iCS_S_MemberCall WS?)+ (WS? dictionaryCallStmt)? (WS? LPAREN subscripts RPAREN)*;

iCS_S_MemberCall : ('.' | '!') WS? (iCS_S_VariableOrProcedureCall | iCS_S_ProcedureOrArrayCall);

iCS_S_DictionaryCall : WS? dictionaryCallStmt;


// atomic call statements ----------------------------------

argsCall : (argCall? WS? (',' | ';') WS?)* argCall (WS? (',' | ';') WS? argCall?)*;

argCall : LPAREN? ((BYVAL | BYREF | PARAMARRAY) WS)? RPAREN? valueStmt;

dictionaryCallStmt : '!' WS? ambiguousIdentifier typeHint?;

argList : LPAREN (WS? arg (WS? ',' WS? arg)*)? WS? RPAREN;

arg : (OPTIONAL WS)? ((BYVAL | BYREF) WS)? (PARAMARRAY WS)? ambiguousIdentifier typeHint? (WS? LPAREN WS? RPAREN)? (WS? asTypeClause)? (WS? argDefaultValue)?;

argDefaultValue : EQ WS? valueStmt;

subscripts : subscript (WS? ',' WS? subscript)*;

subscript : (valueStmt WS TO WS)? valueStmt;


// atomic rules ----------------------------------

ambiguousIdentifier : 
	(IDENTIFIER | ambiguousKeyword)+
;

typeOfStmt : TYPEOF WS valueStmt (WS IS WS type)?;

asTypeClause : AS WS? (NEW WS)? type (WS? fieldLength)?;

baseType : BOOLEAN | BYTE | COLLECTION | DATE | DOUBLE | INTEGER | LONG | SINGLE | STRING | VARIANT;

complexType : ambiguousIdentifier (('.' | '!') ambiguousIdentifier)*;

fieldLength : MULT WS? (INTEGERLITERAL | ambiguousIdentifier);

literal : HEXLITERAL | OCTLITERAL | DATELITERAL | DOUBLELITERAL | INTEGERLITERAL | SHORTLITERAL | STRINGLITERAL | TRUE | FALSE | NOTHING | NULL | EMPTY;

type : (baseType | complexType) (WS? LPAREN WS? RPAREN)?;

typeHint : '&' | '%' | '#' | '!' | '@' | '$';

visibility : PRIVATE | PUBLIC | FRIEND | GLOBAL;

// ambiguous keywords
ambiguousKeyword : 
	ACCESS | ADDRESSOF | ALIAS | AND | ATTRIBUTE | APPACTIVATE | APPEND | AS |
	BEEP | BEGIN | BINARY | BOOLEAN | BYVAL | BYREF | BYTE | 
	CALL | CASE | CLASS | CLOSE | CHDIR | CHDRIVE | COLLECTION | CONST | 
	DATABASE | DATE | DECLARE | DEFBOOL | DEFBYTE | DEFCUR | DEFDBL | DEFDATE | DEFDEC | DEFINT | DEFLNG | DEFOBJ | DEFSNG | DEFSTR | DEFVAR | DELETESETTING | DIM | DO | DOUBLE | 
	EACH | ELSE | ELSEIF | END | ENUM | EQV | ERASE | ERROR | EVENT | 
	FALSE | FILECOPY | FRIEND | FOR | FUNCTION | 
	GET | GLOBAL | GOSUB | GOTO | 
	IF | IMP | IMPLEMENTS | IN | INPUT | IS | INTEGER |
	KILL | 
	LOAD | LOCK | LONG | LOOP | LEN | LET | LIB | LIKE | LSET |
	ME | MID | MKDIR | MOD | 
	NAME | NEXT | NEW | NOT | NOTHING | NULL | 
	ON | OPEN | OPTIONAL | OR | OUTPUT | 
	PARAMARRAY | PRESERVE | PRINT | PRIVATE | PUBLIC | PUT |
	RANDOM | RANDOMIZE | RAISEEVENT | READ | REDIM | REM | RESET | RESUME | RETURN | RMDIR | RSET |
	SAVEPICTURE | SAVESETTING | SEEK | SELECT | SENDKEYS | SET | SETATTR | SHARED | SINGLE | SPC | STATIC | STEP | STOP | STRING | SUB | 
	TAB | TEXT | THEN | TIME | TO | TRUE | TYPE | TYPEOF | 
	UNLOAD | UNLOCK | UNTIL | 
	VARIANT | VERSION | 
	WEND | WHILE | WIDTH | WITH | WITHEVENTS | WRITE |
	XOR
;

remComment : REMCOMMENT;

comment : COMMENT;

endOfLine : WS? (NEWLINE | comment | remComment) WS?;

endOfStatement : (endOfLine | WS? COLON WS?)*;


// lexer rules --------------------------------------------------------------------------------


// keywords
ACCESS : A C C E S S;
ADDRESSOF : A D D R E S S O F;
ALIAS : A L I A S;
AND : A N D;
ATTRIBUTE : A T T R I B U T E;
APPACTIVATE : A P P A C T I V A T E;
APPEND : A P P E N D;
AS : A S;
BEGIN : B E G I N;
BEEP : B E E P;
BINARY : B I N A R Y;
BOOLEAN : B O O L E A N;
BYVAL : B Y V A L;
BYREF : B Y R E F;
BYTE : B Y T E;
CALL : C A L L;
CASE : C A S E;
CHDIR : C H D I R;
CHDRIVE : C H D R I V E;
CLASS : C L A S S;
CLOSE : C L O S E;
COLLECTION : C O L L E C T I O N;
CONST : C O N S T;
DATABASE : D A T A B A S E;
DATE : D A T E;
DECLARE : D E C L A R E;
DEFBOOL : D E F B O O L; 
DEFBYTE : D E F B Y T E;
DEFDATE : D E F D A T E;
DEFDBL : D E F D B L;
DEFDEC : D E F D E C;
DEFCUR : D E F C U R;
DEFINT : D E F I N T;
DEFLNG : D E F L N G;
DEFOBJ : D E F O B J;
DEFSNG : D E F S N G;
DEFSTR : D E F S T R;
DEFVAR : D E F V A R;
DELETESETTING : D E L E T E S E T T I N G;
DIM : D I M;
DO : D O;
DOUBLE : D O U B L E;
EACH : E A C H;
ELSE : E L S E;
ELSEIF : E L S E I F;
EMPTY : E M P T Y;
END_ENUM : E N D WS E N U M;
END_FUNCTION : E N D WS F U N C T I O N;
END_IF : E N D WS I F;
END_PROPERTY : E N D WS P R O P E R T Y;
END_SELECT : E N D WS S E L E C T;
END_SUB : E N D WS S U B;
END_TYPE : E N D WS T Y P E;
END_WITH : E N D WS W I T H;
END : E N D;
ENUM : E N U M;
EQV : E Q V;
ERASE : E R A S E;
ERROR : E R R O R;
EVENT : E V E N T;
EXIT_DO : E X I T WS D O;
EXIT_FOR : E X I T WS F O R;
EXIT_FUNCTION : E X I T WS F U N C T I O N;
EXIT_PROPERTY : E X I T WS P R O P E R T Y;
EXIT_SUB : E X I T WS S U B;
FALSE : F A L S E;
FILECOPY : F I L E C O P Y;
FRIEND : F R I E N D;
FOR : F O R;
FUNCTION : F U N C T I O N;
GET : G E T;
GLOBAL : G L O B A L;
GOSUB : G O S U B;
GOTO : G O T O;
IF : I F;
IMP : I M P;
IMPLEMENTS : I M P L E M E N T S;
IN : I N;
INPUT : I N P U T;
IS : I S;
INTEGER : I N T E G E R;
KILL: K I L L;
LOAD : L O A D;
LOCK : L O C K;
LONG : L O N G;
LOOP : L O O P;
LEN : L E N;
LET : L E T;
LIB : L I B;
LIKE : L I K E;
LINE_INPUT : L I N E WS I N P U T;
LOCK_READ : L O C K WS R E A D;
LOCK_WRITE : L O C K WS W R I T E;
LOCK_READ_WRITE : L O C K WS R E A D WS W R I T E;
LSET : L S E T;
ME : M E;
MID : M I D;
MKDIR : M K D I R;
MOD : M O D;
NAME : N A M E;
NEXT : N E X T;
NEW : N E W;
NOT : N O T;
NOTHING : N O T H I N G;
NULL : N U L L;
ON : O N;
ON_ERROR : O N WS E R R O R;
ON_LOCAL_ERROR : O N WS L O C A L WS E R R O R;
OPEN : O P E N;
OPTIONAL : O P T I O N A L;
OPTION_BASE : O P T I O N WS B A S E;
OPTION_EXPLICIT : O P T I O N WS E X P L I C I T;
OPTION_COMPARE : O P T I O N WS C O M P A R E;
OPTION_PRIVATE_MODULE : O P T I O N WS P R I V A T E WS M O D U L E;
OR : O R;
OUTPUT : O U T P U T;
PARAMARRAY : P A R A M A R R A Y;
PRESERVE : P R E S E R V E;
PRINT : P R I N T;
PRIVATE : P R I V A T E;
PROPERTY_GET : P R O P E R T Y WS G E T;
PROPERTY_LET : P R O P E R T Y WS L E T;
PROPERTY_SET : P R O P E R T Y WS S E T;
PTRSAFE : P T R S A F E;
PUBLIC : P U B L I C;
PUT : P U T;
RANDOM : R A N D O M;
RANDOMIZE : R A N D O M I Z E;
RAISEEVENT : R A I S E E V E N T;
READ : R E A D;
READ_WRITE : R E A D WS W R I T E;
REDIM : R E D I M;
REM : R E M;
RESET : R E S E T;
RESUME : R E S U M E;
RETURN : R E T U R N;
RMDIR : R M D I R;
RSET : R S E T;
SAVEPICTURE : S A V E P I C T U R E;
SAVESETTING : S A V E S E T T I N G;
SEEK : S E E K;
SELECT : S E L E C T;
SENDKEYS : S E N D K E Y S;
SET : S E T;
SETATTR : S E T A T T R;
SHARED : S H A R E D;
SINGLE : S I N G L E;
SPC : S P C;
STATIC : S T A T I C;
STEP : S T E P;
STOP : S T O P;
STRING : S T R I N G;
SUB : S U B;
TAB : T A B;
TEXT : T E X T;
THEN : T H E N;
TIME : T I M E;
TO : T O;
TRUE : T R U E;
TYPE : T Y P E;
TYPEOF : T Y P E O F;
UNLOAD : U N L O A D;
UNLOCK : U N L O C K;
UNTIL : U N T I L;
VARIANT : V A R I A N T;
VERSION : V E R S I O N;
WEND : W E N D;
WHILE : W H I L E;
WIDTH : W I D T H;
WITH : W I T H;
WITHEVENTS : W I T H E V E N T S;
WRITE : W R I T E;
XOR : X O R;


// symbols
AMPERSAND : '&';
ASSIGN : ':=';
DIV : '/';
INTDIV : '\\';
EQ : '=';
GEQ : '>=';
GT : '>';
LEQ : '<=';
LPAREN : '(';
LT : '<';
MINUS : '-';
MULT : '*';
NEQ : '<>';
PLUS : '+';
POW : '^';
RPAREN : ')';
L_SQUARE_BRACKET : '[';
R_SQUARE_BRACKET : ']';


// literals
STRINGLITERAL : '"' (~["\r\n] | '""')* '"';
OCTLITERAL : '&O' [0-8]+ '&'?;
HEXLITERAL : '&H' [0-9A-F]+ '&'?;
SHORTLITERAL : (PLUS|MINUS)? DIGIT+ ('#' | '&' | '@')?;
INTEGERLITERAL : SHORTLITERAL (E SHORTLITERAL)?;
DOUBLELITERAL : (PLUS|MINUS)? DIGIT* '.' DIGIT+ (E SHORTLITERAL)?;

DATELITERAL : '#' DATEORTIME '#';
fragment DATEORTIME : DATEVALUE WS? TIMEVALUE | DATEVALUE | TIMEVALUE;
fragment DATEVALUE : DATEVALUEPART DATESEPARATOR DATEVALUEPART (DATESEPARATOR DATEVALUEPART)?;
fragment DATEVALUEPART : DIGIT+ | MONTHNAME;
fragment DATESEPARATOR : WS? [/,-]? WS?;
fragment MONTHNAME : ENGLISHMONTHNAME | ENGLISHMONTHABBREVIATION;
fragment ENGLISHMONTHNAME : JANUARY | FEBRUARY | MARCH | APRIL | MAY | JUNE | JULY | AUGUST | SEPTEMBER | OCTOBER | NOVEMBER | DECEMBER;
fragment ENGLISHMONTHABBREVIATION : JAN | FEB | MAR | APR | JUN | JUL | AUG | SEP | OCT | NOV | DEC;
fragment TIMEVALUE : DIGIT+ AMPM | DIGIT+ TIMESEPARATOR DIGIT+ (TIMESEPARATOR DIGIT+)? AMPM?;
fragment TIMESEPARATOR : WS? (':' | '.') WS?;
fragment AMPM : WS? (A M | P M | A | P);

JANUARY : J A N U A R Y;
FEBRUARY : F E B R U A R Y;
MARCH : M A R C H;
APRIL : A P R I L;
MAY : M A Y;
JUNE : J U N E;
JULY : J U L Y;
AUGUST : A U G U S T;
SEPTEMBER : S E P T E M B E R;
OCTOBER : O C T O B E R;
NOVEMBER : N O V E M B E R;
DECEMBER : D E C E M B E R;
JAN : J A N;
FEB : F E B;
MAR: M A R;
APR : A P R;
JUN : J U N;
JUL: J U L;
AUG : A U G;
SEP : S E P;
OCT : O C T;
NOV : N O V;
DEC : D E C;

// whitespace, line breaks, comments, ...
LINE_CONTINUATION : [ \t]+ UNDERSCORE '\r'? '\n' -> skip;
NEWLINE : [\r\n\u2028\u2029]+;
REMCOMMENT : COLON? REM WS (LINE_CONTINUATION | ~[\r\n\u2028\u2029])*;
COMMENT : SINGLEQUOTE (LINE_CONTINUATION | ~[\r\n\u2028\u2029])*;
SINGLEQUOTE : '\'';
COLON : ':';
UNDERSCORE : '_';
WS : ([ \t] | LINE_CONTINUATION)+;

// identifier
IDENTIFIER :  (~[\[\]\(\)\r\n\t.,'"|!@#$%^&*-+:=; ])+ | L_SQUARE_BRACKET (~[!\]\r\n])+ R_SQUARE_BRACKET;


// letters
fragment LETTER : [a-zA-Z_äöüÄÖÜ];
fragment DIGIT : [0-9];
fragment LETTERORDIGIT : [a-zA-Z0-9_äöüÄÖÜ];

// case insensitive chars
fragment A:('a'|'A');
fragment B:('b'|'B');
fragment C:('c'|'C');
fragment D:('d'|'D');
fragment E:('e'|'E');
fragment F:('f'|'F');
fragment G:('g'|'G');
fragment H:('h'|'H');
fragment I:('i'|'I');
fragment J:('j'|'J');
fragment K:('k'|'K');
fragment L:('l'|'L');
fragment M:('m'|'M');
fragment N:('n'|'N');
fragment O:('o'|'O');
fragment P:('p'|'P');
fragment Q:('q'|'Q');
fragment R:('r'|'R');
fragment S:('s'|'S');
fragment T:('t'|'T');
fragment U:('u'|'U');
fragment V:('v'|'V');
fragment W:('w'|'W');
fragment X:('x'|'X');
fragment Y:('y'|'Y');
fragment Z:('z'|'Z');