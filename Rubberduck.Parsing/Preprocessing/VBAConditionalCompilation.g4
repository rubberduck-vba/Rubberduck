grammar VBAConditionalCompilation;

compilationUnit : ccBlock EOF;

ccBlock : (ccConst | ccIfBlock | logicalLine)*;

ccConst : WS* HASHCONST WS+ ccVarLhs WS+ EQ WS+ ccExpression ccEol;

logicalLine : extendedLine+;

extendedLine : (lineContinuation | ~(HASHCONST | HASHIF | HASHELSEIF | HASHELSE | HASHENDIF))+ NEWLINE?;

lineContinuation : WS* UNDERSCORE WS* NEWLINE;

ccVarLhs : name;

ccExpression :
    L_PAREN WS* ccExpression WS* R_PAREN
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

ccIf : WS* HASHIF WS+ ccExpression WS+ THEN ccEol;

ccElseIfBlock : ccElseIf ccBlock;

ccElseIf : WS* HASHELSEIF WS+ ccExpression WS+ THEN ccEol;

ccElseBlock : ccElse ccBlock;

ccElse : WS* HASHELSE ccEol;

ccEndIf : WS* HASHENDIF ccEol;

ccEol : (SINGLEQUOTE ~NEWLINE*)? NEWLINE?;

intrinsicFunction : intrinsicFunctionName L_PAREN WS* ccExpression WS* R_PAREN;

intrinsicFunctionName : (INT | FIX | ABS | SGN | LEN | LENB | CBOOL | CBYTE | CCUR | CDATE | CDBL | CINT | CLNG | CLNGLNG | CLNGPTR | CSNG | CSTR | CVAR);

name : IDENTIFIER typeSuffix?;

typeSuffix : AMPERSAND | PERCENT | HASH | EXCLAMATIONMARK | AT | DOLLAR;

literal : HEXLITERAL | OCTLITERAL | DATELITERAL | DOUBLELITERAL | INTEGERLITERAL | SHORTLITERAL | STRINGLITERAL | TRUE | FALSE | NOTHING | NULL | EMPTY;

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
fragment TIMEVALUE : (DIGIT+ AMPM) | (DIGIT+ TIMESEPARATOR DIGIT+ (TIMESEPARATOR DIGIT+)? AMPM?);
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
NOT : N O T;
TRUE : T R U E;
FALSE : F A L S E;
NOTHING : N O T H I N G;
NULL : N U L L;
EMPTY : E M P T Y;
HASHCONST : WS* HASH CONST;
HASHIF : WS* HASH I F;
THEN : T H E N;
HASHELSEIF : WS* HASH E L S E I F;
HASHELSE : WS* HASH E L S E;
HASHENDIF : WS* HASH E N D WS* I F;
INT : I N T;
FIX : F I X;
ABS : A B S;
SGN : S G N;
LEN : L E N;
LENB : L E N B;
CBOOL : C B O O L;
CBYTE : C B Y T E;
CCUR : C C U R;
CDATE : C D A T E;
CDBL : C D B L;
CINT : C I N T;
CLNG : C L N G;
CLNGLNG : C L N G L N G;
CLNGPTR : C L N G P T R;
CSNG : C S N G;
CSTR : C S T R;
CVAR : C V A R;
IS : I S;
LIKE : L I K E;
MOD : M O D;
AND : A N D;
OR : O R;
XOR : X O R;
EQV : E Q V;
IMP : I M P;
CONST : C O N S T;
HASH : '#';
AMPERSAND : '&';
PERCENT : '%';
EXCLAMATIONMARK : '!';
AT : '@';
DOLLAR : '$';
L_PAREN : '(';
R_PAREN : ')';
L_SQUARE_BRACKET : '[';
R_SQUARE_BRACKET : ']';
UNDERSCORE : '_';
EQ : '=';
DIV : '/';
INTDIV : '\\';
GEQ : '>=';
GT : '>';
LEQ : '<=';
LT : '<';
MINUS : '-';
MULT : '*';
NEQ : '<>';
PLUS : '+';
POW : '^';
SINGLEQUOTE : '\'';
DOT : '.';
COMMA : ',';

NEWLINE : '\r' '\n' | [\r\n\u2028\u2029];
WS : [ \t];
IDENTIFIER :  ~[\[\]\(\)\r\n\t.,'"|!@#$%^&*-+:=; 0-9-/\\] ~[\[\]\(\)\r\n\t.,'"|!@#$%^&*-+:=; ]* | L_SQUARE_BRACKET (~[!\]\r\n])+ R_SQUARE_BRACKET;
fragment DIGIT : [0-9];

ANYCHAR : .;

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