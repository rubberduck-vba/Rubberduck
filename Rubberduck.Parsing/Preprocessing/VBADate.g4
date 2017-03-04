grammar VBADate;

compilationUnit : dateLiteral EOF;

dateLiteral : HASH ((WS | LINE_CONTINUATION)+)? dateOrTime HASH;
dateOrTime :
    dateValue
    | timeValue
    | dateValue WS+ timeValue;
dateValue : dateValuePart dateSeparator dateValuePart (dateSeparator dateValuePart)?;
dateValuePart : dateValueNumber | monthName;
dateValueNumber : DIGIT+;
dateSeparator : WS+ | (WS* (SLASH | COMMA | DASH) WS*);
monthName : englishMonthName | englishMonthAbbreviation;
englishMonthName : JANUARY | FEBRUARY | MARCH | APRIL | MAY | JUNE | JULY | AUGUST | SEPTEMBER | OCTOBER | NOVEMBER | DECEMBER;
// MAY is missing because abbreviation = full  name and it doesn't matter which one gets matched.
englishMonthAbbreviation : JAN | FEB | MAR | APR | JUN | JUL | AUG | SEP | OCT | NOV | DEC;
timeValue : (timeValuePart WS* AMPM) | (timeValuePart timeSeparator timeValuePart (timeSeparator timeValuePart)? (WS* AMPM)?);
timeValuePart : DIGIT+;
timeSeparator : WS* (COLON | DOT) WS*;
AMPM : AM | PM | A | P;

LINE_CONTINUATION : [ \t]* '_' [ \t]* '\r'? '\n';

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
AM : A M;
PM : P M;
HASH : '#';
COMMA : ',';
DASH : '-';
SLASH : '/';
COLON : ':';
DOT : '.';
WS : [ \t];
DIGIT : [0-9];
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
