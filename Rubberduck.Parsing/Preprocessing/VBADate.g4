grammar VBADate;

compilationUnit : dateLiteral EOF;

dateLiteral : HASH dateOrTime HASH;
dateOrTime :
    dateValue
    | timeValue
    | dateValue WS+ timeValue;
dateValue : dateValuePart dateSeparator dateValuePart (dateSeparator dateValuePart)?;
dateValuePart : DIGIT+ | monthName;
dateSeparator : WS+ | (WS* (SLASH | COMMA | DASH) WS*);
monthName : englishMonthName | englishMonthAbbreviation;
englishMonthName : J A N U A R Y | F E B R U A R Y | M A R C H | A P R I L | M A Y | J U N E  | A U G U S T | S E P T E M B E R | O C T O B E R | N O V E M B E R | D E C E M B E R;
englishMonthAbbreviation : J A N | F E B | M A R | A P R | J U N | J U L | A U G | S E P |  O C T | N O V | D E C;
timeValue : (timeValuePart WS* AMPM) | (timeValuePart timeSeparator timeValuePart (timeSeparator timeValuePart)? (WS* AMPM)?);
timeValuePart : DIGIT+;
timeSeparator : WS* (COLON | DOT) WS*;
AMPM : AM | PM | A | P;

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