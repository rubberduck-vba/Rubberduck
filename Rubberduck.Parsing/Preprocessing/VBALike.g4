grammar VBALike;

compilationUnit : likePatternString EOF;

likePatternString : likePatternElement*;
likePatternElement : NORMALCHAR | QUESTIONMARK | HASH | STAR | likePatternCharlist;
likePatternCharlist : L_SQUARE_BRACKET EXCLAMATION? DASH? likePatternCharlistElement* DASH? R_SQUARE_BRACKET;
likePatternCharlistElement : NORMALCHAR | likePatternCharlistRange;
likePatternCharlistRange : NORMALCHAR DASH NORMALCHAR;

QUESTIONMARK : '?';
HASH : '#';
STAR : '*';
L_SQUARE_BRACKET : '[';
R_SQUARE_BRACKET : ']';
DASH : '-';
EXCLAMATION : '!';
NORMALCHAR : ~[?#*[\]-!];