grammar VBALike;

compilationUnit : likePatternString EOF;

likePatternString : likePatternElement*;
likePatternElement : likePatternChar | QUESTIONMARK | HASH | STAR | likePatternCharlist;
likePatternChar : ~(QUESTIONMARK | HASH | STAR | L_SQUARE_BRACKET);
likePatternCharlist : L_SQUARE_BRACKET EXCLAMATION? DASH? likePatternCharlistElement* DASH? R_SQUARE_BRACKET;
likePatternCharlistElement : likePatternCharlistChar | likePatternCharlistRange;
likePatternCharlistRange : likePatternCharlistChar DASH likePatternCharlistChar;
likePatternCharlistChar : ~(DASH | R_SQUARE_BRACKET);

QUESTIONMARK : '?';
HASH : '#';
STAR : '*';
L_SQUARE_BRACKET : '[';
R_SQUARE_BRACKET : ']';
DASH : '-';
EXCLAMATION : '!';
ANYCHAR : .;
