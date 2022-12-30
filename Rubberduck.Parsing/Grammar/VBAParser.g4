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

/* VBA grammar based on MS VBAL */

parser grammar VBAParser;

options {
    tokenVocab = VBALexer;
    superClass = VBABaseParser;
    contextSuperClass = VBABaseParserRuleContext;
 }

startRule : module EOF;

module :
    endOfStatement?
    moduleAttributes
    moduleHeader?
    moduleAttributes
    moduleConfigReferences?
    moduleConfig?
    moduleAttributes
    moduleDeclarations
    moduleAttributes
    moduleBody
    moduleAttributes
    // A module can consist of WS as well as line continuations only.
    whiteSpace?
;

moduleHeader : VERSION whiteSpace numberLiteral whiteSpace? CLASS? endOfStatement;

moduleConfigReferences: moduleConfigReferenceElement+;

moduleConfigReferenceElement: 
    OBJECT whiteSpace? EQ whiteSpace? STRINGLITERAL whiteSpace? SEMICOLON whiteSpace? STRINGLITERAL endOfStatement
;

moduleConfig :
    BEGIN (whiteSpace (GUIDLITERAL | expression) whiteSpace unrestrictedIdentifier whiteSpace?)? endOfStatement
        (moduleConfig | moduleConfigProperty | moduleConfigElement)+
    END endOfStatement
;

moduleConfigProperty :
    BEGINPROPERTY whiteSpace unrestrictedIdentifier (LPAREN numberLiteral RPAREN)? (whiteSpace GUIDLITERAL)? endOfStatement
        (moduleConfigProperty | moduleConfigElement)*
    ENDPROPERTY endOfStatement
;

moduleConfigElement :
    (unrestrictedIdentifier | lExpression) whiteSpace? EQ whiteSpace? (shortcut | resource | expression | germanStyleFloatingPointNumber) endOfStatement
;

germanStyleFloatingPointNumber : 
    INTEGERLITERAL COMMA INTEGERLITERAL
    | COMMA INTEGERLITERAL
;

shortcut :
	(POW singleLetter)
	| ((PERCENT | PLUS? POW?) L_BRACE IDENTIFIER R_BRACE)
;

resource :
	DOLLAR? expression COLON (numberLiteral | BARE_HEX_LITERAL | unrestrictedIdentifier)
;

moduleAttributes : (attributeStmt endOfStatement)*;
attributeStmt : ATTRIBUTE whiteSpace attributeName whiteSpace? EQ whiteSpace? attributeValue (whiteSpace? COMMA whiteSpace? attributeValue)*;
attributeName : lExpression;
attributeValue : expression;

moduleDeclarations : (moduleDeclarationsElement endOfStatement)*;

moduleOption : 
    OPTION_BASE whiteSpace numberLiteral                       # optionBaseStmt
    | OPTION_COMPARE whiteSpace (BINARY | TEXT | DATABASE)     # optionCompareStmt
    | OPTION_EXPLICIT                                          # optionExplicitStmt
    | OPTION_PRIVATE_MODULE                                    # optionPrivateModuleStmt
;

moduleDeclarationsElement :
    whiteSpace?
    (attributeStmt
    | declareStmt
    | defDirective
    | enumerationStmt 
    | eventStmt
    | moduleConstStmt
    | implementsStmt
    | moduleVariableStmt
    | moduleOption
    | udtDeclaration)
;

moduleVariableStmt :
	variableStmt
	(endOfLine attributeStmt)*
;

moduleConstStmt :
	constStmt
	(endOfLine attributeStmt)*
;

moduleBody : 
    whiteSpace?
    ((moduleBodyElement | attributeStmt) endOfStatement)*;

moduleBodyElement : 
    functionStmt 
    | propertyGetStmt 
    | propertySetStmt 
    | propertyLetStmt 
    | subStmt 
;

block : (blockStmt endOfStatement)*;

unterminatedBlock : blockStmt (endOfStatement blockStmt)*;

blockStmt : 
    statementLabelDefinition whiteSpace? mainBlockStmt?
    | mainBlockStmt 
;

mainBlockStmt :
    fileStmt
    | attributeStmt
    | constStmt
    | doLoopStmt
    | endStmt
    | eraseStmt
    | errorStmt
    | exitStmt
    | forEachStmt
    | forNextStmt
    | goSubStmt
    | goToStmt
    | ifStmt
    | singleLineIfStmt
    | implementsStmt
    | midStatement
    | letStmt
    | lsetStmt
    | onErrorStmt
    | onGoToStmt
    | onGoSubStmt
    | raiseEventStmt
    | redimStmt
    | resumeStmt
    | returnStmt
    | rsetStmt
    | selectCaseStmt
    | setStmt
    | stopStmt
    | variableStmt
    | whileWendStmt
    | withStmt
    | lineSpecialForm
    | circleSpecialForm
    | scaleSpecialForm
    | pSetSpecialForm
    | unqualifiedObjectPrintStmt
    | callStmt
    | nameStmt
;

// 5.4.5 File Statements
fileStmt :
    openStmt
    | resetStmt
    | closeStmt
    | seekStmt
    | lockStmt
    | unlockStmt
    | lineInputStmt
    | widthStmt
    | printStmt
    | writeStmt
    | inputStmt
    | putStmt
    | getStmt
;


// 5.4.5.1 Open Statement
openStmt : OPEN whiteSpace pathName (whiteSpace modeClause)? (whiteSpace accessClause)? (whiteSpace lock)? whiteSpace AS whiteSpace fileNumber (whiteSpace lenClause)?;
pathName : expression;
modeClause : FOR whiteSpace fileMode;
fileMode : APPEND | BINARY | INPUT | OUTPUT | RANDOM;
accessClause : ACCESS whiteSpace access;
access :
    READ
    | WRITE
    | READ_WRITE
;
lock :
    SHARED
    | LOCK_READ
    | LOCK_WRITE
    | LOCK_READ_WRITE
;
lenClause : LEN whiteSpace? EQ whiteSpace? recLength;
recLength : expression;


// 5.4.5.1.1 File Numbers
fileNumber : markedFileNumber | unmarkedFileNumber;
markedFileNumber : HASH expression;
unmarkedFileNumber : expression;


// 5.4.5.2 Close and Reset Statements
closeStmt : CLOSE (whiteSpace fileNumberList)?;
resetStmt : RESET;
fileNumberList : fileNumber (whiteSpace? COMMA whiteSpace? fileNumber)*;


// 5.4.5.3 Seek Statement
seekStmt : SEEK whiteSpace fileNumber whiteSpace? COMMA whiteSpace? position;
position : expression;


// 5.4.5.4 Lock Statement
lockStmt : LOCK whiteSpace fileNumber (whiteSpace? COMMA whiteSpace? recordRange)?;
recordRange :
    startRecordNumber
    | (startRecordNumber whiteSpace)? TO whiteSpace endRecordNumber
;
startRecordNumber : expression;
endRecordNumber : expression;


// 5.4.5.5 Unlock Statement
unlockStmt : UNLOCK whiteSpace fileNumber (whiteSpace? COMMA whiteSpace? recordRange)?;


// 5.4.5.6 Line Input Statement
lineInputStmt : LINE_INPUT whiteSpace markedFileNumber whiteSpace? COMMA whiteSpace? variableName;
variableName : expression;


// 5.4.5.7 Width Statement
widthStmt : WIDTH whiteSpace markedFileNumber whiteSpace? COMMA whiteSpace? lineWidth;
lineWidth : expression;


// 5.4.5.8   Print Statement
//The unqualifiedObjectPrintStmt is an invocation of the Print member of the enclosing form or report, which also takes an output list as argument.
printMethod : PRINT;
printStmt : printMethod whiteSpace markedFileNumber whiteSpace? COMMA (whiteSpace? outputList)?;
unqualifiedObjectPrintStmt : printMethod (whiteSpace outputList)?;

// 5.4.5.8.1 Output Lists
outputList : outputItem (whiteSpace? outputItem)*;
outputItem :
    outputClause
    | charPosition
    | outputClause whiteSpace? charPosition
;
outputClause : spcClause | tabClause | outputExpression;
charPosition : SEMICOLON | COMMA;
outputExpression : expression;
spcClause : SPC whiteSpace? LPAREN whiteSpace? spcNumber whiteSpace? RPAREN;
spcNumber : expression; 
tabClause : TAB (whiteSpace? tabNumberClause)?;
tabNumberClause : LPAREN whiteSpace? tabNumber whiteSpace? RPAREN;
tabNumber : expression;


// 5.4.5.9 Write Statement
writeStmt : WRITE whiteSpace markedFileNumber whiteSpace? COMMA (whiteSpace? outputList)?;


// 5.4.5.10 Input Statement
inputStmt : INPUT whiteSpace markedFileNumber whiteSpace? COMMA whiteSpace? inputList;
inputList : inputVariable (whiteSpace? COMMA whiteSpace? inputVariable)*;  
inputVariable : expression;


// 5.4.5.11   Put Statement
putStmt : PUT whiteSpace fileNumber whiteSpace? COMMA whiteSpace? recordNumber? whiteSpace? COMMA whiteSpace? data;
recordNumber : expression;
data : expression;


// 5.4.5.12 Get Statement
getStmt : GET whiteSpace fileNumber whiteSpace? COMMA whiteSpace? recordNumber? whiteSpace? COMMA whiteSpace? variable; 
variable : expression;


constStmt : (visibility whiteSpace)? CONST whiteSpace constSubStmt (whiteSpace? COMMA whiteSpace? constSubStmt)*;
constSubStmt : identifier (whiteSpace asTypeClause)? whiteSpace? EQ whiteSpace? expression;

declareStmt : (visibility whiteSpace)? DECLARE whiteSpace (PTRSAFE whiteSpace)? (FUNCTION | SUB) whiteSpace identifier whiteSpace (CDECL whiteSpace)? LIB whiteSpace STRINGLITERAL (whiteSpace ALIAS whiteSpace STRINGLITERAL)? (whiteSpace? argList)? (whiteSpace asTypeClause)? (endOfLine attributeStmt)*;

argList : LPAREN (whiteSpace? arg (whiteSpace? COMMA whiteSpace? arg)*)? whiteSpace? RPAREN;

arg : (OPTIONAL whiteSpace)? ((BYVAL | BYREF) whiteSpace)? (PARAMARRAY whiteSpace)? unrestrictedIdentifier (whiteSpace? LPAREN whiteSpace? RPAREN)? (whiteSpace? asTypeClause)? (whiteSpace? argDefaultValue)?;

argDefaultValue : EQ whiteSpace? expression;

// 5.2.2 Implicit Definition Directives
defDirective : defType whiteSpace letterSpec (whiteSpace? COMMA whiteSpace? letterSpec)*;
defType :
    DEFBOOL | DEFBYTE | DEFINT | DEFLNG | DEFLNGLNG | DEFLNGPTR | DEFCUR |
    DEFSNG | DEFDBL | DEFDATE | 
    DEFSTR | DEFOBJ | DEFVAR
;
// universalLetterRange must appear before letterRange because they both match the same amount in the case of A-Z but we prefer the universalLetterRange.
// singleLetter must appear at the end to prevent premature bailout
letterSpec : universalLetterRange | letterRange | singleLetter;

singleLetter : {MatchesRegex(TextOf(TokenAtRelativePosition(1)),"^[a-zA-Z]$")}? IDENTIFIER;

// We make a separate universalLetterRange rule because it is treated specially in VBA. This makes it easy for users of the parser
// to identify this case. Quoting MS VBAL:
// "A <universal-letter-range> defines a single implicit declared type for every <IDENTIFIER> within 
// a module, even those with a first character that would otherwise fall outside this range if it was 
// interpreted as a <letter-range> from A-Z.""
universalLetterRange : {EqualsString(TextOf(TokenAtRelativePosition(1)),"A") && EqualsString(TextOf(TokenAtRelativePosition(3)),"Z")}? IDENTIFIER MINUS IDENTIFIER;
 
letterRange : singleLetter MINUS singleLetter;


doLoopStmt :
    DO endOfStatement 
    block
    statementLabelDefinition? whiteSpace? LOOP
    |
    DO whiteSpace (WHILE | UNTIL) whiteSpace expression endOfStatement
    block
    statementLabelDefinition? whiteSpace? LOOP
    | 
    DO endOfStatement
    block
    statementLabelDefinition? whiteSpace? LOOP whiteSpace (WHILE | UNTIL) whiteSpace expression
;

enumerationStmt: 
    (visibility whiteSpace)? ENUM whiteSpace identifier endOfStatement 
    enumerationStmt_Constant* 
    END_ENUM
;

enumerationStmt_Constant : identifier (whiteSpace? EQ whiteSpace? expression)? endOfStatement;

// We add "END" as a statement so that it does not get resolved to some nonsensical property.
endStmt : END;

eraseStmt : ERASE whiteSpace expression (whiteSpace? COMMA whiteSpace? expression)*;

errorStmt : ERROR whiteSpace expression;

eventStmt : (visibility whiteSpace)? EVENT whiteSpace identifier whiteSpace? argList;

exitStmt : EXIT_DO | EXIT_FOR | EXIT_FUNCTION | EXIT_PROPERTY | EXIT_SUB;

forEachStmt : 
    FOR whiteSpace EACH whiteSpace expression whiteSpace IN whiteSpace expression 
	(endOfStatement unterminatedBlock)?
    (endOfStatement statementLabelDefinition? whiteSpace? NEXT (whiteSpace expression)? 
	| whiteSpace? COMMA whiteSpace? expression)
;

// expression EQ expression refactored to expression to allow SLL
forNextStmt : 
    FOR whiteSpace expression whiteSpace TO whiteSpace expression stepStmt? whiteSpace* 
	(endOfStatement unterminatedBlock)?
    (endOfStatement statementLabelDefinition? whiteSpace? NEXT (whiteSpace expression)? 
	| whiteSpace? COMMA whiteSpace? expression)
; 

stepStmt : whiteSpace STEP whiteSpace expression;

functionStmt :
    (visibility whiteSpace)? (STATIC whiteSpace)? FUNCTION whiteSpace? functionName (whiteSpace? argList)? (whiteSpace? asTypeClause)? endOfStatement
    block
    statementLabelDefinition? whiteSpace? END_FUNCTION
    (endOfLine attributeStmt)*
;
functionName : identifier;

goSubStmt : GOSUB whiteSpace expression;

goToStmt : GOTO whiteSpace expression;

// 5.4.2.8 If Statement
ifStmt :
    IF whiteSpace booleanExpression whiteSpace THEN endOfStatement
    block
    (statementLabelDefinition? whiteSpace? elseIfBlock)*
    (statementLabelDefinition? whiteSpace? elseBlock?)
    statementLabelDefinition? whiteSpace? END_IF
;
elseIfBlock : 
    ELSEIF whiteSpace booleanExpression whiteSpace THEN endOfStatement block
    | ELSEIF whiteSpace booleanExpression whiteSpace THEN whiteSpace? block
;
elseBlock :
    ELSE endOfStatement block
;

// 5.4.2.9 Single-line If Statement
singleLineIfStmt : ifWithNonEmptyThen | ifWithEmptyThen;
ifWithNonEmptyThen : IF whiteSpace? booleanExpression whiteSpace? THEN whiteSpace? listOrLabel (whiteSpace singleLineElseClause)?;
ifWithEmptyThen : IF whiteSpace? booleanExpression whiteSpace? THEN whiteSpace? emptyThenStatement? singleLineElseClause;
singleLineElseClause : ELSE whiteSpace? listOrLabel?;

// lineNumberLabel should actually be "statement-label" according to MS VBAL but they only allow lineNumberLabels:
// A <statement-label> that occurs as the first element of a <list-or-label> element has the effect 
// as if the <statement-label> was replaced with a <goto-statement> containing the same 
// <statement-label>. This <goto-statement> takes the place of <line-number-label> in 
// <statement-list>.  
listOrLabel :
    lineNumberLabel (whiteSpace? COLON whiteSpace? sameLineStatement?)*
    | (COLON whiteSpace?)? sameLineStatement (whiteSpace? COLON whiteSpace? sameLineStatement?)*
;

sameLineStatement : mainBlockStmt;
emptyThenStatement : (COLON whiteSpace?)+;
booleanExpression : expression;

implementsStmt : IMPLEMENTS whiteSpace expression;

letStmt : (LET whiteSpace)? lExpression whiteSpace? EQ whiteSpace? expression;

lsetStmt : LSET whiteSpace expression whiteSpace? EQ whiteSpace? expression;

onErrorStmt : (ON_ERROR | ON_LOCAL_ERROR) whiteSpace (GOTO whiteSpace expression | RESUME whiteSpace NEXT);

onGoToStmt : ON whiteSpace expression whiteSpace GOTO whiteSpace expression (whiteSpace? COMMA whiteSpace? expression)*;

onGoSubStmt : ON whiteSpace expression whiteSpace GOSUB whiteSpace expression (whiteSpace? COMMA whiteSpace? expression)*;

propertyGetStmt : 
    (visibility whiteSpace)? (STATIC whiteSpace)? PROPERTY_GET whiteSpace functionName (whiteSpace? argList)? (whiteSpace asTypeClause)? endOfStatement 
    block 
    statementLabelDefinition? whiteSpace? END_PROPERTY
    (endOfLine attributeStmt)*
;

propertySetStmt : 
    (visibility whiteSpace)? (STATIC whiteSpace)? PROPERTY_SET whiteSpace subroutineName (whiteSpace? argList)? endOfStatement 
    block 
    statementLabelDefinition? whiteSpace? END_PROPERTY
    (endOfLine attributeStmt)*
;

propertyLetStmt : 
    (visibility whiteSpace)? (STATIC whiteSpace)? PROPERTY_LET whiteSpace subroutineName (whiteSpace? argList)? endOfStatement 
    block 
    statementLabelDefinition? whiteSpace? END_PROPERTY
    (endOfLine attributeStmt)*
;

// 5.4.2.20 RaiseEvent Statement
raiseEventStmt : RAISEEVENT whiteSpace identifier (whiteSpace? LPAREN whiteSpace? eventArgumentList? whiteSpace? RPAREN)?;
eventArgumentList : eventArgument (whiteSpace? COMMA whiteSpace? eventArgument)*;
eventArgument : (BYVAL whiteSpace)? expression;

// 5.4.3.3 ReDim Statement
// To make the grammar non-ambiguous we treat redim statements as index expressions.
// For this to work the argumentList rule had to be changed to support "lower bound arguments", e.g. "1 To 10".
redimStmt : REDIM whiteSpace (PRESERVE whiteSpace)? redimDeclarationList;
redimDeclarationList : redimVariableDeclaration (whiteSpace? COMMA whiteSpace? redimVariableDeclaration)*;
redimVariableDeclaration : expression (whiteSpace asTypeClause)?;

// 5.4.3.5 Mid/MidB/Mid$/MidB$ Statement
// This needs to be explicitly defined to distinguish between Mid as a function and Mid as a keyword.
midStatement : modeSpecifier 
    LPAREN whiteSpace? 
    lExpression whiteSpace? COMMA whiteSpace? expression whiteSpace? (COMMA whiteSpace? expression whiteSpace?)? 
    RPAREN 
    whiteSpace? EQ whiteSpace? 
    expression;
modeSpecifier :	(MID | MIDB) DOLLAR? ;

integerExpression : expression;

callStmt :
    CALL whiteSpace lExpression
    | lExpression (whiteSpace argumentList)?
;

resumeStmt : RESUME (whiteSpace (NEXT | expression))?;

returnStmt : RETURN;

rsetStmt : RSET whiteSpace expression whiteSpace? EQ whiteSpace? expression;

// 5.4.2.11 Stop Statement
stopStmt : STOP;

nameStmt : NAME whiteSpace expression whiteSpace AS whiteSpace expression;

// 5.4.2.10 Select Case Statement
selectCaseStmt :
    SELECT whiteSpace? CASE whiteSpace? selectExpression endOfStatement
    (statementLabelDefinition? whiteSpace? caseClause)*
    statementLabelDefinition? whiteSpace? caseElseClause?
    statementLabelDefinition? whiteSpace? END_SELECT
;
selectExpression : expression;
caseClause :
    CASE whiteSpace rangeClause (whiteSpace? COMMA whiteSpace? rangeClause)* endOfStatement block
;
caseElseClause : CASE whiteSpace? ELSE endOfStatement block;
rangeClause :
    (IS whiteSpace?)? comparisonOperator whiteSpace? expression
    | selectStartValue whiteSpace TO whiteSpace selectEndValue 
    | expression
;
selectStartValue : expression;
selectEndValue : expression;

setStmt : SET whiteSpace lExpression whiteSpace? EQ whiteSpace? expression;

subStmt : 
    (visibility whiteSpace)? (STATIC whiteSpace)? SUB whiteSpace? subroutineName (whiteSpace? argList)? endOfStatement
    block 
    statementLabelDefinition? whiteSpace? END_SUB
    (endOfLine attributeStmt)*
;
subroutineName : identifier;

// 5.2.3.3 User Defined Type Declarations
// member list includes trailing endOfStatement
udtDeclaration : (visibility whiteSpace)? TYPE whiteSpace untypedIdentifier endOfStatement udtMemberList END_TYPE;  
udtMemberList : (udtMember endOfStatement)+; 
udtMember : reservedNameMemberDeclaration | untypedNameMemberDeclaration;
untypedNameMemberDeclaration : untypedIdentifier whiteSpace? optionalArrayClause;
reservedNameMemberDeclaration : unrestrictedIdentifier whiteSpace asTypeClause;
optionalArrayClause : (arrayDim whiteSpace)? asTypeClause;

// 5.2.3.1.3 Array Dimensions and Bounds
arrayDim : LPAREN whiteSpace? boundsList? whiteSpace? RPAREN;
boundsList : dimSpec (whiteSpace? COMMA whiteSpace? dimSpec)*;
dimSpec : (lowerBound whiteSpace?)? upperBound;
lowerBound : constantExpression whiteSpace? TO;
upperBound : constantExpression;

constantExpression : expression;

variableStmt : (DIM | STATIC | visibility) whiteSpace variableListStmt;
variableListStmt : variableSubStmt (whiteSpace? COMMA whiteSpace? variableSubStmt)*;
variableSubStmt : (WITHEVENTS whiteSpace)? identifier (whiteSpace? arrayDim)? (whiteSpace asTypeClause)?;

whileWendStmt : 
    WHILE whiteSpace expression endOfStatement 
    block
    statementLabelDefinition? whiteSpace? WEND
;

withStmt :
    WITH whiteSpace expression endOfStatement 
    block 
    statementLabelDefinition? whiteSpace? END_WITH
;

// Special forms with special syntax, only available in VBA reports or VB6 forms and pictureboxes.
// lineSpecialFormOption is required if expression is missing
lineSpecialForm : expression whiteSpace ((STEP whiteSpace?)? tuple)?
    whiteSpace? MINUS whiteSpace?
	(STEP whiteSpace?)? tuple whiteSpace?
	(COMMA whiteSpace? expression? whiteSpace?)?
	(COMMA whiteSpace? lineSpecialFormOption)?
;
circleSpecialForm : (expression whiteSpace? DOT whiteSpace?)? CIRCLE whiteSpace (STEP whiteSpace?)? tuple whiteSpace? COMMA whiteSpace? expression (whiteSpace? COMMA whiteSpace? expression?)*;
scaleSpecialForm : (expression whiteSpace? DOT whiteSpace?)? SCALE whiteSpace tuple whiteSpace? MINUS whiteSpace? tuple;
pSetSpecialForm : (expression whiteSpace? DOT whiteSpace?)? PSET (whiteSpace STEP)? whiteSpace? tuple whiteSpace? (COMMA whiteSpace? expression)?;
tuple : LPAREN whiteSpace? expression whiteSpace? COMMA whiteSpace? expression whiteSpace? RPAREN;
lineSpecialFormOption : {EqualsStringIgnoringCase(TextOf(TokenAtRelativePosition(1)),"b","bf")}? unrestrictedIdentifier;

unrestrictedIdentifier : identifier | statementKeyword | markerKeyword;
legalLabelIdentifier : { !IsTokenType(TokenTypeAtRelativePosition(1),DOEVENTS,END,CLOSE,ELSE,LOOP,NEXT,RANDOMIZE,REM,RESUME,RETURN,STOP,WEND)}? identifier | markerKeyword;
//The predicate in the following rule has been introduced to lessen the problem that VBA uses the same characters used as type hints in other syntactical constructs, 
//e.g. in the bang notation (see withDictionaryAccessExpr). Generally, it is not legal to have an identifier or opening bracket follow immediately after a type hint.
//The first part of the predicate tries to exclude these two situations. Unfortunately, predicates have to be at the start of a rule. So, an assumption about the number 
//of tokens in the identifier is made. All untypedIdentifers not a foreignNames consist of exactly one token and a typedIdentifier is an untyped one followed by a typeHint,
//again a single token. So, in the majority of situations, the third token is the token following the potential type hint. 
//For foreignNames, no assumption can be made because they consist of a pair of brackets containing arbitrarily many tokens. 
//That is why the second part of the predicate looks at the first character in order to determine whether the identifier is a foreignName. 
identifier : {!IsTokenType(TokenTypeAtRelativePosition(3),IDENTIFIER,L_SQUARE_BRACKET) || IsTokenType(TokenTypeAtRelativePosition(1),L_SQUARE_BRACKET)}? typedIdentifier
             | untypedIdentifier;
untypedIdentifier : identifierValue;
typedIdentifier : untypedIdentifier typeHint;
identifierValue : IDENTIFIER | keyword | foreignName;
foreignName : L_SQUARE_BRACKET foreignIdentifier* R_SQUARE_BRACKET;
foreignIdentifier : ~(L_SQUARE_BRACKET | R_SQUARE_BRACKET) | foreignName;

asTypeClause : AS whiteSpace? (NEW whiteSpace)? type (whiteSpace? fieldLength)?;

baseType : BOOLEAN | BYTE | CURRENCY | DATE | DOUBLE | INTEGER | LONG | LONGLONG | LONGPTR | SINGLE | STRING | VARIANT | ANY;

comparisonOperator : LT | LEQ | GT | GEQ | EQ | NEQ | IS | LIKE;

complexType :
    // Literal Expression has to come before lExpression, otherwise it'll be classified as simple name expression instead.
    literalExpression                                                                               # ctLiteralExpr
    | lExpression                                                                                   # ctLExpr
    | builtInType                                                                                   # ctBuiltInTypeExpr
    | LPAREN whiteSpace? complexType whiteSpace? RPAREN                                             # ctParenthesizedExpr
    | TYPEOF whiteSpace complexType                                                                 # ctTypeofexpr        // To make the grammar SLL, the type-of-is-expression is actually the child of an IS relational op.
    | NEW whiteSpace complexType                                                                    # ctNewExpr
    | HASH expression                                                                               # ctMarkedFileNumberExpr // Added to support special forms such as Input(file1, #file1)
;

fieldLength : MULT whiteSpace? (numberLiteral | identifierValue);

//Statement labels can only appear at the start of a line.
statementLabelDefinition : {IsTokenType(TokenTypeAtRelativePosition(-1),NEWLINE,LINE_CONTINUATION)}? (combinedLabels | identifierStatementLabel | standaloneLineNumberLabel);
identifierStatementLabel : legalLabelIdentifier whiteSpace? COLON;
standaloneLineNumberLabel : 
    lineNumberLabel whiteSpace? COLON
    | lineNumberLabel;
combinedLabels : lineNumberLabel whiteSpace identifierStatementLabel;
// Technically, the negative numbers are illegal but VBE can prettify a
// &HFFFFFFFF into a -1 which becomes a legal line number. Editing the same
// line subsequently then breaks it. 
lineNumberLabel : MINUS? numberLiteral; 

numberLiteral : HEXLITERAL | OCTLITERAL | FLOATLITERAL | INTEGERLITERAL;

type : (baseType | complexType) (whiteSpace? LPAREN whiteSpace? RPAREN)?;

typeHint : PERCENT | AMPERSAND | POW | EXCLAMATIONPOINT | HASH | AT | DOLLAR;

visibility : PRIVATE | PUBLIC | FRIEND | GLOBAL;

// 5.6 Expressions
expression :
    // Literal Expression has to come before lExpression, otherwise it'll be classified as simple name expression instead.
    //The same holds for Built-in Type Expression.
    whiteSpace? LPAREN whiteSpace? expression whiteSpace? RPAREN                                    # parenthesizedExpr
    | TYPEOF whiteSpace expression                                                                  # typeofexpr // To make the grammar SLL, the type-of-is-expression is actually the child of an IS relational op.
    | HASH expression                                                                               # markedFileNumberExpr // Added to support special forms such as Input(file1, #file1)
    | NEW whiteSpace expression                                                                     # newExpr
    | expression whiteSpace? POW whiteSpace? expression                                             # powOp
    | MINUS whiteSpace? expression                                                                  # unaryMinusOp
    | expression whiteSpace? (MULT | DIV) whiteSpace? expression                                    # multOp
    | expression whiteSpace? INTDIV whiteSpace? expression                                          # intDivOp
    | expression whiteSpace? MOD whiteSpace? expression                                             # modOp
    | expression whiteSpace? (PLUS | MINUS) whiteSpace? expression                                  # addOp
    | expression whiteSpace? AMPERSAND whiteSpace? expression                                       # concatOp
    | expression whiteSpace? (EQ | NEQ | LT | GT | LEQ | GEQ | LIKE | IS) whiteSpace? expression    # relationalOp
    | NOT whiteSpace? expression                                                                    # logicalNotOp
    | expression whiteSpace? AND whiteSpace? expression                                             # logicalAndOp
    | expression whiteSpace? OR whiteSpace? expression                                              # logicalOrOp
    | expression whiteSpace? XOR whiteSpace? expression                                             # logicalXorOp
    | expression whiteSpace? EQV whiteSpace? expression                                             # logicalEqvOp
    | expression whiteSpace? IMP whiteSpace? expression                                             # logicalImpOp
    | literalExpression                                                                             # literalExpr
    | {!IsTokenType(TokenTypeAtRelativePosition(2),LPAREN)}? builtInType                            # builtInTypeExpr
    | lExpression                                                                                   # lExpr
;

// 5.6.5 Literal Expressions
literalExpression :
    numberLiteral
    | DATELITERAL
    | STRINGLITERAL
    | literalIdentifier typeHint?
;

literalIdentifier : booleanLiteralIdentifier | objectLiteralIdentifier | variantLiteralIdentifier;
booleanLiteralIdentifier : TRUE | FALSE;
objectLiteralIdentifier : NOTHING;
variantLiteralIdentifier : EMPTY | NULL;

lExpression :
    lExpression LPAREN whiteSpace? argumentList? whiteSpace? RPAREN                                                 # indexExpr
    | lExpression mandatoryLineContinuation? DOT mandatoryLineContinuation? printMethod (whiteSpace outputList)?    # objectPrintExpr
    | lExpression mandatoryLineContinuation? DOT mandatoryLineContinuation? unrestrictedIdentifier                  # memberAccessExpr
    | lExpression mandatoryLineContinuation? dictionaryAccess mandatoryLineContinuation? unrestrictedIdentifier     # dictionaryAccessExpr
    | ME                                                                                                            # instanceExpr
    | identifier                                                                                                    # simpleNameExpr
    | DOT mandatoryLineContinuation? unrestrictedIdentifier                                                         # withMemberAccessExpr
    | dictionaryAccess mandatoryLineContinuation? unrestrictedIdentifier                                            # withDictionaryAccessExpr
    | lExpression mandatoryLineContinuation whiteSpace? LPAREN whiteSpace? argumentList? whiteSpace? RPAREN         # whitespaceIndexExpr
;

//This is a hack to allow attaching identifier references for default members to the exclaramtion mark.
dictionaryAccess : EXCLAMATIONPOINT;

// 3.3.5.3 Special Identifier Forms
builtInType : 
    baseType
    | L_SQUARE_BRACKET whiteSpace? baseType whiteSpace? R_SQUARE_BRACKET
    | OBJECT
    | L_SQUARE_BRACKET whiteSpace? OBJECT whiteSpace? R_SQUARE_BRACKET
;

// 5.6.13.1 Argument Lists
argumentList : whiteSpace? (argument (whiteSpace? COMMA whiteSpace? argument)*)? whiteSpace?
;

requiredArgument : argument;
argument :
    positionalArgument
    | namedArgument
    | missingArgument
;

positionalArgument : argumentExpression;
namedArgument : unrestrictedIdentifier whiteSpace? ASSIGN whiteSpace? argumentExpression;
missingArgument : ;

argumentExpression :
    (BYVAL whiteSpace)? expression
    | addressOfExpression
    // Special case for redim statements. The resolver doesn't have to deal with this because it is "picked apart" in the redim statement.
    | lowerBoundArgumentExpression whiteSpace TO whiteSpace upperBoundArgumentExpression
;

lowerBoundArgumentExpression : expression;
upperBoundArgumentExpression : expression;

// 5.6.16.8   AddressOf Expressions 
addressOfExpression : ADDRESSOF whiteSpace expression;

keyword : 
      ABS
    | ADDRESSOF
    | ALIAS
    | AND
    | ANY
    | ARRAY
    | ATTRIBUTE
    | BEGIN
    | BEGINPROPERTY
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
    | ENDPROPERTY
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
    | LONG
    | LONGLONG
    | LONGPTR
    | ME
    | MID
    | MIDB
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
    | ERROR
    | APPEND
    | BINARY
    | OUTPUT
    | RANDOM
    | RANDOMIZE
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
    | CIRCLE
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
    | PSET
    | PUBLIC
    | RAISEEVENT
    | REDIM
    | RESUME
    | RETURN
    | RSET
    | SCALE
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

endOfLine :
    whiteSpace? NEWLINE
    | whiteSpace? commentOrAnnotation
;

// We expect endOfStatement to consume all trailing whitespace blank statements.
// We have to special case the end of file since infiniftly mant EOF tokens can be consumed at the end of file.
endOfStatement :
    individualNonEOFEndOfStatement+ | whiteSpace? EOF
;

// we expect endOfStatement to consume all trailing whitespace
individualNonEOFEndOfStatement :
	  endOfLine whiteSpace? 
	| whiteSpace? COLON whiteSpace?
;

// Annotations must come before comments because of precedence. ANTLR4 matches as much as possible then chooses the one that comes first.
commentOrAnnotation :
    (annotationList 
    | remComment
    | comment) 
    // all comments must end with a logical line. See VBA Language Spec 3.3.1
    (NEWLINE | EOF)
;
remComment : REM whiteSpace? commentBody;
comment : SINGLEQUOTE commentBody;
commentBody : (~NEWLINE)*;
annotationList : SINGLEQUOTE (AT annotation)+ (COLON commentBody)?;
annotation : annotationName annotationArgList? whiteSpace?;
annotationName : unrestrictedIdentifier;
annotationArgList : 
    whiteSpace? LPAREN whiteSpace? annotationArg whiteSpace? RPAREN
    | whiteSpace? LPAREN whiteSpace? RPAREN
    | whiteSpace? LPAREN annotationArg (whiteSpace? COMMA whiteSpace? annotationArg)+ whiteSpace? RPAREN
    | whiteSpace annotationArg
    | whiteSpace annotationArg (whiteSpace? COMMA whiteSpace? annotationArg)+
;
annotationArg : expression;

mandatoryLineContinuation : LINE_CONTINUATION WS*;
whiteSpace : (WS | LINE_CONTINUATION)+;