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

options { tokenVocab = VBALexer; }

startRule : module EOF;

module :
	endOfStatement
    moduleAttributes
	moduleHeader?
    moduleAttributes
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

moduleConfig :
	BEGIN (whiteSpace GUIDLITERAL whiteSpace unrestrictedIdentifier whiteSpace?)? endOfStatement
	moduleConfigElement+
	END endOfStatement
;

moduleConfigElement :
	unrestrictedIdentifier whiteSpace* EQ whiteSpace* expression (COLON numberLiteral)? endOfStatement
;

moduleAttributes : (attributeStmt endOfStatement)*;
attributeStmt : ATTRIBUTE whiteSpace attributeName whiteSpace? EQ whiteSpace? attributeValue (whiteSpace? COMMA whiteSpace? attributeValue)*;
attributeName : lExpression;
attributeValue : expression;

moduleDeclarations : (moduleDeclarationsElement endOfStatement)*;

moduleOption : 
	OPTION_BASE whiteSpace numberLiteral 					# optionBaseStmt
	| OPTION_COMPARE whiteSpace (BINARY | TEXT | DATABASE) 	# optionCompareStmt
	| OPTION_EXPLICIT 								        # optionExplicitStmt
	| OPTION_PRIVATE_MODULE 						        # optionPrivateModuleStmt
;

moduleDeclarationsElement :
    declareStmt
    | defDirective
	| enumerationStmt 
	| eventStmt
	| constStmt
	| implementsStmt
	| variableStmt
	| moduleOption
	| typeStmt
;

moduleBody : 
	(moduleBodyElement endOfStatement)*;

moduleBodyElement : 
	functionStmt 
	| propertyGetStmt 
	| propertySetStmt 
	| propertyLetStmt 
	| subStmt 
;

block : (blockStmt endOfStatement)*;

blockStmt :
	statementLabelDefinition
    | fileStmt
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
    | circleSpecialForm
    | scaleSpecialForm
    | callStmt
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
printStmt : PRINT whiteSpace markedFileNumber whiteSpace? COMMA (whiteSpace? outputList)?;

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

declareStmt : (visibility whiteSpace)? DECLARE whiteSpace (PTRSAFE whiteSpace)? (FUNCTION | SUB) whiteSpace identifier whiteSpace LIB whiteSpace STRINGLITERAL (whiteSpace ALIAS whiteSpace STRINGLITERAL)? (whiteSpace? argList)? (whiteSpace asTypeClause)?;

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
letterSpec : singleLetter | universalLetterRange | letterRange;
singleLetter : unrestrictedIdentifier;
// We make a separate universalLetterRange rule because it is treated specially in VBA. This makes it easy for users of the parser
// to identify this case. Quoting MS VBAL:
// "A <universal-letter-range> defines a single implicit declared type for every <IDENTIFIER> within 
// a module, even those with a first character that would otherwise fall outside this range if it was 
// interpreted as a <letter-range> from A-Z.""
universalLetterRange : upperCaseA whiteSpace? MINUS whiteSpace? upperCaseZ;
upperCaseA : {_input.Lt(1).Text.Equals("A")}? unrestrictedIdentifier;
upperCaseZ : {_input.Lt(1).Text.Equals("Z")}? unrestrictedIdentifier;
letterRange : firstLetter whiteSpace? MINUS whiteSpace? lastLetter;
firstLetter : unrestrictedIdentifier;
lastLetter : unrestrictedIdentifier;

doLoopStmt :
	DO endOfStatement 
	block
	LOOP
	|
	DO whiteSpace (WHILE | UNTIL) whiteSpace expression endOfStatement
	block
	LOOP
	| 
	DO endOfStatement
	block
	LOOP whiteSpace (WHILE | UNTIL) whiteSpace expression
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
	FOR whiteSpace EACH whiteSpace expression whiteSpace IN whiteSpace expression endOfStatement
	block
	NEXT (whiteSpace expression)?
;

// expression EQ expression refactored to expression to allow SLL
forNextStmt : 
	FOR whiteSpace expression whiteSpace TO whiteSpace expression (whiteSpace STEP whiteSpace expression)? endOfStatement 
	block
	NEXT (whiteSpace expression)?
; 

functionStmt :
	(visibility whiteSpace)? (STATIC whiteSpace)? FUNCTION whiteSpace? functionName (whiteSpace? argList)? (whiteSpace? asTypeClause)? endOfStatement
	block
	END_FUNCTION
;
functionName : identifier;

goSubStmt : GOSUB whiteSpace expression;

goToStmt : GOTO whiteSpace expression;

// 5.4.2.8 If Statement
ifStmt :
     IF whiteSpace booleanExpression whiteSpace THEN endOfStatement
     block
     elseIfBlock*
     elseBlock?
     END_IF
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
ifWithEmptyThen : IF whiteSpace? booleanExpression whiteSpace? THEN endOfStatement whiteSpace? singleLineElseClause;
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
sameLineStatement : blockStmt;
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
	END_PROPERTY
;

propertySetStmt : 
	(visibility whiteSpace)? (STATIC whiteSpace)? PROPERTY_SET whiteSpace subroutineName (whiteSpace? argList)? endOfStatement 
	block 
	END_PROPERTY
;

propertyLetStmt : 
	(visibility whiteSpace)? (STATIC whiteSpace)? PROPERTY_LET whiteSpace subroutineName (whiteSpace? argList)? endOfStatement 
	block 
	END_PROPERTY
;

// 5.4.2.20 RaiseEvent Statement
raiseEventStmt : RAISEEVENT whiteSpace identifier (whiteSpace? LPAREN whiteSpace? eventArgumentList? whiteSpace? RPAREN)?;
eventArgumentList : eventArgument (whiteSpace? COMMA whiteSpace? eventArgument)*;
eventArgument : expression;

// 5.4.3.3 ReDim Statement
// To make the grammar non-ambiguous we treat redim statements as index expressions.
// For this to work the argumentList rule had to be changed to support "lower bound arguments", e.g. "1 To 10".
redimStmt : REDIM whiteSpace (PRESERVE whiteSpace)? redimDeclarationList;
redimDeclarationList : redimVariableDeclaration (whiteSpace? COMMA whiteSpace? redimVariableDeclaration)*;
redimVariableDeclaration : expression (whiteSpace asTypeClause)?;

integerExpression : expression;

callStmt :
    CALL whiteSpace expression
    | expression (whiteSpace argumentList)?
;

resumeStmt : RESUME (whiteSpace (NEXT | expression))?;

returnStmt : RETURN;

rsetStmt : RSET whiteSpace expression whiteSpace? EQ whiteSpace? expression;

// 5.4.2.11 Stop Statement
stopStmt : STOP;

// 5.4.2.10 Select Case Statement
selectCaseStmt :
    SELECT whiteSpace? CASE whiteSpace? selectExpression endOfStatement
    caseClause*
    caseElseClause?
    END_SELECT
;
selectExpression : expression;
caseClause :
    CASE whiteSpace rangeClause (whiteSpace? COMMA whiteSpace? rangeClause)* endOfStatement block
;
caseElseClause : CASE whiteSpace? ELSE endOfStatement block;
rangeClause :
    expression
    | selectStartValue whiteSpace TO whiteSpace selectEndValue   
    | (IS whiteSpace?)? comparisonOperator whiteSpace? expression
;
selectStartValue : expression;
selectEndValue : expression;

setStmt : SET whiteSpace lExpression whiteSpace? EQ whiteSpace? expression;

subStmt : 
	(visibility whiteSpace)? (STATIC whiteSpace)? SUB whiteSpace? subroutineName (whiteSpace? argList)? endOfStatement
	block 
	END_SUB
;
subroutineName : identifier;

typeStmt : 
	(visibility whiteSpace)? TYPE whiteSpace identifier endOfStatement
	typeStmt_Element*
	END_TYPE
;

typeStmt_Element : identifier (whiteSpace? LPAREN (whiteSpace? subscripts)? whiteSpace? RPAREN)? (whiteSpace asTypeClause)? endOfStatement;

variableStmt : (DIM | STATIC | visibility) whiteSpace (WITHEVENTS whiteSpace)? variableListStmt;
variableListStmt : variableSubStmt (whiteSpace? COMMA whiteSpace? variableSubStmt)*;
variableSubStmt : identifier (whiteSpace? LPAREN whiteSpace? (subscripts whiteSpace?)? RPAREN whiteSpace?)? (whiteSpace asTypeClause)?;

whileWendStmt : 
	WHILE whiteSpace expression endOfStatement 
	block
	WEND
;

withStmt :
	WITH whiteSpace expression endOfStatement 
	block 
	END_WITH
;

// Special forms with special syntax, only available in a report.
circleSpecialForm : (expression whiteSpace? DOT whiteSpace?)? CIRCLE whiteSpace (STEP whiteSpace?)? tuple (whiteSpace? COMMA whiteSpace? expression)+;
scaleSpecialForm : (expression whiteSpace? DOT whiteSpace?)? SCALE whiteSpace tuple whiteSpace? MINUS whiteSpace? tuple;
tuple : LPAREN whiteSpace? expression whiteSpace? COMMA whiteSpace? expression whiteSpace? RPAREN;

subscripts : subscript (whiteSpace? COMMA whiteSpace? subscript)*;

subscript : (expression whiteSpace TO whiteSpace)? expression;

unrestrictedIdentifier : identifier | statementKeyword | markerKeyword;
identifier : untypedIdentifier | typedIdentifier;
untypedIdentifier : identifierValue;
typedIdentifier : identifierValue typeHint;
identifierValue : IDENTIFIER | keyword | foreignName;
foreignName : L_SQUARE_BRACKET foreignIdentifier* R_SQUARE_BRACKET;
foreignIdentifier : ~(L_SQUARE_BRACKET | R_SQUARE_BRACKET) | foreignName;

asTypeClause : AS whiteSpace? (NEW whiteSpace)? type (whiteSpace? fieldLength)?;

baseType : BOOLEAN | BYTE | CURRENCY | DATE | DOUBLE | INTEGER | LONG | LONGLONG | LONGPTR | SINGLE | STRING | VARIANT | ANY;

comparisonOperator : LT | LEQ | GT | GEQ | EQ | NEQ | IS | LIKE;

complexType : expression;

fieldLength : MULT whiteSpace? expression;

statementLabelDefinition : statementLabel whiteSpace? COLON;
statementLabel : identifierStatementLabel | lineNumberLabel;
identifierStatementLabel : unrestrictedIdentifier;
lineNumberLabel : numberLiteral;

numberLiteral : HEXLITERAL | OCTLITERAL | FLOATLITERAL | INTEGERLITERAL;

type : (baseType | complexType) (whiteSpace? LPAREN whiteSpace? RPAREN)?;

typeHint : PERCENT | AMPERSAND | POW | EXCLAMATIONPOINT | HASH | AT | DOLLAR;

visibility : PRIVATE | PUBLIC | FRIEND | GLOBAL;

// 5.6 Expressions
expression :
    // Literal Expression has to come before lExpression, otherwise it'll be classified as simple name expression instead.
    literalExpression                                                                               # literalExpr
    | lExpression                                                                                   # lExpr
    | builtInType                                                                                   # builtInTypeExpr
    | LPAREN whiteSpace? expression whiteSpace? RPAREN                                              # parenthesizedExpr
    | TYPEOF whiteSpace expression                                                                  # typeofexpr        // To make the grammar SLL, the type-of-is-expression is actually the child of an IS relational op.
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
    | HASH expression                                                                               # markedFileNumberExpr // Added to support special forms such as Input(file1, #file1)
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
    lExpression whiteSpace? LPAREN whiteSpace? argumentList? whiteSpace? RPAREN                                     # indexExpr
    | lExpression mandatoryLineContinuation? DOT mandatoryLineContinuation? unrestrictedIdentifier                  # memberAccessExpr
    | lExpression mandatoryLineContinuation? EXCLAMATIONPOINT mandatoryLineContinuation? unrestrictedIdentifier     # dictionaryAccessExpr
    | ME                                                                                                            # instanceExpr
    | identifier                                                                                                    # simpleNameExpr
    | DOT mandatoryLineContinuation? unrestrictedIdentifier                                                         # withMemberAccessExpr
    | EXCLAMATIONPOINT mandatoryLineContinuation? unrestrictedIdentifier                                            # withDictionaryAccessExpr
;

// 3.3.5.3 Special Identifier Forms
builtInType : 
    baseType
    | L_SQUARE_BRACKET whiteSpace? baseType whiteSpace? R_SQUARE_BRACKET
    | OBJECT
    | L_SQUARE_BRACKET whiteSpace? OBJECT whiteSpace? R_SQUARE_BRACKET
;

// 5.6.13.1 Argument Lists
argumentList : positionalOrNamedArgumentList;
positionalOrNamedArgumentList :
    (positionalArgumentOrMissing whiteSpace?)* requiredPositionalArgument 
    | (positionalArgumentOrMissing whiteSpace?)* namedArgumentList  
;
positionalArgumentOrMissing :
    positionalArgument whiteSpace? COMMA                                                            # specifiedPositionalArgument
    | whiteSpace? COMMA                                                                             # missingPositionalArgument
;
positionalArgument : argumentExpression;
requiredPositionalArgument : argumentExpression;  
namedArgumentList : namedArgument (whiteSpace? COMMA whiteSpace? namedArgument)*;
namedArgument : unrestrictedIdentifier whiteSpace? ASSIGN whiteSpace? argumentExpression;
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

endOfLine :
    whiteSpace? NEWLINE whiteSpace?
    | whiteSpace? commentOrAnnotation
;

endOfStatement :
    (endOfLine | (whiteSpace? COLON whiteSpace?))*
;

// Annotations must come before comments because of precedence. ANTLR4 matches as much as possible then chooses the one that comes first.
commentOrAnnotation :
    annotationList
    | comment
    | remComment
;
remComment : REM whiteSpace? commentBody;
comment : SINGLEQUOTE commentBody;
commentBody : (LINE_CONTINUATION | ~NEWLINE)*;
annotationList : SINGLEQUOTE (AT annotation whiteSpace?)+;
annotation : annotationName annotationArgList?;
annotationName : unrestrictedIdentifier;
annotationArgList : 
	 whiteSpace annotationArg
	 | whiteSpace annotationArg (whiteSpace? COMMA whiteSpace? annotationArg)+
	 | whiteSpace? LPAREN whiteSpace? RPAREN
	 | whiteSpace? LPAREN whiteSpace? annotationArg whiteSpace? RPAREN
	 | whiteSpace? LPAREN annotationArg (whiteSpace? COMMA whiteSpace? annotationArg)+ whiteSpace? RPAREN;
annotationArg : expression;

mandatoryLineContinuation : WS* LINE_CONTINUATION WS*;
whiteSpace : (WS | LINE_CONTINUATION)+;