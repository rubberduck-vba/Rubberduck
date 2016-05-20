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

parser grammar VBAParser;

options { tokenVocab = VBALexer; }

startRule : module;

module : 
	whiteSpace?
	endOfStatement
	(moduleHeader endOfStatement)?
	moduleConfig? endOfStatement
	moduleAttributes? endOfStatement
	moduleDeclarations? endOfStatement
	moduleBody? endOfStatement
;

moduleHeader : VERSION whiteSpace numberLiteral whiteSpace? CLASS? endOfStatement;

moduleConfig :
	BEGIN (whiteSpace GUIDLITERAL whiteSpace unrestrictedIdentifier whiteSpace?)? endOfStatement
	moduleConfigElement+
	END
;

moduleConfigElement :
	unrestrictedIdentifier whiteSpace* EQ whiteSpace* valueStmt (COLON numberLiteral)? endOfStatement
;

moduleAttributes : (attributeStmt endOfStatement)+;

moduleDeclarations : moduleDeclarationsElement (endOfStatement moduleDeclarationsElement)* endOfStatement;

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
	moduleBodyElement (endOfStatement moduleBodyElement)* endOfStatement;

moduleBodyElement : 
	functionStmt 
	| propertyGetStmt 
	| propertySetStmt 
	| propertyLetStmt 
	| subStmt 
;

attributeStmt : ATTRIBUTE whiteSpace attributeName whiteSpace? EQ whiteSpace? attributeValue (whiteSpace? COMMA whiteSpace? attributeValue)*;
attributeName : implicitCallStmt_InStmt;
attributeValue : valueStmt;

block : blockStmt (endOfStatement blockStmt)* endOfStatement;

blockStmt :
	statementLabelDefinition
    // Put before the implicitCallStmt_InBlock rule so that RESET statements are not treated as function calls but as resetStmt.
    | fileStmt
	| attributeStmt
	| constStmt
	| doLoopStmt
    | endStmt
	| eraseStmt
	| errorStmt
    | exitStmt
	| explicitCallStmt
	| forEachStmt
	| forNextStmt
	| goSubStmt
	| goToStmt
	| ifStmt
    | singleLineIfStmt
	| implementsStmt
	| letStmt
	| lsetStmt
	| midStmt
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
	| implicitCallStmt_InBlock
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
pathName : valueStmt;
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
recLength : valueStmt;


// 5.4.5.1.1 File Numbers
fileNumber : markedFileNumber | unmarkedFileNumber;
markedFileNumber : HASH valueStmt;
unmarkedFileNumber : valueStmt;


// 5.4.5.2 Close and Reset Statements
closeStmt : CLOSE (whiteSpace fileNumberList)?;
resetStmt : RESET;
fileNumberList : fileNumber (whiteSpace? COMMA whiteSpace? fileNumber)*;


// 5.4.5.3 Seek Statement
seekStmt : SEEK whiteSpace fileNumber whiteSpace? COMMA whiteSpace? position;
position : valueStmt;


// 5.4.5.4 Lock Statement
lockStmt : LOCK whiteSpace fileNumber (whiteSpace? COMMA whiteSpace? recordRange)?;
recordRange :
    startRecordNumber
    | (startRecordNumber whiteSpace)? TO whiteSpace endRecordNumber
;
startRecordNumber : valueStmt;
endRecordNumber : valueStmt;


// 5.4.5.5 Unlock Statement
unlockStmt : UNLOCK whiteSpace fileNumber (whiteSpace? COMMA whiteSpace? recordRange)?;


// 5.4.5.6 Line Input Statement
lineInputStmt : LINE_INPUT whiteSpace markedFileNumber whiteSpace? COMMA whiteSpace? variableName;
variableName : valueStmt;


// 5.4.5.7 Width Statement
widthStmt : WIDTH whiteSpace markedFileNumber whiteSpace? COMMA whiteSpace? lineWidth;
lineWidth : valueStmt;


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
outputExpression : valueStmt;
spcClause : SPC whiteSpace? LPAREN whiteSpace? spcNumber whiteSpace? RPAREN;
spcNumber : valueStmt; 
tabClause : TAB (whiteSpace? tabNumberClause)?;
tabNumberClause : LPAREN whiteSpace? tabNumber whiteSpace? RPAREN;
tabNumber : valueStmt;


// 5.4.5.9 Write Statement
writeStmt : WRITE whiteSpace markedFileNumber whiteSpace? COMMA (whiteSpace? outputList)?;


// 5.4.5.10 Input Statement
inputStmt : INPUT whiteSpace markedFileNumber whiteSpace? COMMA whiteSpace? inputList;
inputList : inputVariable (whiteSpace? COMMA whiteSpace? inputVariable)*;  
inputVariable : valueStmt;


// 5.4.5.11   Put Statement
putStmt : PUT whiteSpace fileNumber whiteSpace? COMMA whiteSpace? recordNumber? whiteSpace? COMMA whiteSpace? data;
recordNumber : valueStmt;
data : valueStmt;


// 5.4.5.12 Get Statement
getStmt : GET whiteSpace fileNumber whiteSpace? COMMA whiteSpace? recordNumber? whiteSpace? COMMA whiteSpace? variable; 
variable : valueStmt;



constStmt : (visibility whiteSpace)? CONST whiteSpace constSubStmt (whiteSpace? COMMA whiteSpace? constSubStmt)*;

constSubStmt : identifier typeHint? (whiteSpace asTypeClause)? whiteSpace? EQ whiteSpace? valueStmt;

declareStmt : (visibility whiteSpace)? DECLARE whiteSpace (PTRSAFE whiteSpace)? (FUNCTION | SUB) whiteSpace identifier typeHint? whiteSpace LIB whiteSpace STRINGLITERAL (whiteSpace ALIAS whiteSpace STRINGLITERAL)? (whiteSpace? argList)? (whiteSpace asTypeClause)?;

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
	block?
	LOOP
	|
	DO whiteSpace (WHILE | UNTIL) whiteSpace valueStmt endOfStatement
	block?
	LOOP
	| 
	DO endOfStatement
	block?
	LOOP whiteSpace (WHILE | UNTIL) whiteSpace valueStmt
;

enumerationStmt: 
	(visibility whiteSpace)? ENUM whiteSpace identifier endOfStatement 
	enumerationStmt_Constant* 
	END_ENUM
;

enumerationStmt_Constant : identifier (whiteSpace? EQ whiteSpace? valueStmt)? endOfStatement;

// We add "END" as a statement so that it does not get resolved to some nonsensical property.
endStmt : END;

eraseStmt : ERASE whiteSpace valueStmt (whiteSpace? COMMA whiteSpace? valueStmt)*;

errorStmt : ERROR whiteSpace valueStmt;

eventStmt : (visibility whiteSpace)? EVENT whiteSpace identifier whiteSpace? argList;

exitStmt : EXIT_DO | EXIT_FOR | EXIT_FUNCTION | EXIT_PROPERTY | EXIT_SUB;

forEachStmt : 
	FOR whiteSpace EACH whiteSpace valueStmt whiteSpace IN whiteSpace valueStmt endOfStatement
	block?
	NEXT (whiteSpace valueStmt)?
;

forNextStmt : 
	FOR whiteSpace valueStmt whiteSpace? EQ whiteSpace? valueStmt whiteSpace TO whiteSpace valueStmt (whiteSpace STEP whiteSpace valueStmt)? endOfStatement 
	block?
	NEXT (whiteSpace valueStmt)?
; 

functionStmt :
	(visibility whiteSpace)? (STATIC whiteSpace)? FUNCTION whiteSpace? functionName typeHint? (whiteSpace? argList)? (whiteSpace? asTypeClause)? endOfStatement
	block?
	END_FUNCTION
;
functionName : identifier;

goSubStmt : GOSUB whiteSpace valueStmt;

goToStmt : GOTO whiteSpace valueStmt;

// 5.4.2.8 If Statement
ifStmt :
     IF whiteSpace booleanExpression whiteSpace THEN endOfStatement
     block?
     elseIfBlock*
     elseBlock?
     END_IF
;
elseIfBlock :
     ELSEIF whiteSpace booleanExpression whiteSpace THEN endOfStatement block?
     | ELSEIF whiteSpace booleanExpression whiteSpace THEN whiteSpace? block?
;
elseBlock :
     ELSE endOfStatement block?
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
booleanExpression : valueStmt;

implementsStmt : IMPLEMENTS whiteSpace valueStmt;

letStmt : (LET whiteSpace)? valueStmt whiteSpace? EQ whiteSpace? valueStmt;

lsetStmt : LSET whiteSpace valueStmt whiteSpace? EQ whiteSpace? valueStmt;

midStmt : MID whiteSpace? LPAREN whiteSpace? argsCall whiteSpace? RPAREN;

onErrorStmt : (ON_ERROR | ON_LOCAL_ERROR) whiteSpace (GOTO whiteSpace valueStmt | RESUME whiteSpace NEXT);

onGoToStmt : ON whiteSpace valueStmt whiteSpace GOTO whiteSpace valueStmt (whiteSpace? COMMA whiteSpace? valueStmt)*;

onGoSubStmt : ON whiteSpace valueStmt whiteSpace GOSUB whiteSpace valueStmt (whiteSpace? COMMA whiteSpace? valueStmt)*;

propertyGetStmt : 
	(visibility whiteSpace)? (STATIC whiteSpace)? PROPERTY_GET whiteSpace functionName typeHint? (whiteSpace? argList)? (whiteSpace asTypeClause)? endOfStatement 
	block? 
	END_PROPERTY
;

propertySetStmt : 
	(visibility whiteSpace)? (STATIC whiteSpace)? PROPERTY_SET whiteSpace subroutineName (whiteSpace? argList)? endOfStatement 
	block? 
	END_PROPERTY
;

propertyLetStmt : 
	(visibility whiteSpace)? (STATIC whiteSpace)? PROPERTY_LET whiteSpace subroutineName (whiteSpace? argList)? endOfStatement 
	block? 
	END_PROPERTY
;

raiseEventStmt : RAISEEVENT whiteSpace identifier (whiteSpace? LPAREN whiteSpace? (argsCall whiteSpace?)? RPAREN)?;

redimStmt : REDIM whiteSpace (PRESERVE whiteSpace)? redimSubStmt (whiteSpace? COMMA whiteSpace? redimSubStmt)*;

redimSubStmt : implicitCallStmt_InStmt whiteSpace? LPAREN whiteSpace? subscripts whiteSpace? RPAREN (whiteSpace asTypeClause)?;

resumeStmt : RESUME (whiteSpace (NEXT | valueStmt))?;

returnStmt : RETURN;

rsetStmt : RSET whiteSpace valueStmt whiteSpace? EQ whiteSpace? valueStmt;

// 5.4.2.11 Stop Statement
stopStmt : STOP;

selectCaseStmt : 
	SELECT whiteSpace CASE whiteSpace valueStmt endOfStatement 
	sC_Case*
	END_SELECT
;

sC_Selection :
    IS whiteSpace? comparisonOperator whiteSpace? valueStmt                       # caseCondIs
    | valueStmt whiteSpace TO whiteSpace valueStmt                                # caseCondTo
    | valueStmt                                                                   # caseCondValue
;

sC_Case : 
	CASE whiteSpace sC_Cond endOfStatement
	block?
;

sC_Cond :
    ELSE                                                                           # caseCondElse
    | sC_Selection (whiteSpace? COMMA whiteSpace? sC_Selection)*                   # caseCondSelection
;

setStmt : SET whiteSpace valueStmt whiteSpace? EQ whiteSpace? valueStmt;

subStmt : 
	(visibility whiteSpace)? (STATIC whiteSpace)? SUB whiteSpace? subroutineName (whiteSpace? argList)? endOfStatement
	block? 
	END_SUB
;
subroutineName : identifier;

typeStmt : 
	(visibility whiteSpace)? TYPE whiteSpace identifier endOfStatement
	typeStmt_Element*
	END_TYPE
;

typeStmt_Element : identifier (whiteSpace? LPAREN (whiteSpace? subscripts)? whiteSpace? RPAREN)? (whiteSpace asTypeClause)? endOfStatement;

valueStmt : 
	literal                                                                                         # vsLiteral
	| implicitCallStmt_InStmt                                                                       # vsICS
	| LPAREN whiteSpace? valueStmt whiteSpace? RPAREN                                               # vsStruct
	| NEW whiteSpace? valueStmt                                                                     # vsNew
	| typeOfIsExpression                                                                            # vsTypeOf
	| midStmt                                                                                       # vsMid
	| ADDRESSOF whiteSpace? valueStmt                                                               # vsAddressOf
	| unrestrictedIdentifier whiteSpace? ASSIGN whiteSpace? valueStmt                              # vsAssign
	| valueStmt whiteSpace? POW whiteSpace? valueStmt                                               # vsPow
	| MINUS whiteSpace? valueStmt                                                                   # vsNegation
	| valueStmt whiteSpace? (MULT | DIV) whiteSpace? valueStmt                                      # vsMult
	| valueStmt whiteSpace? INTDIV whiteSpace? valueStmt                                            # vsIntDiv
	| valueStmt whiteSpace? MOD whiteSpace? valueStmt                                               # vsMod
	| valueStmt whiteSpace? (PLUS | MINUS) whiteSpace? valueStmt                                    # vsAdd
	| valueStmt whiteSpace? AMPERSAND whiteSpace? valueStmt                                         # vsAmp
	| valueStmt whiteSpace? (EQ | NEQ | LT | GT | LEQ | GEQ | LIKE | IS) whiteSpace? valueStmt      # vsRelational
	| NOT whiteSpace? valueStmt                                                                     # vsNot
	| valueStmt whiteSpace? AND whiteSpace? valueStmt                                               # vsAnd
	| valueStmt whiteSpace? OR whiteSpace? valueStmt                                                # vsOr
	| valueStmt whiteSpace? XOR whiteSpace? valueStmt                                               # vsXor
	| valueStmt whiteSpace? EQV whiteSpace? valueStmt                                               # vsEqv
	| valueStmt whiteSpace? IMP whiteSpace? valueStmt                                               # vsImp
    // Added so that functions such as the Input Function, which takes a file number as argument, is supported.
    | markedFileNumber                                                                              # vsMarkedFileNumber
;

typeOfIsExpression : TYPEOF whiteSpace valueStmt (whiteSpace IS whiteSpace type)?;

variableStmt : (DIM | STATIC | visibility) whiteSpace (WITHEVENTS whiteSpace)? variableListStmt;

variableListStmt : variableSubStmt (whiteSpace? COMMA whiteSpace? variableSubStmt)*;

variableSubStmt : identifier typeHint? (whiteSpace? LPAREN whiteSpace? (subscripts whiteSpace?)? RPAREN whiteSpace?)? (whiteSpace asTypeClause)?;

whileWendStmt : 
	WHILE whiteSpace valueStmt endOfStatement 
	block?
	WEND
;

withStmt :
	WITH whiteSpace withStmtExpression endOfStatement 
	block? 
	END_WITH
;

// Special forms with special syntax, only available in a report.
circleSpecialForm : (valueStmt whiteSpace? DOT whiteSpace?)? CIRCLE whiteSpace (STEP whiteSpace?)? tuple (whiteSpace? COMMA whiteSpace? valueStmt)+;
scaleSpecialForm : (valueStmt whiteSpace? DOT whiteSpace?)? SCALE whiteSpace tuple whiteSpace? MINUS whiteSpace? tuple;
tuple : LPAREN whiteSpace? valueStmt whiteSpace? COMMA whiteSpace? valueStmt whiteSpace? RPAREN;

withStmtExpression : valueStmt;

explicitCallStmt : CALL whiteSpace explicitCallStmtExpression;

explicitCallStmtExpression : 
    implicitCallStmt_InStmt? DOT identifier typeHint? (whiteSpace? LPAREN whiteSpace? argsCall whiteSpace? RPAREN)? (whiteSpace? LPAREN subscripts RPAREN)*   # ECS_MemberCall
    | identifier typeHint? (whiteSpace? LPAREN whiteSpace? argsCall whiteSpace? RPAREN)? (whiteSpace? LPAREN subscripts RPAREN)*                              # ECS_ProcedureCall
;

implicitCallStmt_InBlock :
	iCS_B_MemberProcedureCall 
	| iCS_B_ProcedureCall
;

iCS_B_MemberProcedureCall : implicitCallStmt_InStmt? whiteSpace? DOT whiteSpace? unrestrictedIdentifier typeHint? (whiteSpace argsCall)? (whiteSpace? dictionaryCallStmt)? (whiteSpace? LPAREN subscripts RPAREN)*;

iCS_B_ProcedureCall : identifier (whiteSpace argsCall)? (whiteSpace? LPAREN subscripts RPAREN)*;

implicitCallStmt_InStmt :
	iCS_S_MembersCall
	| iCS_S_VariableOrProcedureCall
	| iCS_S_ProcedureOrArrayCall
	| iCS_S_DictionaryCall
;

iCS_S_VariableOrProcedureCall : identifier typeHint? (whiteSpace? dictionaryCallStmt)? (whiteSpace? LPAREN subscripts RPAREN)*;
iCS_S_ProcedureOrArrayCall : (identifier | baseType) typeHint? whiteSpace? LPAREN whiteSpace? (argsCall whiteSpace?)? RPAREN (whiteSpace? dictionaryCallStmt)? (whiteSpace? LPAREN subscripts RPAREN)*;

iCS_S_VariableOrProcedureCallUnrestricted : unrestrictedIdentifier typeHint? (whiteSpace? dictionaryCallStmt)? (whiteSpace? LPAREN subscripts RPAREN)*;
iCS_S_ProcedureOrArrayCallUnrestricted : (unrestrictedIdentifier | baseType) typeHint? whiteSpace? LPAREN whiteSpace? (argsCall whiteSpace?)? RPAREN (whiteSpace? dictionaryCallStmt)? (whiteSpace? LPAREN subscripts RPAREN)*;

iCS_S_MembersCall : (iCS_S_VariableOrProcedureCall | iCS_S_ProcedureOrArrayCall)? (iCS_S_MemberCall whiteSpace?)+ (whiteSpace? dictionaryCallStmt)? (whiteSpace? LPAREN subscripts RPAREN)*;

iCS_S_MemberCall : (DOT | EXCLAMATIONPOINT) whiteSpace? (iCS_S_VariableOrProcedureCallUnrestricted | iCS_S_ProcedureOrArrayCallUnrestricted);

iCS_S_DictionaryCall : whiteSpace? dictionaryCallStmt;

argsCall : (argCall? whiteSpace? (COMMA | SEMICOLON) whiteSpace?)* argCall (whiteSpace? (COMMA | SEMICOLON) whiteSpace? argCall?)*;

argCall : LPAREN? ((BYVAL | BYREF | PARAMARRAY) whiteSpace)? RPAREN? valueStmt;

dictionaryCallStmt : EXCLAMATIONPOINT whiteSpace? unrestrictedIdentifier typeHint?;

argList : LPAREN (whiteSpace? arg (whiteSpace? COMMA whiteSpace? arg)*)? whiteSpace? RPAREN;

arg : (OPTIONAL whiteSpace)? ((BYVAL | BYREF) whiteSpace)? (PARAMARRAY whiteSpace)? unrestrictedIdentifier typeHint? (whiteSpace? LPAREN whiteSpace? RPAREN)? (whiteSpace? asTypeClause)? (whiteSpace? argDefaultValue)?;

argDefaultValue : EQ whiteSpace? valueStmt;

subscripts : subscript (whiteSpace? COMMA whiteSpace? subscript)*;

subscript : (valueStmt whiteSpace TO whiteSpace)? valueStmt;

unrestrictedIdentifier : identifier | statementKeyword | markerKeyword;

identifier : IDENTIFIER | keyword;

asTypeClause : AS whiteSpace? (NEW whiteSpace)? type (whiteSpace? fieldLength)?;

baseType : BOOLEAN | BYTE | CURRENCY | DATE | DOUBLE | INTEGER | LONG | LONGLONG | LONGPTR | SINGLE | STRING | VARIANT;

comparisonOperator : LT | LEQ | GT | GEQ | EQ | NEQ | IS | LIKE;

complexType : identifier ((DOT | EXCLAMATIONPOINT) identifier)*;

fieldLength : MULT whiteSpace? (numberLiteral | identifier);

statementLabelDefinition : statementLabel whiteSpace? COLON;
statementLabel : identifierStatementLabel | lineNumberLabel;
identifierStatementLabel : unrestrictedIdentifier;
lineNumberLabel : numberLiteral;

literal : numberLiteral | DATELITERAL | STRINGLITERAL | TRUE | FALSE | NOTHING | NULL | EMPTY;

numberLiteral : HEXLITERAL | OCTLITERAL | FLOATLITERAL | INTEGERLITERAL;

type : (baseType | complexType) (whiteSpace? LPAREN whiteSpace? RPAREN)?;

typeHint : PERCENT | AMPERSAND | POW | EXCLAMATIONPOINT | HASH | AT | DOLLAR;

visibility : PRIVATE | PUBLIC | FRIEND | GLOBAL;

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
     | COLLECTION
     | CSNG
     | CSTR
     | CURRENCY
     | CVAR
     | CVERR
     | DATABASE
     | DATE
     | DEBUG
     | DELETESETTING
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
     | OPTIONAL
     | OR
     | PARAMARRAY
     | PRESERVE
     | PSET
     | REM
     | RMDIR
     | SENDKEYS
     | SETATTR
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
     | EXIT_DO 
     | EXIT_FOR 
     | EXIT_FUNCTION 
     | EXIT_PROPERTY 
     | EXIT_SUB
     | END_SELECT
     | END_WITH
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
    | PRINT
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
    whiteSpace? commentOrAnnotation?
;

endOfStatement :
    (((endOfLine NEWLINE whiteSpace?)|(whiteSpace? COLON whiteSpace?)))*
    | endOfLine EOF
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
annotationArg : valueStmt;

whiteSpace : (WS | LINE_CONTINUATION)+;