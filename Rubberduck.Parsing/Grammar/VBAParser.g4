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
	lineLabel
	| attributeStmt
	| closeStmt
	| constStmt
	| deftypeStmt
	| doLoopStmt
	| eraseStmt
	| errorStmt
    | exitStmt
	| explicitCallStmt
	| forEachStmt
	| forNextStmt
	| getStmt
	| goSubStmt
	| goToStmt
	| ifThenElseStmt
	| implementsStmt
	| inputStmt
	| letStmt
	| lineInputStmt
	| lockStmt
	| lsetStmt
	| midStmt
	| onErrorStmt
	| onGoToStmt
	| onGoSubStmt
	| openStmt
	| printStmt
	| putStmt
	| raiseEventStmt
	| redimStmt
	| resetStmt
	| resumeStmt
	| returnStmt
	| rsetStmt
	| seekStmt
	| selectCaseStmt
	| setStmt
    | stopStmt
	| unlockStmt
	| variableStmt
	| whileWendStmt
	| widthStmt
	| withStmt
	| writeStmt
	| implicitCallStmt_InBlock
;

closeStmt : CLOSE (whiteSpace fileNumber (whiteSpace? COMMA whiteSpace? fileNumber)*)?;

constStmt : (visibility whiteSpace)? CONST whiteSpace constSubStmt (whiteSpace? COMMA whiteSpace? constSubStmt)*;

constSubStmt : identifier typeHint? (whiteSpace asTypeClause)? whiteSpace? EQ whiteSpace? valueStmt;

declareStmt : (visibility whiteSpace)? DECLARE whiteSpace (PTRSAFE whiteSpace)? ((FUNCTION typeHint?) | SUB) whiteSpace identifier typeHint? whiteSpace LIB whiteSpace STRINGLITERAL (whiteSpace ALIAS whiteSpace STRINGLITERAL)? (whiteSpace? argList)? (whiteSpace asTypeClause)?;

deftypeStmt : 
	(
		DEFBOOL | DEFBYTE | DEFINT | DEFLNG | DEFLNGLNG | DEFLNGPTR | DEFCUR |
		DEFSNG | DEFDBL | DEFDATE | 
		DEFSTR | DEFOBJ | DEFVAR
	) whiteSpace
	letterrange (whiteSpace? COMMA whiteSpace? letterrange)*
;

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

getStmt : GET whiteSpace fileNumber whiteSpace? COMMA whiteSpace? valueStmt? whiteSpace? COMMA whiteSpace? valueStmt;

goSubStmt : GOSUB whiteSpace valueStmt;

goToStmt : GOTO whiteSpace valueStmt;

ifThenElseStmt : 
	IF whiteSpace ifConditionStmt whiteSpace THEN whiteSpace blockStmt (whiteSpace ELSE whiteSpace blockStmt)?	# inlineIfThenElse
	| ifBlockStmt ifElseIfBlockStmt* ifElseBlockStmt? END_IF			# blockIfThenElse
;

ifBlockStmt : 
	IF whiteSpace ifConditionStmt whiteSpace THEN endOfStatement 
	block?
;

ifConditionStmt : valueStmt;

ifElseIfBlockStmt : 
	ELSEIF whiteSpace ifConditionStmt whiteSpace THEN endOfStatement
	block?
;

ifElseBlockStmt : 
	ELSE endOfStatement 
	block?
;

implementsStmt : IMPLEMENTS whiteSpace valueStmt;

inputStmt : INPUT whiteSpace fileNumber (whiteSpace? COMMA whiteSpace? valueStmt)+;

letStmt : (LET whiteSpace)? valueStmt whiteSpace? EQ whiteSpace? valueStmt;

lineInputStmt : LINE_INPUT whiteSpace fileNumber whiteSpace? COMMA whiteSpace? valueStmt;

lockStmt : LOCK whiteSpace valueStmt (whiteSpace? COMMA whiteSpace? valueStmt (whiteSpace TO whiteSpace valueStmt)?)?;

lsetStmt : LSET whiteSpace valueStmt whiteSpace? EQ whiteSpace? valueStmt;

midStmt : MID whiteSpace? LPAREN whiteSpace? argsCall whiteSpace? RPAREN;

onErrorStmt : (ON_ERROR | ON_LOCAL_ERROR) whiteSpace (GOTO whiteSpace valueStmt | RESUME whiteSpace NEXT);

// TODO: only first valueStmt is correct, rest should be IDENTIFIER/INTEGERs?
onGoToStmt : ON whiteSpace valueStmt whiteSpace GOTO whiteSpace valueStmt (whiteSpace? COMMA whiteSpace? valueStmt)*;

onGoSubStmt : ON whiteSpace valueStmt whiteSpace GOSUB whiteSpace valueStmt (whiteSpace? COMMA whiteSpace? valueStmt)*;

openStmt : 
	OPEN whiteSpace valueStmt whiteSpace FOR whiteSpace (APPEND | BINARY | INPUT | OUTPUT | RANDOM) 
	(whiteSpace ACCESS whiteSpace (READ | WRITE | READ_WRITE))?
	(whiteSpace (SHARED | LOCK_READ | LOCK_WRITE | LOCK_READ_WRITE))?
	whiteSpace AS whiteSpace fileNumber
	(whiteSpace LEN whiteSpace? EQ whiteSpace? valueStmt)?
;

outputList :
	outputList_Expression (whiteSpace? (SEMICOLON | COMMA) whiteSpace? outputList_Expression?)*
	| outputList_Expression? (whiteSpace? (SEMICOLON | COMMA) whiteSpace? outputList_Expression?)+
;

outputList_Expression : 
	valueStmt
	| (SPC | TAB) (whiteSpace? LPAREN whiteSpace? argsCall whiteSpace? RPAREN)?
;

printStmt : PRINT whiteSpace fileNumber whiteSpace? COMMA (whiteSpace? outputList)?;

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

putStmt : PUT whiteSpace fileNumber whiteSpace? COMMA whiteSpace? valueStmt? whiteSpace? COMMA whiteSpace? valueStmt;

raiseEventStmt : RAISEEVENT whiteSpace identifier (whiteSpace? LPAREN whiteSpace? (argsCall whiteSpace?)? RPAREN)?;

redimStmt : REDIM whiteSpace (PRESERVE whiteSpace)? redimSubStmt (whiteSpace? COMMA whiteSpace? redimSubStmt)*;

redimSubStmt : implicitCallStmt_InStmt whiteSpace? LPAREN whiteSpace? subscripts whiteSpace? RPAREN (whiteSpace asTypeClause)?;

resetStmt : RESET;

resumeStmt : RESUME (whiteSpace (NEXT | valueStmt))?;

returnStmt : RETURN;

rsetStmt : RSET whiteSpace valueStmt whiteSpace? EQ whiteSpace? valueStmt;

// 5.4.2.11 Stop Statement
stopStmt : STOP;

seekStmt : SEEK whiteSpace fileNumber whiteSpace? COMMA whiteSpace? valueStmt;

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

unlockStmt : UNLOCK whiteSpace fileNumber (whiteSpace? COMMA whiteSpace? valueStmt (whiteSpace TO whiteSpace valueStmt)?)?;

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
;

typeOfIsExpression : TYPEOF whiteSpace valueStmt (whiteSpace IS whiteSpace type)?;

variableStmt : (DIM | STATIC | visibility) whiteSpace (WITHEVENTS whiteSpace)? variableListStmt;

variableListStmt : variableSubStmt (whiteSpace? COMMA whiteSpace? variableSubStmt)*;

variableSubStmt : identifier (whiteSpace? LPAREN whiteSpace? (subscripts whiteSpace?)? RPAREN whiteSpace?)? typeHint? (whiteSpace asTypeClause)?;

whileWendStmt : 
	WHILE whiteSpace valueStmt endOfStatement 
	block?
	WEND
;

widthStmt : WIDTH whiteSpace fileNumber whiteSpace? COMMA whiteSpace? valueStmt;

withStmt :
	WITH whiteSpace withStmtExpression endOfStatement 
	block? 
	END_WITH
;

withStmtExpression : valueStmt;

writeStmt : WRITE whiteSpace fileNumber whiteSpace? COMMA (whiteSpace? outputList)?;

fileNumber : HASH? valueStmt;

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

letterrange : identifier (whiteSpace? MINUS whiteSpace? identifier)?;

lineLabel : (identifier | numberLiteral) COLON;

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
     | CIRCLE
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
     | END_IF
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
     | SCALE
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
;

markerKeyword : AS;

statementKeyword :
    CALL
    | CASE
    | CLOSE
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
    | END
    | ENDIF
    | ENUM
    | ERASE
    | EVENT
    | EXIT
    | FOR
    | FRIEND
    | FUNCTION
    | GET
    | GLOBAL
    | GOSUB
    | GOTO
    | IF
    | IMPLEMENTS
    | INPUT
    | LET
    | LOCK
    | LOOP
    | LSET
    | NEXT
    | ON
    | OPEN
    | OPTION
    | PRINT
    | PRIVATE
    | PUBLIC
    | PUT
    | RAISEEVENT
    | REDIM
    | RESUME
    | RETURN
    | RSET
    | SEEK
    | SELECT
    | SET
    | STATIC
    | STOP
    | SUB
    | TYPE
    | UNLOCK
    | WEND
    | WHILE
    | WITH
    | WRITE
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
    | RESET
    | LINE_INPUT
    | WIDTH
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
	 | whiteSpace? LPAREN whiteSpace? annotationArg whiteSpace? RPAREN
	 | whiteSpace? LPAREN annotationArg (whiteSpace? COMMA whiteSpace? annotationArg)+ whiteSpace? RPAREN;
annotationArg : valueStmt;

whiteSpace : (WS | LINE_CONTINUATION)+;