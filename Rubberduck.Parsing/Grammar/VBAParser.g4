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

startRule : module EOF;

module : 
	WS?
	endOfStatement
	(moduleHeader endOfStatement)?
	moduleConfig? endOfStatement
	moduleAttributes? endOfStatement
	moduleDeclarations? endOfStatement
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

moduleDeclarations : moduleDeclarationsElement (endOfStatement moduleDeclarationsElement)* endOfStatement;

moduleOption : 
	OPTION_BASE WS SHORTLITERAL 					# optionBaseStmt
	| OPTION_COMPARE WS (BINARY | TEXT | DATABASE) 	# optionCompareStmt
	| OPTION_EXPLICIT 								# optionExplicitStmt
	| OPTION_PRIVATE_MODULE 						# optionPrivateModuleStmt
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

attributeStmt : ATTRIBUTE WS implicitCallStmt_InStmt WS? EQ WS? literal (WS? COMMA WS? literal)*;

block : blockStmt (endOfStatement blockStmt)* endOfStatement;

blockStmt :
	lineLabel
    | appactivateStmt
	| attributeStmt
	| beepStmt
	| chdirStmt
	| chdriveStmt
	| closeStmt
	| constStmt
	| dateStmt
	| deleteSettingStmt
	| deftypeStmt
	| doLoopStmt
	| endStmt
	| eraseStmt
	| errorStmt
	| exitStmt
	| explicitCallStmt
	| filecopyStmt
	| forEachStmt
	| forNextStmt
	| getStmt
	| goSubStmt
	| goToStmt
	| ifThenElseStmt
	| implementsStmt
	| inputStmt
	| killStmt
	| letStmt
	| lineInputStmt
	| loadStmt
	| lockStmt
	| lsetStmt
	| midStmt
	| mkdirStmt
	| nameStmt
	| onErrorStmt
	| onGoToStmt
	| onGoSubStmt
	| openStmt
	| printStmt
	| putStmt
	| raiseEventStmt
	| randomizeStmt
	| redimStmt
	| resetStmt
	| resumeStmt
	| returnStmt
	| rmdirStmt
	| rsetStmt
	| savepictureStmt
	| saveSettingStmt
	| seekStmt
	| selectCaseStmt
	| sendkeysStmt
	| setattrStmt
	| setStmt
	| stopStmt
	| timeStmt
	| unloadStmt
	| unlockStmt
	| variableStmt
	| whileWendStmt
	| widthStmt
	| withStmt
	| writeStmt
	| implicitCallStmt_InBlock
;

appactivateStmt : APPACTIVATE WS valueStmt (WS? COMMA WS? valueStmt)?;

beepStmt : BEEP;

chdirStmt : CHDIR WS valueStmt;

chdriveStmt : CHDRIVE WS valueStmt;

closeStmt : CLOSE (WS fileNumber (WS? COMMA WS? fileNumber)*)?;

constStmt : (visibility WS)? CONST WS constSubStmt (WS? COMMA WS? constSubStmt)*;

constSubStmt : ambiguousIdentifier typeHint? (WS asTypeClause)? WS? EQ WS? valueStmt;

dateStmt : DATE WS? EQ WS? valueStmt;

declareStmt : (visibility WS)? DECLARE WS (PTRSAFE WS)? ((FUNCTION typeHint?) | SUB) WS ambiguousIdentifier typeHint? WS LIB WS STRINGLITERAL (WS ALIAS WS STRINGLITERAL)? (WS? argList)? (WS asTypeClause)?;

deftypeStmt : 
	(
		DEFBOOL | DEFBYTE | DEFINT | DEFLNG | DEFCUR | 
		DEFSNG | DEFDBL | DEFDEC | DEFDATE | 
		DEFSTR | DEFOBJ | DEFVAR
	) WS
	letterrange (WS? COMMA WS? letterrange)*
;

deleteSettingStmt : DELETESETTING WS valueStmt WS? COMMA WS? valueStmt (WS? COMMA WS? valueStmt)?;

doLoopStmt :
	DO endOfStatement 
	block?
	LOOP
	|
	DO WS (WHILE | UNTIL) WS valueStmt endOfStatement
	block?
	LOOP
	| 
	DO endOfStatement
	block
	LOOP WS (WHILE | UNTIL) WS valueStmt
;

endStmt : END;

enumerationStmt: 
	(visibility WS)? ENUM WS ambiguousIdentifier endOfStatement 
	enumerationStmt_Constant* 
	END_ENUM
;

enumerationStmt_Constant : ambiguousIdentifier (WS? EQ WS? valueStmt)? endOfStatement;

eraseStmt : ERASE WS valueStmt;

errorStmt : ERROR WS valueStmt;

eventStmt : (visibility WS)? EVENT WS ambiguousIdentifier WS? argList;

exitStmt : EXIT_DO | EXIT_FOR | EXIT_FUNCTION | EXIT_PROPERTY | EXIT_SUB;

filecopyStmt : FILECOPY WS valueStmt WS? COMMA WS? valueStmt;

forEachStmt : 
	FOR WS EACH WS ambiguousIdentifier typeHint? WS IN WS valueStmt endOfStatement
	block?
	NEXT (WS ambiguousIdentifier)?
;

forNextStmt : 
	FOR WS ambiguousIdentifier typeHint? (WS asTypeClause)? WS? EQ WS? valueStmt WS TO WS valueStmt (WS STEP WS valueStmt)? endOfStatement 
	block?
	NEXT (WS ambiguousIdentifier)?
; 

functionStmt :
	(visibility WS)? (STATIC WS)? FUNCTION WS? ambiguousIdentifier typeHint? (WS? argList)? (WS? asTypeClause)? endOfStatement
	block?
	END_FUNCTION
;

getStmt : GET WS fileNumber WS? COMMA WS? valueStmt? WS? COMMA WS? valueStmt;

goSubStmt : GOSUB WS valueStmt;

goToStmt : GOTO WS valueStmt;

ifThenElseStmt : 
	IF WS ifConditionStmt WS THEN WS blockStmt (WS ELSE WS blockStmt)?	# inlineIfThenElse
	| ifBlockStmt ifElseIfBlockStmt* ifElseBlockStmt? END_IF			# blockIfThenElse
;

ifBlockStmt : 
	IF WS ifConditionStmt WS THEN endOfStatement 
	block?
;

ifConditionStmt : valueStmt;

ifElseIfBlockStmt : 
	ELSEIF WS ifConditionStmt WS THEN endOfStatement
	block?
;

ifElseBlockStmt : 
	ELSE endOfStatement 
	block?
;

implementsStmt : IMPLEMENTS WS ambiguousIdentifier;

inputStmt : INPUT WS fileNumber (WS? COMMA WS? valueStmt)+;

killStmt : KILL WS valueStmt;

letStmt : (LET WS)? implicitCallStmt_InStmt WS? EQ WS? valueStmt;

lineInputStmt : LINE_INPUT WS fileNumber WS? COMMA WS? valueStmt;

loadStmt : LOAD WS valueStmt;

lockStmt : LOCK WS valueStmt (WS? COMMA WS? valueStmt (WS TO WS valueStmt)?)?;

lsetStmt : LSET WS implicitCallStmt_InStmt WS? EQ WS? valueStmt;

midStmt : MID WS? LPAREN WS? argsCall WS? RPAREN;

mkdirStmt : MKDIR WS valueStmt;

nameStmt : NAME WS valueStmt WS AS WS valueStmt;

onErrorStmt : (ON_ERROR | ON_LOCAL_ERROR) WS (GOTO WS valueStmt | RESUME WS NEXT);

onGoToStmt : ON WS valueStmt WS GOTO WS valueStmt (WS? COMMA WS? valueStmt)*;

onGoSubStmt : ON WS valueStmt WS GOSUB WS valueStmt (WS? COMMA WS? valueStmt)*;

openStmt : 
	OPEN WS valueStmt WS FOR WS (APPEND | BINARY | INPUT | OUTPUT | RANDOM) 
	(WS ACCESS WS (READ | WRITE | READ_WRITE))?
	(WS (SHARED | LOCK_READ | LOCK_WRITE | LOCK_READ_WRITE))?
	WS AS WS fileNumber
	(WS LEN WS? EQ WS? valueStmt)?
;

outputList :
	outputList_Expression (WS? (SEMICOLON | COMMA) WS? outputList_Expression?)*
	| outputList_Expression? (WS? (SEMICOLON | COMMA) WS? outputList_Expression?)+
;

outputList_Expression : 
	valueStmt
	| (SPC | TAB) (WS? LPAREN WS? argsCall WS? RPAREN)?
;

printStmt : PRINT WS fileNumber WS? COMMA (WS? outputList)?;

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

putStmt : PUT WS fileNumber WS? COMMA WS? valueStmt? WS? COMMA WS? valueStmt;

raiseEventStmt : RAISEEVENT WS ambiguousIdentifier (WS? LPAREN WS? (argsCall WS?)? RPAREN)?;

randomizeStmt : RANDOMIZE (WS valueStmt)?;

redimStmt : REDIM WS (PRESERVE WS)? redimSubStmt (WS? COMMA WS? redimSubStmt)*;

redimSubStmt : implicitCallStmt_InStmt WS? LPAREN WS? subscripts WS? RPAREN (WS asTypeClause)?;

resetStmt : RESET;

resumeStmt : RESUME (WS (NEXT | ambiguousIdentifier))?;

returnStmt : RETURN;

rmdirStmt : RMDIR WS valueStmt;

rsetStmt : RSET WS implicitCallStmt_InStmt WS? EQ WS? valueStmt;

savepictureStmt : SAVEPICTURE WS valueStmt WS? COMMA WS? valueStmt;

saveSettingStmt : SAVESETTING WS valueStmt WS? COMMA WS? valueStmt WS? COMMA WS? valueStmt WS? COMMA WS? valueStmt;

seekStmt : SEEK WS fileNumber WS? COMMA WS? valueStmt;

selectCaseStmt : 
	SELECT WS CASE WS valueStmt endOfStatement 
	sC_Case*
	END_SELECT
;

sC_Selection :
    IS WS? comparisonOperator WS? valueStmt                       # caseCondIs
    | valueStmt WS TO WS valueStmt                                # caseCondTo
    | valueStmt                                                   # caseCondValue
;

sC_Case : 
	CASE WS sC_Cond endOfStatement
	block?
;

sC_Cond :
    ELSE                                                            # caseCondElse
    | sC_Selection (WS? COMMA WS? sC_Selection)*                      # caseCondSelection
;

sendkeysStmt : SENDKEYS WS valueStmt (WS? COMMA WS? valueStmt)?;

setattrStmt : SETATTR WS valueStmt WS? COMMA WS? valueStmt;

setStmt : SET WS implicitCallStmt_InStmt WS? EQ WS? valueStmt;

stopStmt : STOP;

subStmt : 
	(visibility WS)? (STATIC WS)? SUB WS? ambiguousIdentifier (WS? argList)? endOfStatement
	block? 
	END_SUB
;

timeStmt : TIME WS? EQ WS? valueStmt;

typeStmt : 
	(visibility WS)? TYPE WS ambiguousIdentifier endOfStatement
	typeStmt_Element*
	END_TYPE
;

typeStmt_Element : ambiguousIdentifier (WS? LPAREN (WS? subscripts)? WS? RPAREN)? (WS asTypeClause)? endOfStatement;

typeOfStmt : TYPEOF WS valueStmt (WS IS WS type)?;

unloadStmt : UNLOAD WS valueStmt;

unlockStmt : UNLOCK WS fileNumber (WS? COMMA WS? valueStmt (WS TO WS valueStmt)?)?;

valueStmt : 
	literal                                                                         # vsLiteral
	| implicitCallStmt_InStmt                                                       # vsICS
	| LPAREN WS? valueStmt (WS? COMMA WS? valueStmt)* RPAREN                        # vsStruct
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

variableStmt : (DIM | STATIC | visibility) WS (WITHEVENTS WS)? variableListStmt;

variableListStmt : variableSubStmt (WS? COMMA WS? variableSubStmt)*;

variableSubStmt : ambiguousIdentifier (WS? LPAREN WS? (subscripts WS?)? RPAREN WS?)? typeHint? (WS asTypeClause)?;

whileWendStmt : 
	WHILE WS valueStmt endOfStatement 
	block?
	WEND
;

widthStmt : WIDTH WS fileNumber WS? COMMA WS? valueStmt;

withStmt : 
	WITH WS (implicitCallStmt_InStmt | (NEW WS type)) endOfStatement 
	block? 
	END_WITH
;

writeStmt : WRITE WS fileNumber WS? COMMA (WS? outputList)?;

fileNumber : HASH? valueStmt;

explicitCallStmt : 
	eCS_ProcedureCall 
	| eCS_MemberProcedureCall 
;

eCS_ProcedureCall : CALL WS ambiguousIdentifier typeHint? (WS? LPAREN WS? argsCall WS? RPAREN)? (WS? LPAREN subscripts RPAREN)*;

eCS_MemberProcedureCall : CALL WS implicitCallStmt_InStmt? DOT ambiguousIdentifier typeHint? (WS? LPAREN WS? argsCall WS? RPAREN)? (WS? LPAREN subscripts RPAREN)*;

implicitCallStmt_InBlock :
	iCS_B_MemberProcedureCall 
	| iCS_B_ProcedureCall
;

iCS_B_MemberProcedureCall : implicitCallStmt_InStmt? DOT ambiguousIdentifier typeHint? (WS argsCall)? (WS? dictionaryCallStmt)? (WS? LPAREN subscripts RPAREN)*;

iCS_B_ProcedureCall : certainIdentifier (WS argsCall)? (WS? LPAREN subscripts RPAREN)*;

implicitCallStmt_InStmt :
	iCS_S_MembersCall
	| iCS_S_VariableOrProcedureCall
	| iCS_S_ProcedureOrArrayCall
	| iCS_S_DictionaryCall
;

iCS_S_VariableOrProcedureCall : ambiguousIdentifier typeHint? (WS? dictionaryCallStmt)? (WS? LPAREN subscripts RPAREN)*;

iCS_S_ProcedureOrArrayCall : (ambiguousIdentifier | baseType) typeHint? WS? LPAREN WS? (argsCall WS?)? RPAREN (WS? dictionaryCallStmt)? (WS? LPAREN subscripts RPAREN)*;

iCS_S_MembersCall : (iCS_S_VariableOrProcedureCall | iCS_S_ProcedureOrArrayCall)? (iCS_S_MemberCall WS?)+ (WS? dictionaryCallStmt)? (WS? LPAREN subscripts RPAREN)*;

iCS_S_MemberCall : (DOT | EXCLAMATIONPOINT) WS? (iCS_S_VariableOrProcedureCall | iCS_S_ProcedureOrArrayCall);

iCS_S_DictionaryCall : WS? dictionaryCallStmt;

argsCall : (argCall? WS? (COMMA | SEMICOLON) WS?)* argCall (WS? (COMMA | SEMICOLON) WS? argCall?)*;

argCall : LPAREN? ((BYVAL | BYREF | PARAMARRAY) WS)? RPAREN? valueStmt;

dictionaryCallStmt : EXCLAMATIONPOINT WS? ambiguousIdentifier typeHint?;

argList : LPAREN (WS? arg (WS? COMMA WS? arg)*)? WS? RPAREN;

arg : (OPTIONAL WS)? ((BYVAL | BYREF) WS)? (PARAMARRAY WS)? ambiguousIdentifier typeHint? (WS? LPAREN WS? RPAREN)? (WS? asTypeClause)? (WS? argDefaultValue)?;

argDefaultValue : EQ WS? valueStmt;

subscripts : subscript (WS? COMMA WS? subscript)*;

subscript : (valueStmt WS TO WS)? valueStmt;

ambiguousIdentifier : 
	(IDENTIFIER | ambiguousKeyword)+
;

asTypeClause : AS WS? (NEW WS)? type (WS? fieldLength)?;

baseType : BOOLEAN | BYTE | COLLECTION | DATE | DOUBLE | INTEGER | LONG | SINGLE | STRING | VARIANT;

certainIdentifier : 
	IDENTIFIER (ambiguousKeyword | IDENTIFIER)*
	| ambiguousKeyword (ambiguousKeyword | IDENTIFIER)+
;

comparisonOperator : LT | LEQ | GT | GEQ | EQ | NEQ | IS | LIKE;

complexType : ambiguousIdentifier ((DOT | EXCLAMATIONPOINT) ambiguousIdentifier)*;

fieldLength : MULT WS? (INTEGERLITERAL | ambiguousIdentifier);

letterrange : certainIdentifier (WS? MINUS WS? certainIdentifier)?;

lineLabel : ambiguousIdentifier COLON;

literal : HEXLITERAL | OCTLITERAL | DATELITERAL | DOUBLELITERAL | INTEGERLITERAL | SHORTLITERAL | STRINGLITERAL | TRUE | FALSE | NOTHING | NULL | EMPTY;

type : (baseType | complexType) (WS? LPAREN WS? RPAREN)?;

typeHint : PERCENT | AMPERSAND | EXP | EXCLAMATIONPOINT | HASH | AT | DOLLAR;

visibility : PRIVATE | PUBLIC | FRIEND | GLOBAL;

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