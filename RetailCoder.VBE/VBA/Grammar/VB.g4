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

/*
* Visual Basic 6.0 Grammar for ANTLR4, Version 1.0 
*
* This is an approximate grammar for Visual Basic 6.0, derived 
* from the Visual Basic 6.0 language reference 
* http://msdn.microsoft.com/en-us/library/aa338033%28v=vs.60%29.aspx 
* and tested against MSDN VB6 statement examples as well as several Visual 
* Basic 6.0 code repositories.
*
* Characteristics:
*
* 1. This grammar is line-based and takes into account whitespace, so that
*    member calls (e.g. "A.B") are distinguished from contextual object calls 
*    in WITH statements (e.g. "A .B").
*
* 2. Keywords can be used as identifiers depending on the context, enabling
*    e.g. "A.Type", but not "Type.B".
*
*
* Known limitations:
*
* 1. Preprocessor statements (#if, #else, ...) must not interfere with regular
*    statements.
*
* 2. Comments are skipped.
*
*
* Change log:
*
* v1.0 Initial revision
*/

/*
 Changes for VBA / C#:
	- Added "Option Compare Database"
	- Renamed parser rules to PascalCase
*/

grammar VB;

// module ----------------------------------

StartRule : Module EOF;

Module : 
	WS? NEWLINE*
	(ModuleHeader NEWLINE+)?
	ModuleConfig? NEWLINE*
	ModuleAttributes? NEWLINE*
	ModuleOptions? NEWLINE*
	ModuleBody? NEWLINE*
	WS?
;

ModuleHeader : VERSION WS DOUBLELITERAL WS CLASS;

ModuleConfig :
	BEGIN NEWLINE+ 
	(AmbiguousIdentifier WS? EQ WS? Aiteral NEWLINE)+ 
	END NEWLINE+
;

ModuleAttributes : (AttributeStmt NEWLINE+)+;

ModuleOptions : (ModuleOption NEWLINE+)+;

ModuleOption : 
	OPTION_BASE WS INTEGERLITERAL # OptionBaseStmt
	| OPTION_COMPARE WS (BINARY | TEXT | DATABASE) # OptionCompareStmt
	| OPTION_EXPLICIT # OptionExplicitStmt
	| OPTION_PRIVATE_MODULE # OptionPrivateModuleStmt
;

ModuleBody : 
	ModuleBodyElement (NEWLINE+ ModuleBodyElement)*
;

ModuleBodyElement : 
	ModuleBlock
	| DeclareStmt
	| EnumerationStmt 
	| EventStmt
	| FunctionStmt 
	| MacroIfThenElseStmt
	| PropertyGetStmt 
	| PropertySetStmt 
	| PropertyLetStmt 
	| SubStmt 
	| TypeStmt
;


// Block ----------------------------------

ModuleBlock : Block;

AttributeStmt : ATTRIBUTE WS ImplicitCallStmt_InStmt WS? EQ WS? Literal (WS? ',' WS? Literal)*;

bBock : BlockStmt (NEWLINE+ WS? BlockStmt)*;

BlockStmt : 
	AppActivateStmt
	| AttributeStmt
	| BeepStmt
	| ChdirStmt
	| ChdriveStmt
	| CloseStmt
	| ConstStmt
	| DateStmt
	| DeleteSettingStmt
	| DefTypeStmt
	| DoLoopStmt
	| EndStmt
	| EraseStmt
	| ErrorStmt
	| ExitStmt
	| ExplicitCallStmt
	| FileCopyStmt
	| ForEachStmt
	| ForNextStmt
	| GetStmt
	| GoSubStmt
	| GoToStmt
	| IfThenElseStmt
	| ImplementsStmt
	| ImplicitCallStmt_InBlock
	| InputStmt
	| KillStmt
	| LetStmt
	| LineInputStmt
	| LineLabel
	| LoadStmt
	| LockStmt
	| LSetStmt
	| MacroIfThenElseStmt
	| MidStmt
	| MkdirStmt
	| NameStmt
	| OnErrorStmt
	| OnGoToStmt
	| OnGoSubStmt
	| OpenStmt
	| PrintStmt
	| PutStmt
	| RaiseEventStmt
	| RandomizeStmt
	| RedimStmt
	| ResetStmt
	| ResumeStmt
	| ReturnStmt
	| RmdirStmt
	| RSetStmt
	| SavePictureStmt
	| SaveSettingStmt
	| SeekStmt
	| SelectCaseStmt
	| SendkeysStmt
	| SetAttrStmt
	| SetStmt
	| StopStmt
	| TimeStmt
	| UnloadStmt
	| UnlockStmt
	| VariableStmt
	| WhileWendStmt
	| WidthStmt
	| WithStmt
	| WriteStmt
;


// statements ----------------------------------

AppActivateStmt : APPACTIVATE WS ValueStmt (WS? ',' WS? ValueStmt)?;

BeepStmt : BEEP;

ChdirStmt : CHDIR WS ValueStmt;

ChdriveStmt : CHDRIVE WS ValueStmt;

CloseStmt : CLOSE (WS ValueStmt (WS? ',' WS? ValueStmt)*)?;

ConstStmt : (Visibility WS)? CONST WS ConstSubStmt (WS? ',' WS? ConstSubStmt)*;

ConstSubStmt : AmbiguousIdentifier TypeHint? (WS AsTypeClause)? WS? EQ WS? ValueStmt;

DateStmt : DATE WS? EQ WS? ValueStmt;

DeclareStmt : (Visibility WS)? DECLARE WS (FUNCTION | SUB) WS AmbiguousIdentifier WS LIB WS STRINGLITERAL (WS ALIAS WS STRINGLITERAL)? (WS? ArgList)? (WS AsTypeClause)?;

DefTypeStmt : 
	(
		DEFBOOL | DEFBYTE | DEFINT | DEFLNG | DEFCUR | 
		DEFSNG | DEFDBL | DEFDEC | DEFDATE | 
		DEFSTR | DEFOBJ | DEFVAR
	) WS
	LetterRange (WS? ',' WS? LetterRange)*
;

DeleteSettingStmt : DELETESETTING WS ValueStmt WS? ',' WS? ValueStmt (WS? ',' WS? ValueStmt)?;

DoLoopStmt :
	DO NEWLINE+ 
	(Block NEWLINE+)? 
	LOOP
	|
	DO WS (WHILE | UNTIL) WS ValueStmt NEWLINE+ 
	(Block NEWLINE+)? 
	LOOP
	| 
	DO NEWLINE+ 
	(Block NEWLINE+) 
	LOOP WS (WHILE | UNTIL) WS ValueStmt
;

EndStmt : END;

EnumerationStmt: 
	(Visibility WS)? ENUM WS AmbiguousIdentifier NEWLINE+ 
	(EnumerationStmt_Constant)* 
	END_ENUM
;

EnumerationStmt_Constant : AmbiguousIdentifier (WS? EQ WS? ValueStmt)? NEWLINE+;

EraseStmt : ERASE WS ValueStmt;

ErrorStmt : ERROR WS ValueStmt;

EventStmt : (Visibility WS)? EVENT WS AmbiguousIdentifier WS? ArgList;

ExitStmt : EXIT_DO | EXIT_FOR | EXIT_FUNCTION | EXIT_PROPERTY | EXIT_SUB;

FileCopyStmt : FILECOPY WS ValueStmt WS? ',' WS? ValueStmt;

ForEachStmt : 
	FOR WS EACH WS AmbiguousIdentifier TypeHint? WS IN WS ValueStmt NEWLINE+ 
	(Block NEWLINE+)?
	NEXT (WS AmbiguousIdentifier)?
;

ForNextStmt : 
	FOR WS AmbiguousIdentifier TypeHint? (WS AsTypeClause)? WS? EQ WS? ValueStmt WS TO WS ValueStmt (WS STEP WS ValueStmt)? NEWLINE+ 
	(Block NEWLINE+)? 
	NEXT (WS AmbiguousIdentifier)?
; 

FunctionStmt :
	(Visibility WS)? (STATIC WS)? FUNCTION WS AmbiguousIdentifier (WS? ArgList)? (WS AsTypeClause)? NEWLINE+
	(Block NEWLINE+)?
	END_FUNCTION
;

GetStmt : GET WS ValueStmt WS? ',' WS? ValueStmt? WS? ',' WS? ValueStmt;

GoSubStmt : GOSUB WS ValueStmt;

GoToStmt : GOTO WS ValueStmt;

IfThenElseStmt : 
	IF WS IfConditionStmt WS THEN WS BlockStmt (WS ELSE WS BlockStmt)? # inlineIfThenElse
	| IfBlockStmt IfElseIfBlockStmt* IfElseBlockStmt? END_IF # blockIfThenElse
;

IfBlockStmt : 
	IF WS IfConditionStmt WS THEN NEWLINE+ 
	(Block NEWLINE+)?
;

IfConditionStmt : ValueStmt;

IfElseIfBlockStmt : 
	ELSEIF WS IfConditionStmt WS THEN NEWLINE+ 
	(Block NEWLINE+)?
;

IfElseBlockStmt : 
	ELSE NEWLINE+ 
	(Block NEWLINE+)?
;

ImplementsStmt : IMPLEMENTS WS AmbiguousIdentifier;

InputStmt : INPUT WS ValueStmt (WS? ',' WS? ValueStmt)+;

KillStmt : KILL WS ValueStmt;

LetStmt : (LET WS)? ImplicitCallStmt_InStmt WS? EQ WS? ValueStmt;

LineInputStmt : LINE_INPUT WS ValueStmt WS? ',' WS? ValueStmt;

LoadStmt : LOAD WS ValueStmt;

LockStmt : LOCK WS ValueStmt (WS? ',' WS? ValueStmt (WS TO WS ValueStmt)?)?;

LSetStmt : LSET WS ImplicitCallStmt_InStmt WS? EQ WS? ValueStmt;

MacroIfThenElseStmt : MacroIfBlockStmt MacroElseIfBlockStmt* MacroElseBlockStmt? MACRO_END_IF;

MacroIfBlockStmt : 
	MACRO_IF WS IfConditionStmt WS THEN NEWLINE+ 
	(ModuleBody NEWLINE+)?
;

MacroElseIfBlockStmt : 
	MACRO_ELSEIF WS IfConditionStmt WS THEN NEWLINE+ 
	(ModuleBody NEWLINE+)?
;

MacroElseBlockStmt : 
	MACRO_ELSE NEWLINE+ 
	(ModuleBody NEWLINE+)?
;

MidStmt : MID WS? LPAREN WS? ArgsCall WS? RPAREN;

MkdirStmt : MKDIR WS ValueStmt;

NameStmt : NAME WS ValueStmt WS AS WS ValueStmt;

OnErrorStmt : ON_ERROR WS (GOTO WS ValueStmt | RESUME WS NEXT);

OnGoToStmt : ON WS ValueStmt WS GOTO WS ValueStmt (WS? ',' WS? ValueStmt)*;

OnGoSubStmt : ON WS ValueStmt WS GOSUB WS ValueStmt (WS? ',' WS? ValueStmt)*;

OpenStmt : 
	OPEN WS ValueStmt WS FOR WS (APPEND | BINARY | INPUT | OUTPUT | RANDOM) 
	(WS ACCESS WS (READ | WRITE | READ_WRITE))?
	(WS (SHARED | LOCK_READ | LOCK_WRITE | LOCK_READ_WRITE))?
	WS AS WS ValueStmt
	(WS LEN WS? EQ WS? ValueStmt)?
;

OutputList :
	OutputList_Expression (WS? (';' | ',') WS? OutputList_Expression?)*
	| OutputList_Expression? (WS? (';' | ',') WS? OutputList_Expression?)+
;

OutputList_Expression : 
	ValueStmt
	| (SPC | TAB) (WS? LPAREN WS? ArgsCall WS? RPAREN)?
;

PrintStmt : PRINT WS ValueStmt WS? ',' (WS? OutputList)?;

PropertyGetStmt : 
	(Visibility WS)? (STATIC WS)? PROPERTY_GET WS AmbiguousIdentifier (WS? ArgList)? (WS AsTypeClause)? NEWLINE+ 
	(Block NEWLINE+)? 
	END_PROPERTY
;

PropertySetStmt : 
	(Visibility WS)? (STATIC WS)? PROPERTY_SET WS AmbiguousIdentifier (WS? ArgList)? NEWLINE+ 
	(Block NEWLINE+)? 
	END_PROPERTY
;

PropertyLetStmt : 
	(Visibility WS)? (STATIC WS)? PROPERTY_LET WS AmbiguousIdentifier (WS? ArgList)? NEWLINE+ 
	(Block NEWLINE+)? 
	END_PROPERTY
;

PutStmt : PUT WS ValueStmt WS? ',' WS? ValueStmt? WS? ',' WS? ValueStmt;

RaiseEventStmt : RAISEEVENT WS AmbiguousIdentifier (WS? LPAREN WS? (ArgsCall WS?)? RPAREN)?;

RandomizeStmt : RANDOMIZE (WS ValueStmt)?;

RedimStmt : REDIM WS (PRESERVE WS)? RedimSubStmt (WS?',' WS? RedimSubStmt)*;

RedimSubStmt : ImplicitCallStmt_InStmt WS? LPAREN WS? Subscripts WS? RPAREN (WS AsTypeClause)?;

ResetStmt : RESET;

ResumeStmt : RESUME (WS (NEXT | AmbiguousIdentifier))?;

ReturnStmt : RETURN;

RmdirStmt : RMDIR WS ValueStmt;

RSetStmt : RSET WS ImplicitCallStmt_InStmt WS? EQ WS? ValueStmt;

SavePictureStmt : SAVEPICTURE WS ValueStmt WS? ',' WS? ValueStmt;

SaveSettingStmt : SAVESETTING WS ValueStmt WS? ',' WS? ValueStmt WS? ',' WS? ValueStmt WS? ',' WS? ValueStmt;

SeekStmt : SEEK WS ValueStmt WS? ',' WS? ValueStmt;

SelectCaseStmt : 
	SELECT WS CASE WS ValueStmt NEWLINE+ 
	SC_Case* 
	SC_CaseElse?
	WS? END_SELECT
;

SC_Case : 
	CASE WS SC_Cond WS? (':'? NEWLINE* | NEWLINE+)  
	(Block NEWLINE+)?
;

SC_Cond : 
	IS WS? ComparisonOperator WS? ValueStmt # caseCondIs
	| ValueStmt (WS? ',' WS? ValueStmt)* # caseCondValue
	| INTEGERLITERAL WS TO WS ValueStmt (WS? ',' WS? ValueStmt)* # caseCondTo
;

SC_CaseElse : 
	CASE WS ELSE WS? (':'? NEWLINE* | NEWLINE+)  
	(Block NEWLINE+)?
;

SendkeysStmt : SENDKEYS WS ValueStmt (WS? ',' WS? ValueStmt)?;

SetAttrStmt : SETATTR WS ValueStmt WS? ',' WS? ValueStmt;

SetStmt : SET WS ImplicitCallStmt_InStmt WS? EQ WS? ValueStmt;

StopStmt : STOP;

SubStmt : 
	(Visibility WS)? (STATIC WS)? SUB WS AmbiguousIdentifier (WS? ArgList)? NEWLINE+ 
	(Block NEWLINE+)? 
	END_SUB
;

TimeStmt : TIME WS? EQ WS? ValueStmt;

TypeStmt : 
	(Visibility WS)? Type WS AmbiguousIdentifier NEWLINE+ 
	(TypeStmt_Element)*
	END_Type
;

TypeStmt_Element : AmbiguousIdentifier (WS? LPAREN (WS? Subscripts)? WS? RPAREN)? (WS AsTypeClause)? NEWLINE+;

TypeOfStmt : TypeOF WS ValueStmt (WS IS WS Type)?;

UnloadStmt : UNLOAD WS ValueStmt;

UnlockStmt : UNLOCK WS ValueStmt (WS? ',' WS? ValueStmt (WS TO WS ValueStmt)?)?;

ValueStmt : 
	Literal # vsLiteral
	| MidStmt # vsMid
	| NEW WS ValueStmt # vsNew
	| ImplicitCallStmt_InStmt # vsValueCalls
	| TypeOfStmt # vsTypeOf
	| LPAREN WS? ValueStmt (WS? ',' WS? ValueStmt)* RPAREN # vsStruct
	| ImplicitCallStmt_InStmt WS? ASSIGN WS? ValueStmt # vsAssign
	| ValueStmt WS? PLUS WS? ValueStmt # vsAdd
	| PLUS WS? ValueStmt # vsPlus
	| ADDRESSOF WS ValueStmt # vsAddressOf
	| ValueStmt WS AMPERSAND WS ValueStmt # vsAmp
	| ValueStmt WS AND WS ValueStmt # vsAnd
	| ValueStmt WS? LT WS? ValueStmt # vsLt
	| ValueStmt WS? LEQ WS? ValueStmt # vsLeq
	| ValueStmt WS? GT WS? ValueStmt # vsGt
	| ValueStmt WS? GEQ WS? ValueStmt # vsGeq
	| ValueStmt WS? EQ WS? ValueStmt # vsEq
	| ValueStmt WS? NEQ WS? ValueStmt # vsNeq
	| ValueStmt WS? DIV WS? ValueStmt # vsDiv
	| ValueStmt WS EQV WS ValueStmt # vsEqv
	| ValueStmt WS IMP WS ValueStmt # vsImp
	| ValueStmt WS IS WS ValueStmt # vsIs
	| ValueStmt WS LIKE WS ValueStmt # vsLike
	| ValueStmt WS? MINUS WS? ValueStmt # vsMinus
	| MINUS WS? ValueStmt # vsNegation
	| ValueStmt WS? MOD WS? ValueStmt # vsMod
	| ValueStmt WS? MULT WS? ValueStmt # vsMult
	| NOT WS ValueStmt # vsNot
	| ValueStmt WS? OR WS? ValueStmt # vsOr
	| ValueStmt WS? POW WS? ValueStmt # vsPow
	| ValueStmt WS? XOR WS? ValueStmt # vsXor
;

VariableStmt : (DIM | STATIC | Visibility) WS (WITHEVENTS WS)? VariableListStmt;

VariableListStmt : VariableSubStmt (WS? ',' WS? VariableSubStmt)*;

VariableSubStmt : AmbiguousIdentifier (WS? LPAREN WS? (Subscripts WS?)? RPAREN WS?)? TypeHint? (WS AsTypeClause)?;

WhileWendStmt : 
	WHILE WS ValueStmt NEWLINE+ 
	(Block NEWLINE)* 
	WEND
;

WidthStmt : WIDTH WS ValueStmt WS? ',' WS? ValueStmt;

WithStmt : 
	WITH WS ImplicitCallStmt_InStmt NEWLINE+ 
	(Block NEWLINE+)? 
	END_WITH
;

WriteStmt : WRITE WS ValueStmt WS? ',' (WS? OutputList)?;


// complex call statements ----------------------------------

ExplicitCallStmt : 
	ECS_ProcedureCall 
	| ECS_MemberProcedureCall 
;

ECS_ProcedureCall : CALL WS AmbiguousIdentifier TypeHint? (WS? LPAREN WS? (ArgsCall WS?)? RPAREN)?;

ECS_MemberProcedureCall : CALL WS VariableCallStmt? MemberPropertyCallStmt* '.' AmbiguousIdentifier TypeHint? (WS? LPAREN WS? (ArgsCall WS?)? RPAREN)?;


ImplicitCallStmt_InBlock :
	ICS_B_SubCall
	| ICS_B_FunctionCall
	| ICS_B_MemberSubCall
	| ICS_B_MemberFunctionCall
;

// CertainIdentifier instead of AmbiguousIdentifier for preventing ambiguity with statement keywords 
ICS_B_SubCall : CertainIdentifier (WS ArgsCall)?;

ICS_B_FunctionCall : FunctionOrArrayCallStmt DictionaryCallStmt?;

ICS_B_MemberSubCall : ImplicitCallStmt_InStmt* MemberSubCallStmt;

ICS_B_MemberFunctionCall : ImplicitCallStmt_InStmt* MemberFunctionOrArrayCallStmt DictionaryCallStmt?;


ImplicitCallStmt_InStmt :
	ICS_S_VariableCall
	| ICS_S_FunctionOrArrayCall
	| ICS_S_DictionaryCall
	| ICS_S_MembersCall
;

ICS_S_VariableCall : VariableCallStmt DictionaryCallStmt?;

ICS_S_FunctionOrArrayCall : FunctionOrArrayCallStmt DictionaryCallStmt?;

ICS_S_DictionaryCall : DictionaryCallStmt;

ICS_S_MembersCall : (VariableCallStmt | FunctionOrArrayCallStmt)? MemberCall_Value+ DictionaryCallStmt?;


// member call statements ----------------------------------

MemberPropertyCallStmt : '.' AmbiguousIdentifier;

MemberFunctionOrArrayCallStmt : '.' FunctionOrArrayCallStmt;

MemberSubCallStmt : '.' AmbiguousIdentifier (WS ArgsCall)?;

MemberCall_Value : MemberPropertyCallStmt | MemberFunctionOrArrayCallStmt;


// atomic call statements ----------------------------------

VariableCallStmt : AmbiguousIdentifier TypeHint?;

DictionaryCallStmt : '!' AmbiguousIdentifier TypeHint?;

FunctionOrArrayCallStmt : (AmbiguousIdentifier | BaseType) TypeHint? WS? LPAREN WS? (ArgsCall WS?)? RPAREN;


ArgsCall : (ArgCall? WS? (',' | ';') WS?)* ArgCall (WS? (',' | ';') WS? ArgCall?)*;

ArgCall : ((BYVAL | BYREF | PARAMARRAY) WS)? ValueStmt;


// atomic rules for statements

ArgList : LPAREN (WS? Arg (WS? ',' WS? Arg)*)? WS? RPAREN;

Arg : (OPTIONAL WS)? ((BYVAL | BYREF) WS)? (PARAMARRAY WS)? AmbiguousIdentifier (WS? LPAREN WS? RPAREN)? (WS AsTypeClause)? (WS? ArgDefaultValue)?;

ArgDefaultValue : EQ WS? (Literal | AmbiguousIdentifier);

Subscripts : Subscript (WS? ',' WS? Subscript)*;

Subscript : (ValueStmt WS TO WS)? ValueStmt;


// atomic rules ----------------------------------

AmbiguousIdentifier : 
	(IDENTIFIER | AmbiguousKeyword)+
	| L_SQUARE_BRACKET (IDENTIFIER | AmbiguousKeyword)+ R_SQUARE_BRACKET
;

AsTypeClause : AS WS (NEW WS)? Type (WS FieldLength)?;

BaseType : BOOLEAN | BYTE | COLLECTION | DATE | DOUBLE | INTEGER | LONG | SINGLE | STRING | VARIANT;

CertainIdentifier : 
	IDENTIFIER (AmbiguousKeyword | IDENTIFIER)*
	| AmbiguousKeyword (AmbiguousKeyword | IDENTIFIER)+
;

ComparisonOperator : LT | LEQ | GT | GEQ | EQ | NEQ | IS | LIKE;

ComplexType : AmbiguousIdentifier ('.' AmbiguousIdentifier)*;

FieldLength : MULT WS? (INTEGERLITERAL | AmbiguousIdentifier);

LetterRange : CertainIdentifier (WS? MINUS WS? CertainIdentifier)?;

LineLabel : AmbiguousIdentifier ':';

Literal : COLORLITERAL | DATELITERAL | DOUBLELITERAL | FILENUMBER | INTEGERLITERAL | STRINGLITERAL | TRUE | FALSE | NOTHING | NULL;

Type : (BaseType | ComplexType) (WS? LPAREN WS? RPAREN)?;

TypeHint : '&' | '%' | '#' | '!' | '@' | '$';

Visibility : PRIVATE | PUBLIC | FRIEND | GLOBAL;


// ambiguous keywords
AmbiguousKeyword : 
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
	TAB | TEXT | THEN | TIME | TO | TRUE | Type | TypeOF | 
	UNLOAD | UNLOCK | UNTIL | 
	VARIANT | VERSION | 
	WEND | WHILE | WIDTH | WITH | WITHEVENTS | WRITE |
	XOR
;


// lexer rules --------------------------------------------------------------------------------


// keywords
ACCESS : A C C E S S;
ADDRESSOF : A D D R E S S O F;
ALIAS : A L I A S;
AND : A N D;
ATTRIBUTE : A T T R I B U T E;
APPACTIVATE : A P P A C T I V A T E;
APPEND : A P P E N D;
AS : A S;
BEGIN : B E G I N;
BEEP : B E E P;
BINARY : B I N A R Y;
BOOLEAN : B O O L E A N;
BYVAL : B Y V A L;
BYREF : B Y R E F;
BYTE : B Y T E;
CALL : C A L L;
CASE : C A S E;
CHDIR : C H D I R;
CHDRIVE : C H D R I V E;
CLASS : C L A S S;
CLOSE : C L O S E;
COLLECTION : C O L L E C T I O N;
CONST : C O N S T;
DATABASE : D A T A B A S E;
DATE : D A T E;
DECLARE : D E C L A R E;
DEFBOOL : D E F B O O L; 
DEFBYTE : D E F B Y T E;
DEFDATE : D E F D A T E;
DEFDBL : D E F D B L;
DEFDEC : D E F D E C;
DEFCUR : D E F C U R;
DEFINT : D E F I N T;
DEFLNG : D E F L N G;
DEFOBJ : D E F O B J;
DEFSNG : D E F S N G;
DEFSTR : D E F S T R;
DEFVAR : D E F V A R;
DELETESETTING : D E L E T E S E T T I N G;
DIM : D I M;
DO : D O;
DOUBLE : D O U B L E;
EACH : E A C H;
ELSE : E L S E;
ELSEIF : E L S E I F;
END_ENUM : E N D ' ' E N U M;
END_FUNCTION : E N D ' ' F U N C T I O N;
END_IF : E N D ' ' I F;
END_PROPERTY : E N D ' ' P R O P E R T Y;
END_SELECT : E N D ' ' S E L E C T;
END_SUB : E N D ' ' S U B;
END_Type : E N D ' ' T Y P E;
END_WITH : E N D ' ' W I T H;
END : E N D;
ENUM : E N U M;
EQV : E Q V;
ERASE : E R A S E;
ERROR : E R R O R;
EVENT : E V E N T;
EXIT_DO : E X I T ' ' D O;
EXIT_FOR : E X I T ' ' F O R;
EXIT_FUNCTION : E X I T ' ' F U N C T I O N;
EXIT_PROPERTY : E X I T ' ' P R O P E R T Y;
EXIT_SUB : E X I T ' ' S U B;
FALSE : F A L S E;
FILECOPY : F I L E C O P Y;
FRIEND : F R I E N D;
FOR : F O R;
FUNCTION : F U N C T I O N;
GET : G E T;
GLOBAL : G L O B A L;
GOSUB : G O S U B;
GOTO : G O T O;
IF : I F;
IMP : I M P;
IMPLEMENTS : I M P L E M E N T S;
IN : I N;
INPUT : I N P U T;
IS : I S;
INTEGER : I N T E G E R;
KILL: K I L L;
LOAD : L O A D;
LOCK : L O C K;
LONG : L O N G;
LOOP : L O O P;
LEN : L E N;
LET : L E T;
LIB : L I B;
LIKE : L I K E;
LINE_INPUT : L I N E ' ' I N P U T;
LOCK_READ : L O C K ' ' R E A D;
LOCK_WRITE : L O C K ' ' W R I T E;
LOCK_READ_WRITE : L O C K ' ' R E A D ' ' W R I T E;
LSET : L S E T;
MACRO_IF : '#' I F;
MACRO_ELSEIF : '#' E L S E I F;
MACRO_ELSE : '#' E L S E;
MACRO_END_IF : '#' E N D ' ' I F;
ME : M E;
MID : M I D;
MKDIR : M K D I R;
MOD : M O D;
NAME : N A M E;
NEXT : N E X T;
NEW : N E W;
NOT : N O T;
NOTHING : N O T H I N G;
NULL : N U L L;
ON : O N;
ON_ERROR : O N ' ' E R R O R;
OPEN : O P E N;
OPTIONAL : O P T I O N A L;
OPTION_BASE : O P T I O N ' ' B A S E;
OPTION_EXPLICIT : O P T I O N ' ' E X P L I C I T;
OPTION_COMPARE : O P T I O N ' ' C O M P A R E;
OPTION_PRIVATE_MODULE : O P T I O N ' ' P R I V A T E ' ' M O D U L E;
OR : O R;
OUTPUT : O U T P U T;
PARAMARRAY : P A R A M A R R A Y;
PRESERVE : P R E S E R V E;
PRINT : P R I N T;
PRIVATE : P R I V A T E;
PROPERTY_GET : P R O P E R T Y ' ' G E T;
PROPERTY_LET : P R O P E R T Y ' ' L E T;
PROPERTY_SET : P R O P E R T Y ' ' S E T;
PUBLIC : P U B L I C;
PUT : P U T;
RANDOM : R A N D O M;
RANDOMIZE : R A N D O M I Z E;
RAISEEVENT : R A I S E E V E N T;
READ : R E A D;
READ_WRITE : R E A D ' ' W R I T E;
REDIM : R E D I M;
REM : R E M;
RESET : R E S E T;
RESUME : R E S U M E;
RETURN : R E T U R N;
RMDIR : R M D I R;
RSET : R S E T;
SAVEPICTURE : S A V E P I C T U R E;
SAVESETTING : S A V E S E T T I N G;
SEEK : S E E K;
SELECT : S E L E C T;
SENDKEYS : S E N D K E Y S;
SET : S E T;
SETATTR : S E T A T T R;
SHARED : S H A R E D;
SINGLE : S I N G L E;
SPC : S P C;
STATIC : S T A T I C;
STEP : S T E P;
STOP : S T O P;
STRING : S T R I N G;
SUB : S U B;
TAB : T A B;
TEXT : T E X T;
THEN : T H E N;
TIME : T I M E;
TO : T O;
TRUE : T R U E;
Type : T Y P E;
TypeOF : T Y P E O F;
UNLOAD : U N L O A D;
UNLOCK : U N L O C K;
UNTIL : U N T I L;
VARIANT : V A R I A N T;
VERSION : V E R S I O N;
WEND : W E N D;
WHILE : W H I L E;
WIDTH : W I D T H;
WITH : W I T H;
WITHEVENTS : W I T H E V E N T S;
WRITE : W R I T E;
XOR : X O R;


// symbols
AMPERSAND : '&';
ASSIGN : ':=';
DIV : '\\' | '/';
EQ : '=';
GEQ : '>=';
GT : '>';
LEQ : '<=';
LPAREN : '(';
LT : '<';
MINUS : '-';
MULT : '*';
NEQ : '<>';
PLUS : '+';
POW : '^';
RPAREN : ')';
L_SQUARE_BRACKET : '[';
R_SQUARE_BRACKET : ']';


// literals
STRINGLITERAL : '"' (~["\r\n] | '""')* '"';
DATELITERAL : '#' (~[#\r\n])* '#';
COLORLITERAL : '&H' [0-9A-F]+ '&'?;
INTEGERLITERAL : (PLUS|MINUS)? ('0'..'9')+ ( ('e' | 'E') INTEGERLITERAL)* ('#' | '&')?;
DOUBLELITERAL : (PLUS|MINUS)? ('0'..'9')* '.' ('0'..'9')+ ( ('e' | 'E') (PLUS|MINUS)? ('0'..'9')+)* ('#' | '&')?;
FILENUMBER : '#' LETTERORDIGIT+;
// identifier
IDENTIFIER : LETTER (LETTERORDIGIT)*;
// whitespace, line breaks, comments, ...
LINE_CONTINUATION : ' ' '_' '\r'? '\n' -> skip;
NEWLINE : WS? ('\r'? '\n' | ':' ' ') WS?;
COMMENT : WS? ('\'' | ':'? REM ' ') (LINE_CONTINUATION | ~('\n' | '\r'))* -> skip;
WS : [ \t]+;


// letters
fragment LETTER : [a-zA-Z_‰ˆ¸ƒ÷‹];
fragment LETTERORDIGIT : [a-zA-Z0-9_‰ˆ¸ƒ÷‹];

// case insensitive chars
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