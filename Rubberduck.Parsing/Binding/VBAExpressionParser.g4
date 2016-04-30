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

parser grammar VBAExpressionParser;

options { tokenVocab = VBALexer; }

startRule : expression EOF;

// 5.1 Module Body Structure
unrestrictedName : name | reservedIdentifierName;
name : untypedName | typedName;
// Added to allow expressions like VBA.String$
reservedIdentifierName : reservedUntypedName | reservedTypedName;
reservedUntypedName : reservedIdentifier;
reservedTypedName : reservedIdentifier typeSuffix;
untypedName : IDENTIFIER | FOREIGNNAME | reservedProcedureName | specialForm | optionCompareArgument | OBJECT | uncategorizedKeyword | ERROR;
typedName : typedNameValue typeSuffix;
typedNameValue : IDENTIFIER | reservedProcedureName | specialForm | optionCompareArgument | OBJECT | uncategorizedKeyword | ERROR;
typeSuffix : PERCENT | AMPERSAND | POW | EXCLAMATIONPOINT | HASH | AT | DOLLAR;

optionCompareArgument : BINARY | TEXT | DATABASE;

// 3.3.5.3 Special Identifier Forms
builtInType : 
    reservedTypeIdentifier
    | L_SQUARE_BRACKET whiteSpace? reservedTypeIdentifier whiteSpace? R_SQUARE_BRACKET
    | OBJECT
    | L_SQUARE_BRACKET whiteSpace? OBJECT whiteSpace? R_SQUARE_BRACKET
;

// 5.6 Expressions
expression :
    lExpression                                                                                     # lExpr
    | LPAREN whiteSpace? expression whiteSpace? RPAREN                                              # parenthesizedExpr
    | typeOfIsExpression                                                                            # typeOfIsExpr
    | newExpression                                                                                 # newExpr
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
;

// 5.6.5 Literal Expressions
literalExpression :
    numberLiteral
    | DATELITERAL
    | STRINGLITERAL
    | literalIdentifier typeSuffix?
;
numberLiteral : HEXLITERAL | OCTLITERAL | FLOATLITERAL | INTEGERLITERAL;

// 5.6.6 Parenthesized Expressions
parenthesizedExpression : LPAREN whiteSpace? expression whiteSpace? RPAREN;

// 5.6.7 TypeOf…Is Expressions
typeOfIsExpression : TYPEOF whiteSpace expression whiteSpace IS whiteSpace typeExpression;

// 5.6.8 New Expressions
newExpression : NEW whiteSpace typeExpression;

lExpression :
    lExpression whiteSpace? LPAREN whiteSpace? argumentList? whiteSpace? RPAREN                   # indexExpr
    | lExpression DOT unrestrictedName                                                            # memberAccessExpr
    | lExpression LINE_CONTINUATION whiteSpace? DOT unrestrictedName                              # memberAccessExpr
    | lExpression EXCLAMATIONPOINT unrestrictedName                                               # dictionaryAccessExpr
    | lExpression LINE_CONTINUATION EXCLAMATIONPOINT unrestrictedName                             # dictionaryAccessExpr
    | lExpression LINE_CONTINUATION EXCLAMATIONPOINT LINE_CONTINUATION unrestrictedName           # dictionaryAccessExpr
    | instanceExpression                                                                          # instanceExpr
    | simpleNameExpression                                                                        # simpleNameExpr
    | withExpression                                                                              # withExpr
;

// 5.6.12 Member Access Expressions
memberAccessExpression :
    lExpression DOT unrestrictedName
    | lExpression LINE_CONTINUATION whiteSpace? DOT unrestrictedName
;

// 5.6.13 Index Expressions
indexExpression : lExpression whiteSpace? LPAREN whiteSpace? argumentList? whiteSpace? RPAREN;

// 5.6.14 Dictionary Access Expressions
dictionaryAccessExpression :
    lExpression EXCLAMATIONPOINT unrestrictedName
    | lExpression LINE_CONTINUATION EXCLAMATIONPOINT unrestrictedName
    | lExpression LINE_CONTINUATION EXCLAMATIONPOINT LINE_CONTINUATION unrestrictedName
;

// 5.6.13.1 Argument Lists
argumentList : positionalOrNamedArgumentList;
positionalOrNamedArgumentList :
    (positionalArgument? whiteSpace? COMMA whiteSpace?)* requiredPositionalArgument 
    | (positionalArgument? whiteSpace? COMMA whiteSpace?)* namedArgumentList  
;
positionalArgument : argumentExpression;
requiredPositionalArgument : argumentExpression;  
namedArgumentList : namedArgument (whiteSpace? COMMA whiteSpace? namedArgument)*;
namedArgument : unrestrictedName whiteSpace? ASSIGN whiteSpace? argumentExpression;
argumentExpression :
    (BYVAL whiteSpace)? expression
    | addressOfExpression
;

// 5.6.10 Simple Name Expressions
simpleNameExpression : name;

// 5.6.11 Instance Expressions
instanceExpression : ME;

// 5.6.15 With Expressions
withExpression : withMemberAccessExpression | withDictionaryAccessExpression;
withMemberAccessExpression : DOT unrestrictedName;  
withDictionaryAccessExpression : EXCLAMATIONPOINT unrestrictedName;

// 5.6.16.1 Constant Expressions
constantExpression : expression;

// 5.6.16.7 Type Expressions
typeExpression : builtInType | definedTypeExpression;
definedTypeExpression : simpleNameExpression | memberAccessExpression;

// 5.6.16.8   AddressOf Expressions 
addressOfExpression : ADDRESSOF whiteSpace procedurePointerExpression;
procedurePointerExpression : memberAccessExpression | simpleNameExpression;

// 3.3.5.2   Reserved Identifiers and IDENTIFIER
reservedIdentifier :
    statementKeyword
    | markerKeyword
    | operatorIdentifier
    | specialForm
    | reservedName
    | literalIdentifier
    | remKeyword
    | reservedTypeIdentifier // Added to allow expressions like VBA.String$
;
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
    | END_IF
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
remKeyword : REM;
markerKeyword :
    ANY
    | AS
    | BYREF
    | BYVAL
    | CASE
    | EACH
    | ELSE
    | IN
    | NEW
    | SHARED
    | UNTIL
    | WITHEVENTS
    | WRITE
    | OPTIONAL
    | PARAMARRAY
    | PRESERVE
    | SPC
    | TAB
    | THEN
    | TO
;
operatorIdentifier :
    ADDRESSOF
    | AND
    | EQV
    | IMP
    | IS
    | LIKE
    | NEW
    | MOD
    | NOT
    | OR
    | TYPEOF
    | XOR
;
reservedName :
    ME
    | reservedProcedureName
;
reservedProcedureName :
    ABS
    | CBOOL
    | CBYTE
    | CCUR
    | CDATE
    | CDBL
    | CDEC
    | CINT
    | CLNG
    | CLNGLNG
    | CLNGPTR
    | CSNG
    | CSTR
    | CVAR
    | CVERR
    | DEBUG
    | DOEVENTS
    | FIX
    | INT
    | LEN
    | LENB
    | PSET
    | SCALE
    | SGN
    | MID 
    | MIDB 
    | MIDTYPESUFFIX 
    | MIDBTYPESUFFIX
;
specialForm :
    ARRAY
    | CIRCLE
    | INPUT
    | INPUTB
    | LBOUND
    | SCALE
    | UBOUND
;
reservedTypeIdentifier :
    BOOLEAN
    | BYTE
    | CURRENCY
    | DATE
    | DOUBLE
    | INTEGER
    | LONG
    | LONGLONG
    | LONGPTR
    | SINGLE
    | STRING
    | VARIANT
;

uncategorizedKeyword : 
	ALIAS | ATTRIBUTE | APPACTIVATE |
	BEEP | BEGIN | CLASS | CHDIR | CHDRIVE | COLLECTION | DELETESETTING |
	FILECOPY | KILL | LOAD | LIB | MKDIR | NAME | ON |
	RANDOMIZE | RMDIR |
	SAVEPICTURE | SAVESETTING | SENDKEYS | SETATTR |
	TAB | TIME | UNLOAD | VERSION
;

literalIdentifier : booleanLiteralIdentifier | objectLiteralIdentifier | variantLiteralIdentifier;
booleanLiteralIdentifier : TRUE | FALSE;
objectLiteralIdentifier : NOTHING;
variantLiteralIdentifier : EMPTY | NULL;

whiteSpace : (WS | LINE_CONTINUATION)+;