Attribute VB_Name = "modGlobals"
'***************************************************************************
'*
'* PROJECT NAME:    SMART INDENTER VB6
'* AUTHOR & DATE:   STEPHEN BULLEN, Office Automation Ltd.
'*                  15 July 1999
'*
'*                  COPYRIGHT © 1999-2004 BY OFFICE AUTOMATION LTD
'*
'* CONTACT:         stephen@oaltd.co.uk
'* WEB SITE:        http://www.oaltd.co.uk
'*
'* DESCRIPTION:     Adds items to the VBE environment to recreate the indenting
'*                  for the current procedure, module or project.
'*
'* THIS MODULE:     Contains global variables and constants used in the addin
'*
'* PROCEDURES:
'***************************************************************************

Option Explicit
Option Compare Text
Option Base 1

'Define some public constants for the Registry keys
Public Const psREG_SECTION As String = "Office Automation Ltd."
Public Const psAPP_TITLE As String = "Smart Indenter VB"
Public Const psREG_KEY As String = "Smart Indenter"
Public Const psMENU_TAG As String = "INDENT2K"

'Define some constants for the text of the menu items
Public Const msMENU_TEXT As String = "&Smart Indent"
Public Const msMENU_PROC As String = "Indent &Procedure"
Public Const msMENU_MOD As String = "Indent &Module"
Public Const msMENU_PROJ As String = "Indent Pro&ject"
Public Const msMENU_UNDO As String = "&Undo Indenting"
Public Const msMENU_FORM As String = "&Indenting Options"
Public Const msMENU_ABOUT As String = "About Smart Indenter"
Public Const msMENU_PROJ_WIN As String = "Indent"

'Define a collection to store the menu item click event handlers
Public poBtnEvents As New Collection
Attribute poBtnEvents.VB_VarUserMemId = 1073741824
Public poBtnUndo As New Collection
Attribute poBtnUndo.VB_VarUserMemId = 1073741825
Public poVBE As VBIDE.VBE
Attribute poVBE.VB_VarUserMemId = 1073741826

'UDT to store Undo information
Public Type uUndo
    oMod As CodeModule
    sName As String
    lStartLine As Long
    lEndLine As Long
    asOriginal() As String
    asIndented() As String
End Type

Public pauUndo() As uUndo
Attribute pauUndo.VB_VarUserMemId = 1073741827
Public piUndoCount As Integer
Attribute piUndoCount.VB_VarUserMemId = 1073741828

Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Attribute SetParent.VB_UserMemId = 1879048240

