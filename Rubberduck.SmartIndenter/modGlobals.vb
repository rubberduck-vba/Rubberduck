Option Explicit On
Option Compare Text

Imports Microsoft.Vbe.Interop

Module modGlobals
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
    '*
    '***************************************************************************
    '*
    '* CHANGE HISTORY
    '*
    '*  DATE        NAME                DESCRIPTION
    '*  07/15/1999  Stephen Bullen      Initial version
    '*  11/26/2015  Mathieu Guindon     Slight modifications to build in VB.NET
    '*
    '***************************************************************************

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
    Public poBtnUndo As New Collection
    Public poVBE As VBE

    'UDT to store Undo information
    Public Structure uUndo
        Public oMod As CodeModule
        Public sName As String
        Public lStartLine As Long
        Public lEndLine As Long
        Public asOriginal() As String
        Public asIndented() As String
    End Structure

    Public pauUndo() As uUndo
    Public piUndoCount As Integer

    Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

End Module
