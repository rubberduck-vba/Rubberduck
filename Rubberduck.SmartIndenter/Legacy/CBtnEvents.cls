VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBtnEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'* THIS MODULE:     Contains the class to trap the menu item clicking
'*
'* PROCEDURES:
'*   oHook_Click    Handles clicking on the VBE menu items
'*
'***************************************************************************

Option Explicit
Option Compare Text
Option Base 1

Public WithEvents oHook As CommandBarEvents
Attribute oHook.VB_VarHelpID = -1

Private Sub oHook_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    'Do the appropriate action, depending on the menu item's parameter value
    Select Case CommandBarControl.Parameter
    Case "Indent2KProc"
        IndentProcedure

    Case "Indent2KMod"
        IndentModule

    Case "Indent2KProj"
        IndentProject

    Case "Indent2KUndo"
        UndoIndenting

    Case "Indent2KForm"
        frmOptions.Show vbModal

    Case "Indent2KAbout"
        frmAbout.Show vbModal

    Case "Indent2KPrjWin"
        IndentFromProjectWindow
    End Select

End Sub

