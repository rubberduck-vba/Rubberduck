Option Explicit On
Option Compare Text

Imports Microsoft.Vbe.Interop

Public Class CBtnEvents
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
    '*
    '* CHANGE HISTORY
    '*
    '*  DATE        NAME                DESCRIPTION
    '*  07/15/1999  Stephen Bullen      Initial version
    '*  11/26/2015  Mathieu Guindon     Slight modifications to build in VB.NET
    '*
    '***************************************************************************

    Public WithEvents oHook As CommandBarEvents

    Private Sub oHook_Click(ByVal CommandBarControl As Object, ByRef handled As Boolean, ByRef CancelDefault As Boolean) Handles oHook.Click

        'Do the appropriate action, depending on the menu item's parameter value
        Select Case CommandBarControl.Parameter
        Case "Indent2KProc":
            'IndentProcedure

        Case "Indent2KMod":
            'IndentModule

        Case "Indent2KProj":
            'IndentProject

        Case "Indent2KUndo":
            'UndoIndenting

        Case "Indent2KForm":
            'todo: tell Rubberduck to bring up the SmartIndenter page of the Rubberduck options dialog
            'frmOptions.Show vbModal

        Case "Indent2KAbout":
            'todo: remove the "About" menu item (Rubberduck's "About" window credits SmartIndenter)
            'frmAbout.Show vbModal

        Case "Indent2KPrjWin":
            'IndentFromProjectWindow
        End Select

    End Sub

End Class
