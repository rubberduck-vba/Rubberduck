VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'*
'* MODULE NAME:     INDENTER FOR VB5
'* AUTHOR & DATE:   STEPHEN BULLEN, Business Modelling Solutions Ltd.
'*                  12 May 1998
'*
'* CONTACT:         Stephen@BMSLtd.co.uk
'* WEB SITE:        http://www.BMSLtd.co.uk
'*
'* DESCRIPTION:     Adds items to the VBE environment to recreate the indenting
'*                  for the current procedure, module or project.
'*
'* THIS MODULE:     Sets up the menu items in the VBE and handles their events.
'*                  Command bar controls in the VBE do not have a working OnAction
'*                  property, so we have to trap and respond to the command bar
'*                  events instead.  This is done by the vbeMenus class array.
'*
'* PROCEDURES:
'*   IDTExtensibility_OnConnection       Routine run when addin is loaded
'*   IDTExtensibility_OnDisconnection    Routine run when addin is closed
'*   IDTExtensibility_OnAddInsUpdate     Stub for addin connection
'*   IDTExtensibility_OnStartupComplete  Stub for addin connection
'*   SetUpEditMenus                      Adds our menu items to the Edit menu
'*   SetUpCodeWindowMenus                Adds our menu items to the Code Window popup menu
'*
'***************************************************************************

Option Explicit
Option Compare Text
Option Base 1

Implements IDTExtensibility

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)

    Set poVBE = VBInst
    SetUpMenus

End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)

    On Error Resume Next

    Unload frmProgress
    Unload frmVBEOnKey
    Unload frmAbout
    Unload frmOptions

    RemoveMenus

End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)

End Sub
