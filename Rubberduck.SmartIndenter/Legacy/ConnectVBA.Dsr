VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   12600
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   12960
   _ExtentX        =   22860
   _ExtentY        =   22225
   _Version        =   393216
   Description     =   $"ConnectVBA.dsx":0000
   DisplayName     =   "Smart Indenter v3.5"
   AppName         =   "Visual Basic for Applications IDE"
   AppVer          =   "6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0"
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'*
'* MODULE NAME:     INDENTER FOR VB6/VBA
'* AUTHOR & DATE:   STEPHEN BULLEN, Office Automation Ltd.
'*                  12 May 1998
'*
'* CONTACT:         stephen@oaltd.co.uk
'* WEB SITE:        http://www.oaltd.co.uk
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
'*   AddinInstance_OnConnection          Routine run when addin is loaded
'*   AddinInstance_OnDisconnection       Routine run when addin is closed
'*
'***************************************************************************

Option Explicit

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

    Set poVBE = Application
    SetUpMenus

End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    On Error Resume Next

    Unload frmProgress
    Unload frmVBEOnKey
    Unload frmAbout
    Unload frmOptions

    RemoveMenus

End Sub
