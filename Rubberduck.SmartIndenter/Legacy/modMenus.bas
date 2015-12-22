Attribute VB_Name = "modMenus"
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
'* THIS MODULE:     Sets up the menu items in the VBE and handles their events.
'*                  Command bar controls in the VBE do not have a working OnAction
'*                  property, so we have to trap and respond to the command bar
'*                  events instead.  This is done by the vbeMenus class array.
'*
'* PROCEDURES:
'*  SetUpMenus                      Add our menus to the VBE
'*  CreateSubMenus                  Add our cascading menus
'*  RemoveMenus                     Remove our menus from the VBE
'*
'***************************************************************************

Option Explicit
Option Compare Text
Option Base 1


'
' Adds the menus to the VBE Code Window menu and sets up the objects to trap
' the command bar events for the new controls
'
Sub SetUpMenus()
Attribute SetUpMenus.VB_UserMemId = 1610612736

    Dim oBar As CommandBar, oCtl As CommandBarControl, oMenu As CommandBarPopup
    Dim oEvt As CBtnEvents, oBtn As CommandBarButton

    'Ignore errors, so we can set objects, then check if they were set OK
    On Error Resume Next

    RemoveMenus

    'Hook 'Indent Procedure' hot-key, if enabled
    If (GetSetting(psREG_SECTION, psREG_KEY, "EnableProcHotKey", "Y") = "Y") Then
        VBEOnKey GetSetting(psREG_SECTION, psREG_KEY, "ProcHotKeyCode", "+^P"), "Indent2KProc"
    End If

    'Hook 'Indent Module' hot-key, if enabled
    If (GetSetting(psREG_SECTION, psREG_KEY, "EnableModHotKey", "Y") = "Y") Then
        VBEOnKey GetSetting(psREG_SECTION, psREG_KEY, "ModHotKeyCode", "+^M"), "Indent2KMod"
    End If

    'About box on Addins menu
    Set oMenu = poVBE.CommandBars.ActiveMenuBar.FindControl(id:=30038)    '30038 = Addins Menu
    Set oBar = oMenu.CommandBar

    Set oBtn = oBar.Controls.Add(Type:=msoControlButton, Parameter:="Indent2KAbout", temporary:=True)
    With oBtn
        .Caption = msMENU_ABOUT
        .FaceId = 2564
        .Style = msoButtonIconAndCaption
        .Tag = psMENU_TAG
    End With
    Set oEvt = New CBtnEvents
    Set oEvt.oHook = poVBE.Events.CommandBarEvents(oBtn)
    poBtnEvents.Add oEvt

    'VBE Project Explorer
    If IsOfficeVBE Then
        AddProjectExplorerMenu "Project Window"
    Else
        AddProjectExplorerMenu "Project Window Project"
        AddProjectExplorerMenu "Project Window Form Folder"
        AddProjectExplorerMenu "Project Window Module/Class Folder"
    End If

    ' VBE Code Window shortcut Menu
    Set oBar = poVBE.CommandBars("Code Window")

    Set oCtl = oBar.FindControl(id:=473)         '473 = Object Browser

    If oCtl Is Nothing Then
        Set oMenu = oBar.Controls.Add(Type:=msoControlPopup, temporary:=True)
    Else
        Set oMenu = oBar.Controls.Add(Type:=msoControlPopup, before:=oCtl.Index, temporary:=True)
    End If

    oMenu.BeginGroup = True

    CreateSubMenus oMenu

    ' VBE Edit Menu
    Set oMenu = poVBE.CommandBars.ActiveMenuBar.FindControl(id:=30003)    '30003 = Edit Menu
    Set oBar = oMenu.CommandBar

    Set oCtl = oBar.FindControl(id:=14)          '14 = Outdent

    If oCtl Is Nothing Then
        Set oMenu = oBar.Controls.Add(Type:=msoControlPopup, temporary:=True)
        oMenu.BeginGroup = True
    Else
        Set oMenu = oBar.Controls.Add(Type:=msoControlPopup, before:=oCtl.Index + 1, temporary:=True)
    End If

    CreateSubMenus oMenu

End Sub

'
' Adds an 'Indent' option to the Project explorer menu(s)
'
Sub AddProjectExplorerMenu(ByVal sBar As String)

    Dim oBar As CommandBar
    Dim oEvt As CBtnEvents, oBtn As CommandBarButton

    On Error Resume Next
    Set oBar = poVBE.CommandBars(sBar)

    Set oBtn = oBar.Controls.Add(Type:=msoControlButton, Parameter:="Indent2KPrjWin", temporary:=True)
    With oBtn
        .Caption = msMENU_PROJ_WIN
        .FaceId = 2564
        .Style = msoButtonIconAndCaption
        .Tag = psMENU_TAG & Rnd
        .BeginGroup = True
    End With
    Set oEvt = New CBtnEvents
    Set oEvt.oHook = poVBE.Events.CommandBarEvents(oBtn)
    poBtnEvents.Add oEvt

End Sub


'
' Adds the Proc, Mod, Proj and Options items to our popup menus
'
Sub CreateSubMenus(oMenu As CommandBarPopup)
Attribute CreateSubMenus.VB_UserMemId = 1610612737
    Dim oEvt As CBtnEvents, oBtn As CommandBarButton

    With oMenu
        .Caption = msMENU_TEXT
        .Parameter = "Indent2K"
        .Tag = psMENU_TAG
    End With

    With oMenu.CommandBar.Controls
        'Indent current procedure
        Set oBtn = .Add(Type:=msoControlButton, Parameter:="Indent2KProc", temporary:=True)
        With oBtn
            .Caption = msMENU_PROC
            .FaceId = 2564
            .Style = msoButtonIconAndCaption
            .Tag = psMENU_TAG
        End With
        Set oEvt = New CBtnEvents
        Set oEvt.oHook = poVBE.Events.CommandBarEvents(oBtn)
        poBtnEvents.Add oEvt

        'Indent current module
        Set oBtn = .Add(Type:=msoControlButton, Parameter:="Indent2KMod", temporary:=True)
        With oBtn
            .Caption = msMENU_MOD
            .FaceId = 472
            .Style = msoButtonIconAndCaption
            .Tag = psMENU_TAG
        End With
        Set oEvt = New CBtnEvents
        Set oEvt.oHook = poVBE.Events.CommandBarEvents(oBtn)
        poBtnEvents.Add oEvt

        'Indent project
        Set oBtn = .Add(Type:=msoControlButton, Parameter:="Indent2KProj", temporary:=True)
        With oBtn
            .Caption = msMENU_PROJ
            .FaceId = 2557
            .Style = msoButtonIconAndCaption
            .Tag = psMENU_TAG
        End With
        Set oEvt = New CBtnEvents
        Set oEvt.oHook = poVBE.Events.CommandBarEvents(oBtn)
        poBtnEvents.Add oEvt

        'Undo
        Set oBtn = .Add(Type:=msoControlButton, Parameter:="Indent2KUndo", temporary:=True)
        With oBtn
            .Caption = msMENU_UNDO
            .FaceId = 128
            .Style = msoButtonIconAndCaption
            .Tag = psMENU_TAG
            .Enabled = False
        End With

        poBtnUndo.Add oBtn

        Set oEvt = New CBtnEvents
        Set oEvt.oHook = poVBE.Events.CommandBarEvents(oBtn)
        poBtnEvents.Add oEvt

        'Show indenting options
        Set oBtn = .Add(Type:=msoControlButton, Parameter:="Indent2KForm", temporary:=True)
        With oBtn
            .Caption = msMENU_FORM
            .FaceId = 222
            .Style = msoButtonIconAndCaption
            .Tag = psMENU_TAG
        End With
        Set oEvt = New CBtnEvents
        Set oEvt.oHook = poVBE.Events.CommandBarEvents(oBtn)
        poBtnEvents.Add oEvt
    End With

End Sub

'
' Delete our menu items
'
Sub RemoveMenus()
Attribute RemoveMenus.VB_UserMemId = 1610612738

'    Dim oCtl As CommandBarControl
'
'    On Error Resume Next
'
'    Set oCtl = poVBE.CommandBars.FindControl(Tag:=psMENU_TAG)
'    Do Until oCtl Is Nothing
'        oCtl.Delete
'        Set oCtl = poVBE.CommandBars.FindControl(Tag:=psMENU_TAG)
'    Loop

End Sub

