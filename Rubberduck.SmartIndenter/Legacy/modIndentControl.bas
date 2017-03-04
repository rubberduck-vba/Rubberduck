Attribute VB_Name = "modIndentControl"
'***************************************************************************
'*
'* PROJECT NAME:    SMART INDENTER
'* AUTHOR & DATE:   STEPHEN BULLEN, Office Automation Ltd.
'*                  14 July 1999
'*
'*                  COPYRIGHT © 1999-2004 BY OFFICE AUTOMATION LTD
'*
'* CONTACT:         stephen@oaltd.co.uk
'* WEB SITE:        http://www.oaltd.co.uk
'*
'* DESCRIPTION:     Adds items to the VBE environment to recreate the indenting
'*                  for the current procedure, module or project.
'*
'* THIS MODULE:     Identifies where and what to indent.
'*
'* PROCEDURES:
'*   IndentProcedure  Handles the "Current Procedure" menu item
'*   IndentModule     Handles the "Current Module" menu item
'*   IndentProject    Handles the "Current Project" menu item
'*   UndoIndenting    Undoes the last indenting
'*   fnHasCodeModule  Identifies if a given VBComponent has a code module to indent
'*   fnHasCode        Checks to see if a given code module contains any code
'*   CheckActivePane  Stores and restores the selections for the active code pane
'*   IsOfficeVBE      Checks if the VBE we're working in is the Office 2000 VBE
'*   SetParentToVBE   Sets the userform's parent as the VBE main window
'*   GetTabWidth      Reads the user's Tab Width setting from the registry
'*   fnPackComment    Removes the gap between code and an EOL comment
'*
'***************************************************************************
'*
'* CHANGE HISTORY
'*
'*  DATE        NAME                DESCRIPTION
'*  14/07/1999  Stephen Bullen      Initial version
'*  14/04/2000  Stephen Bullen      Added SetParentToVBE and Undo ability
'*  10/05/2000  Stephen Bullen      Added functions for shortcut keys
'*  28/05/2000  Stephen Bullen      Corrected Undo routine for line labels and EOL comments
'*  07/10/2004  Stephen Bullen      Changed to Office Automation
'*
'***************************************************************************

Option Explicit
Option Compare Text
Option Base 1

Dim mbCodeChanged As Boolean, msProc As String
Attribute mbCodeChanged.VB_VarUserMemId = 1073741824
Attribute msProc.VB_VarUserMemId = 1073741824
Dim miStartLine As Long, miEndline As Long, mlTotLines As Long
Attribute miStartLine.VB_VarUserMemId = 1073741826
Attribute miEndline.VB_VarUserMemId = 1073741826
Attribute mlTotLines.VB_VarUserMemId = 1073741826
Dim moVBP As VBProject, moVBC As VBComponent, moCM As CodeModule, moCP As CodePane
Attribute moVBP.VB_VarUserMemId = 1073741829
Attribute moVBC.VB_VarUserMemId = 1073741829
Attribute moCM.VB_VarUserMemId = 1073741829
Attribute moCP.VB_VarUserMemId = 1073741829

'
' Subroutine: IndentProcedure
'
' Purpose:    Locates the active procedure, checks it for code and indents it
'
Sub IndentProcedure()
Attribute IndentProcedure.VB_UserMemId = 1610612736

    Dim lProcType As Long
    Dim i As Integer, iTop As Integer

    On Error Resume Next

    'Don't do anything if not in a code window
    If poVBE.ActiveWindow.Type <> vbext_wt_CodeWindow Then Exit Sub

    'Try to get the active project
    Set moVBP = poVBE.ActiveVBProject

    'If we couldn't get it, display a message and quit
    If moVBP Is Nothing Then
        MsgBox "Could not identify current project.", vbOKOnly, psAPP_TITLE
        Exit Sub
    End If

    'If the project is locked, display a message and quit
    If IsOfficeVBE Then
        If moVBP.Protection = 1 Then
            MsgBox "The current project is locked and can't be indented.", vbOKOnly, psAPP_TITLE
            Exit Sub
        End If
    End If

    'Try to find the active code pane
    Set moCP = poVBE.ActiveCodePane

    'If we couldn't get it, display a message and quit
    If moCP Is Nothing Then
        MsgBox "Could not identify current module.", vbOKOnly, psAPP_TITLE
        Exit Sub
    End If

    'Check if the module contains any code
    If Not fnHasCode(moCP.CodeModule) Then
        MsgBox "The current module does not contain any code to indent.", vbOKOnly, psAPP_TITLE
        Exit Sub
    End If

    'Get where the current selection is in the module
    moCP.GetSelection miStartLine, 0, 0, 0

    Set moCM = moCP.CodeModule

    'Try to get the procedure name
    msProc = moCM.ProcOfLine(miStartLine, lProcType)

    If msProc <> "" Then
        'If we got a procedure name, find its start and end lines and quit the loop
        miStartLine = moCM.ProcStartLine(msProc, lProcType)
        miEndline = miStartLine + moCM.ProcCountLines(msProc, lProcType) - 1
    Else
        msProc = "Declarations"
        miStartLine = 1
        miEndline = moCM.CountOfDeclarationLines
    End If

    'Store the currently active pane
    CheckActivePane "Store"

    'Show the status bar user form.  The activate of the userform runs the indenting
    'routine, so it can update the status bar form as it progresses.
    With frmProgress
        .Min = 0
        .Max = miEndline - miStartLine
        .Action = "Procedure"
        .Show vbModal
    End With

    'Restore the currently active pane
    CheckActivePane "Restore"

End Sub

'
' Subroutine: IndentModule
'
' Purpose:    Locates the active module, checks it for code and indents it
'
Sub IndentModule()
Attribute IndentModule.VB_UserMemId = 1610612737

    On Error Resume Next

    'Try to get the active project
    Set moVBP = poVBE.ActiveVBProject

    'Don't do anything if not in a code window
    If poVBE.ActiveWindow.Type <> vbext_wt_CodeWindow Then Exit Sub

    'If we couldn't get it, display a message and quit
    If moVBP Is Nothing Then
        MsgBox "Could not identify current project.", vbOKOnly, psAPP_TITLE
        Exit Sub
    End If

    'If the project is locked, display a message and quit
    If IsOfficeVBE Then
        If moVBP.Protection = 1 Then
            MsgBox "The current project is locked and can't be indented.", vbOKOnly, psAPP_TITLE
            Exit Sub
        End If
    End If

    'Try to find the active code pane
    Set moCP = poVBE.ActiveCodePane

    'If we couldn't get it, display a message and quit
    If moCP Is Nothing Then
        MsgBox "Could not identify current module.", vbOKOnly, psAPP_TITLE
        Exit Sub
    End If

    'Check if the module contains any code
    If Not fnHasCode(moCP.CodeModule) Then
        MsgBox "The current module does not contain any code to indent.", vbOKOnly, psAPP_TITLE
        Exit Sub
    End If

    'Store the currently active pane
    CheckActivePane "Store"

    Set moCM = moCP.CodeModule

    'Show the status bar user form.  The activate of the userform runs the indenting
    'routine, so it can update the status bar form as it progresses.
    With frmProgress
        .Min = 0
        .Max = moCM.CountOfLines - 1
        .Action = "Module"
        .Show vbModal
    End With

    'Restore the currently active pane
    CheckActivePane "Restore"

End Sub

'
' Subroutine: IndentProject
'
' Purpose:    Locates the active project, checks it for code and indents it
'
Sub IndentProject()
Attribute IndentProject.VB_UserMemId = 1610612738

    Dim bSomeCode As Boolean
    Dim lLineCount As Long, lLinesDone As Long

    On Error Resume Next

    'Don't do anything if not in a code window
    If poVBE.ActiveWindow.Type <> vbext_wt_CodeWindow Then Exit Sub

    'Try to get the active project
    Set moVBP = poVBE.ActiveVBProject

    'If we couldn't get it, display a message and quit
    If moVBP Is Nothing Then
        MsgBox "Could not identify current project.", vbOKOnly, psAPP_TITLE
        Exit Sub
    End If

    'If the project is locked, display a message and quit
    If IsOfficeVBE Then
        If moVBP.Protection = 1 Then
            MsgBox "The current project is locked and can't be indented.", vbOKOnly, psAPP_TITLE
            Exit Sub
        End If
    End If

    'Check if the project contains any code, by looping through its VBComponents
    bSomeCode = False
    lLineCount = 0

    For Each moVBC In moVBP.VBComponents
        If fnHasCodeModule(moVBC) Then
            If fnHasCode(moVBC.CodeModule) Then
                bSomeCode = True
                lLineCount = lLineCount + moVBC.CodeModule.CountOfLines
            End If
        End If
    Next

    'If we didn't find any code, display a message and quit
    If Not bSomeCode Then
        MsgBox "The current project, '" & moVBP.Name & "', does not contain any code", vbOKOnly, psAPP_TITLE
        Exit Sub
    End If

    'Store the currently active pane
    CheckActivePane "Store"

    'Show the status bar user form.  The activate of the userform runs the indenting
    'routine, so it can update the status bar form as it progresses.
    With frmProgress
        .Min = 0
        .Max = lLineCount - 1
        .Action = "Project"
        .Show vbModal
    End With

    'Restore the currently active pane
    CheckActivePane "Restore"

End Sub


'
' Subroutine: IndentFromProjectWindow
'
' Purpose:    Locates the active project, checks it for code and indents it
'
Sub IndentFromProjectWindow()

    Dim bSomeCode As Boolean
    Dim lLineCount As Long, lLinesDone As Long

    On Error Resume Next

    'Try to get the active project
    Set moVBP = poVBE.ActiveVBProject

    'If we couldn't get it, display a message and quit
    If moVBP Is Nothing Then
        MsgBox "Could not identify current project.", vbOKOnly, psAPP_TITLE
        Exit Sub
    End If

    'If the project is locked, display a message and quit
    If IsOfficeVBE Then
        If moVBP.Protection = 1 Then
            MsgBox "The current project is locked and can't be indented.", vbOKOnly, psAPP_TITLE
            Exit Sub
        End If
    End If

    If poVBE.SelectedVBComponent Is Nothing Then
        'Clicked on a project node, so indent the whole thing
        Set moVBP = poVBE.ActiveVBProject

        'Check if the project contains any code, by looping through its VBComponents
        bSomeCode = False
        lLineCount = 0

        For Each moVBC In moVBP.VBComponents
            If fnHasCodeModule(moVBC) Then
                If fnHasCode(moVBC.CodeModule) Then
                    bSomeCode = True
                    lLineCount = lLineCount + moVBC.CodeModule.CountOfLines
                End If
            End If
        Next

        'If we didn't find any code, display a message and quit
        If Not bSomeCode Then
            MsgBox "The current project, '" & moVBP.Name & "', does not contain any code", vbOKOnly, psAPP_TITLE
            Exit Sub
        End If

        'Store the currently active pane
        CheckActivePane "Store"

        'Show the status bar user form.  The activate of the userform runs the indenting
        'routine, so it can update the status bar form as it progresses.
        With frmProgress
            .Min = 0
            .Max = lLineCount - 1
            .Action = "Project"
            .Show vbModal
        End With

        'Restore the currently active pane
        CheckActivePane "Restore"
    Else
        'Clicked on a VBComponent node, so indent that

        'Check if the module contains any code
        If Not fnHasCode(poVBE.SelectedVBComponent.CodeModule) Then
            MsgBox "The selected module does not contain any code to indent.", vbOKOnly, psAPP_TITLE
            Exit Sub
        End If

        'Store the currently active pane
        CheckActivePane "Store"

        Set moCM = poVBE.SelectedVBComponent.CodeModule

        'Show the status bar user form.  The activate of the userform runs the indenting
        'routine, so it can update the status bar form as it progresses.
        With frmProgress
            .Min = 0
            .Max = moCM.CountOfLines - 1
            .Action = "Module"
            .Show vbModal
        End With

        'Restore the currently active pane
        CheckActivePane "Restore"
    End If

End Sub

'
' Undo the last indenting
'
Sub UndoIndenting()
Attribute UndoIndenting.VB_UserMemId = 1610612739

    Dim iMod As Integer
    Dim oCtls As CommandBarControls, oCtl As CommandBarControl

    On Error Resume Next

    'Don't do anything if not in a code window
    If poVBE.ActiveWindow.Type <> vbext_wt_CodeWindow Then Exit Sub

    mbCodeChanged = False

    'Get the total number of lines we indented before
    mlTotLines = 0
    For iMod = 1 To piUndoCount
        With pauUndo(iMod)
            If Not .oMod Is Nothing Then mlTotLines = mlTotLines + .lEndLine - .lStartLine + 1
        End With
    Next

    With frmProgress
        .Min = 0
        .Max = mlTotLines
        .Progress = 0
        .Action = "Undo"
        .Show vbModal
    End With

    'Had any of the code changed?
    If mbCodeChanged Then
        MsgBox "Some of the code has been changed since doing the indenting." & vbCrLf & _
               "The indenting was only undone for those lines that have not changed.", vbOKOnly, psAPP_TITLE
    End If

    'Clear out the old undo information
    Erase pauUndo

    'Disable all of our Undo buttons
    For Each oCtl In poBtnUndo
        oCtl.Enabled = False
    Next

End Sub

'
' Perform the action required within the progress bar display
'
Sub DoAction(ByVal sAction As String)
Attribute DoAction.VB_UserMemId = 1610612740

    Dim iMod As Integer, iLine As Long, lProgress As Long
    Dim sCurr As String, sIndented As String, i As Long
    Dim colMemInfo As New Collection, oMemInfo As CMemberInfo, oCM As Object, iMembers As Long

    On Error Resume Next

    Select Case sAction
    Case "Procedure"
        'Just rebuilding a procedure, so pass the procedure name and line boundaries
        RebuildModule moCM, msProc, miStartLine, miEndline, 0

    Case "Module"
        'Rebuilding a module, so pass the module name and number of lines therein
        RebuildModule moCM, moCM.Parent.Name, 1, moCM.CountOfLines, 0

    Case "Project"
        'Loop through the components to rebuild their indenting
        For Each moVBC In moVBP.VBComponents
            If fnHasCodeModule(moVBC) Then
                If fnHasCode(moVBC.CodeModule) Then
                    Set moCM = moVBC.CodeModule
                    'Pass the module name, number of lines, how many have been done already and how many are there in total
                    RebuildModule moCM, moCM.Parent.Name, 1, moCM.CountOfLines, lProgress

                    'Increment the number of lines done for next time round
                    lProgress = lProgress + moCM.CountOfLines
                End If
            End If
        Next

    Case "Undo"
        mbCodeChanged = False

        'Loop through, putting back the original code, ensuring we don't overwrite changed code
        For iMod = 1 To piUndoCount
            With pauUndo(iMod)
                If Not .oMod Is Nothing Then
                    frmProgress.MessageText = "Undoing '" & .sName & "'"

                    'Store all the procedure attributes in the module
                    Set oCM = .oMod
                    iMembers = oCM.Members.Count

                    If iMembers > 0 Then
                        For i = 1 To iMembers
                            If oCM.Members(i).CodeLocation >= .lStartLine And oCM.Members(i).CodeLocation <= .lEndLine Then
                                Set oMemInfo = New CMemberInfo
                                CopyMemberInfo oCM.Members(i), oMemInfo
                                colMemInfo.Add oMemInfo, CStr(oCM.Members(i).CodeLocation)
                            End If
                        Next
                    End If

                    For iLine = 0 To .lEndLine - .lStartLine

                        'Update the progress indicator
                        If (lProgress + iLine) Mod Int(mlTotLines / 200 + 1) = 0 Then
                            frmProgress.Progress = (lProgress + iLine)
                        End If

                        'Did we change the line during our indenting?
                        If .asIndented(iLine) <> .asOriginal(iLine) Then

                            sCurr = Trim$(.oMod.Lines(.lStartLine + iLine, 1))
                            sIndented = Trim$(.asIndented(iLine))

                            'Has the line been changed since we did our indenting?
                            If sCurr = sIndented Then

                                'Only if nothing's changed can we put the original code back
                                .oMod.ReplaceLine .lStartLine + iLine, .asOriginal(iLine)

                            ElseIf fnPackComment(sCurr) = fnPackComment(sIndented) Then

                                'We can undo if the only change is the position of the comment
                                .oMod.ReplaceLine .lStartLine + iLine, .asOriginal(iLine)

                            Else
                                mbCodeChanged = True
                            End If
                        End If

                        'Write back our members' properties
                        If colMemInfo.Count > 0 Then
                            For i = 1 To iMembers
                                CopyMemberInfo colMemInfo(CStr(oCM.Members(i).CodeLocation)), oCM.Members(i)
                            Next
                        End If
                    Next

                    lProgress = lProgress + .lEndLine - .lStartLine + 1
                End If
            End With
        Next

    End Select

End Sub


'
' Identify if a given VBComponent has a Code Module to indent
'
Private Function fnHasCodeModule(moVBC As VBComponent) As Boolean

    Dim i As Long

    On Error Resume Next

    Select Case moVBC.Type
    Case 7, 10, 4   'vbext_ct_PropPage = 7, vbext_ct_RelatedDocument = 10, vbext_ct_ResFile = 4
        'These don't have code modules

    Case Else

        Err.Clear
        i = moVBC.CodeModule.CountOfLines

        fnHasCodeModule = (Err.Number = 0)
        Err.Clear
    End Select

End Function

'
' Checks if the given code module actually contains any code.
'
Private Function fnHasCode(moCM As CodeModule) As Boolean

    Dim i As Long

    For i = 1 To moCM.CountOfLines
        If Trim$(moCM.Lines(i, 1)) <> "" Then
            fnHasCode = True
            Exit Function
        End If
    Next

End Function


'
' Stores and restores the visible lines and selection in the active code pane.
' When we are replacing code lines, the topline of the code module and the
' current selection can become reset.
'
Sub CheckActivePane(sType As String)
Attribute CheckActivePane.VB_UserMemId = 1610612743

    'Remember these values for when we come to restore the settings
    Static iStartLine As Long, iStartCol As Long, iEndline As Long, iEndCol As Long, iTop As Long
    Dim oCtls As CommandBarControls, oCtl As CommandBarControl, bCanUndo As Boolean

    On Error Resume Next

    'Using the active code pane
    With poVBE.ActiveCodePane

        'Store or restore the topline and selection settings as appropriate
        If sType = "Store" Then
            iTop = .TopLine
            .GetSelection iStartLine, iStartCol, iEndline, iEndCol

            'Clear the undo array
            Erase pauUndo
            piUndoCount = 0
        Else
            .TopLine = iTop
            .SetSelection iStartLine, iStartCol, iEndline, iEndCol

            'Are we storing Undo information?
            bCanUndo = (GetSetting(psREG_SECTION, psREG_KEY, "EnableUndo", "Y") = "Y")

            'Enable all of our Undo buttons
            For Each oCtl In poBtnUndo
                oCtl.Enabled = bCanUndo
            Next
        End If
    End With

End Sub

'
' Identifies if the IDE we're working in is an Office IDE, or a VB6 IDE
'
Function IsOfficeVBE()
Attribute IsOfficeVBE.VB_UserMemId = 1610612744

    Dim s As String
    On Error Resume Next

    s = poVBE.FullName

    IsOfficeVBE = (Err.Number <> 0)
    Err.Clear

End Function


'
' Read the user's tab width from the registry
'
Function GetTabWidth() As Integer
Attribute GetTabWidth.VB_UserMemId = 1610612745

    Dim sKey As String, i As Integer

    On Error Resume Next

    'Are we in the Office or VB IDE?
    If IsOfficeVBE Then
        If Left(poVBE.Version, 1) = "6" Then
            'Excel 97
            sKey = "\Software\Microsoft\VBA\6.0\Common\"
        Else
            'Office 2000
            sKey = "\Software\Microsoft\VBA\7.0\Common\"
        End If
    Else
        'VB
        sKey = "\Software\Microsoft\VBA\Microsoft Visual Basic\"
    End If

    'Default to 4 spaces
    i = Val("&H" & Mid(funGetRegValue(HKEY_CURRENT_USER, sKey, "TabWidth", "0x000004"), 3))
    If i = 0 Then i = 4

    GetTabWidth = i

End Function

'
' Removes the gap between code and an end-of-line comment
'
Function fnPackComment(ByVal sLine As String) As String
Attribute fnPackComment.VB_UserMemId = 1610612746

    Dim iPos As Long, iString As Long, iCmt As Long, iRem As Long

    fnPackComment = sLine

    iPos = 1

    Do
        'Find the positions of the first string, comment or Rem
        iString = InStr(iPos, sLine, """", 0)
        iCmt = InStr(iPos, sLine, "'", 0)
        iRem = InStr(iPos, sLine, " Rem ", 0)

        'If no comments or Rem, quit the function
        If iCmt = 0 And iRem = 0 Then Exit Function

        'Get the position of the first comment or Rem
        If iRem > 0 And iRem < iCmt Then iCmt = iRem

        If iString > 0 And iString < iCmt Then
            'If a string comes before a comment, skip to after the comment and try again
            iPos = InStr(iString + 1, sLine, """", 0) + 1
        Else
            'Comment came first, so strip out any space between the end of the line and the comment
            fnPackComment = Trim$(Left$(sLine, iCmt - 1)) & " " & Trim$(Mid$(sLine, iCmt))
            Exit Do
        End If
    Loop

End Function

'
' Copy a Procedure's attributes between a Member object and our
' CMemberInfo class
'
Sub CopyMemberInfo(ByRef CopyFrom As Object, ByRef CopyTo As Object)
Attribute CopyMemberInfo.VB_UserMemId = 1610612747

    'In case the properties are read-only or not applicable.
    On Error Resume Next

    With CopyTo
        If .Bindable <> CopyFrom.Bindable Then .Bindable = CopyFrom.Bindable
        If .Browsable <> CopyFrom.Browsable Then .Browsable = CopyFrom.Browsable
        If .Category <> CopyFrom.Category Then .Category = CopyFrom.Category
        If .DefaultBind <> CopyFrom.DefaultBind Then .DefaultBind = CopyFrom.DefaultBind
        If .Description <> CopyFrom.Description Then .Description = CopyFrom.Description
        If .DisplayBind <> CopyFrom.DisplayBind Then .DisplayBind = CopyFrom.DisplayBind
        If .HelpContextID <> CopyFrom.HelpContextID Then .HelpContextID = CopyFrom.HelpContextID
        If .Hidden <> CopyFrom.Hidden Then .Hidden = CopyFrom.Hidden
        If .PropertyPage <> CopyFrom.PropertyPage Then .PropertyPage = CopyFrom.PropertyPage
        If .RequestEdit <> CopyFrom.RequestEdit Then .RequestEdit = CopyFrom.RequestEdit
        If .StandardMethod <> CopyFrom.StandardMethod Then .StandardMethod = CopyFrom.StandardMethod
        If .UIDefault <> CopyFrom.UIDefault Then .UIDefault = CopyFrom.UIDefault
    End With

End Sub

'
' Makes a UserForm a child of the VBE window
'
Sub SetParentToVBE(oForm As Object)
Attribute SetParentToVBE.VB_UserMemId = 1610612748

    Dim hWndForm As Long

    On Error Resume Next

    hWndForm = oForm.hWnd

    If hWndForm <> 0 Then
        'SetParent hWndForm, poVBE.MainWindow.hwnd
    End If

End Sub

