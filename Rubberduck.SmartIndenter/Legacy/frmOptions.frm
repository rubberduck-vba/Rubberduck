VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Smart Indenter Options"
   ClientHeight    =   7440
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11205
      Begin VB.CheckBox cbIndentCompilerStuff 
         Caption         =   "Indent compiler directives within procs"
         Height          =   255
         Left            =   7920
         TabIndex        =   30
         ToolTipText     =   "Whether to indent the code in a #If block inside a procedure."
         Top             =   2910
         Width           =   3015
      End
      Begin VB.CheckBox cbCompilerCol1 
         Caption         =   "Force compiler directives to column 1"
         Height          =   255
         Left            =   7920
         TabIndex        =   29
         ToolTipText     =   "Ensures that all compiler directives (e.g. #If) remain in column 1"
         Top             =   2610
         Width           =   3015
      End
      Begin VB.CheckBox cbAlignIgnoreOps 
         Caption         =   "Ignore operators when aligning"
         Height          =   255
         Left            =   8160
         TabIndex        =   28
         ToolTipText     =   "If a continued line starts with an operator, aligns that or the text following it."
         Top             =   1710
         Width           =   2655
      End
      Begin VB.CheckBox cbAlignDim 
         Caption         =   "Align Dim's in column:"
         Height          =   255
         Left            =   7920
         TabIndex        =   11
         ToolTipText     =   "Aligns the 'As' part of a Declaration line"
         Top             =   3210
         Width           =   1935
      End
      Begin VB.CheckBox cbEnableUndo 
         Caption         =   "Enable Undo"
         Height          =   255
         Left            =   7935
         TabIndex        =   14
         ToolTipText     =   "Allows the ability to undo the indenting, but slows down the routine."
         Top             =   3510
         Width           =   3120
      End
      Begin VB.CheckBox cbHotKeyMod 
         Caption         =   "Indent Module Hot-key:"
         Height          =   255
         Left            =   7935
         TabIndex        =   17
         ToolTipText     =   "Enable hot-key to indent the current module."
         Top             =   4110
         Width           =   2295
      End
      Begin VB.CheckBox cbHotKeyProc 
         Caption         =   "Indent Procedure Hot-key:"
         Height          =   255
         Left            =   7935
         TabIndex        =   15
         ToolTipText     =   "Enable hot-key to indent the current procedure."
         Top             =   3810
         Width           =   2265
      End
      Begin VB.TextBox ebHotKeyProc 
         Height          =   285
         Left            =   10305
         TabIndex        =   16
         Text            =   "+^P"
         Top             =   3780
         Width           =   720
      End
      Begin VB.TextBox ebHotKeyMod 
         Height          =   285
         Left            =   10305
         TabIndex        =   18
         Text            =   "+^M"
         Top             =   4080
         Width           =   720
      End
      Begin VB.Frame Frame2 
         Caption         =   "End-of-line Comment Handling"
         Height          =   1455
         Left            =   7935
         TabIndex        =   19
         Top             =   4410
         Width           =   3100
         Begin VB.VScrollBar spnEOLAlignCol 
            Height          =   240
            Left            =   2175
            Max             =   100
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1095
            Width           =   195
         End
         Begin VB.TextBox tbEOLAlignCol 
            Height          =   285
            Left            =   1680
            TabIndex        =   24
            Text            =   "123"
            ToolTipText     =   "The column number to align end-of-line comments to."
            Top             =   1065
            Width           =   720
         End
         Begin VB.OptionButton obEOLSameGap 
            Caption         =   "Maintain gap "
            Height          =   255
            Left            =   120
            TabIndex        =   21
            ToolTipText     =   "Maintains the same gap between the code and the end-of-line comment"
            Top             =   525
            Width           =   2535
         End
         Begin VB.OptionButton obEOLStdGap 
            Caption         =   "Force standard gap"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Forces a standard gap (two indent widths) between the code and the end-of-line column"
            Top             =   810
            Width           =   2535
         End
         Begin VB.OptionButton obEOLAlignCol 
            Caption         =   "Align in column:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Aligns all end-of-line comments in the same column"
            Top             =   1095
            Width           =   1455
         End
         Begin VB.OptionButton obEOLAbsolute 
            Caption         =   "Maintain absolute position"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            ToolTipText     =   "Keeps end-of-line comments in the same column as prior to indenting the code"
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.VScrollBar spnAlignDimCol 
         Height          =   240
         Left            =   10800
         Max             =   100
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3210
         Width           =   195
      End
      Begin VB.TextBox tbAlignDimCol 
         Height          =   285
         Left            =   10305
         TabIndex        =   12
         Text            =   "123"
         Top             =   3180
         Width           =   720
      End
      Begin VB.CheckBox cbIndentProc 
         Caption         =   "Indent everything within a procedure"
         Height          =   255
         Left            =   7920
         TabIndex        =   4
         ToolTipText     =   "Indents all the code within a procedure by one step."
         Top             =   210
         Width           =   3135
      End
      Begin VB.CheckBox cbIndentDim 
         Caption         =   "Indent first declaration block"
         Height          =   255
         Left            =   8160
         TabIndex        =   6
         ToolTipText     =   "Sets whether the first declaration block in a procedure is indented."
         Top             =   810
         Width           =   2655
      End
      Begin VB.CheckBox cbIndentCmt 
         Caption         =   "Indent comments to align with code"
         Height          =   255
         Left            =   7920
         TabIndex        =   7
         ToolTipText     =   "Indent all the comments to left-align with the code."
         Top             =   1110
         Width           =   3135
      End
      Begin VB.CheckBox cbAlignCont 
         Caption         =   "Align continued strings and parameters"
         Height          =   255
         Left            =   7920
         TabIndex        =   8
         ToolTipText     =   "In continued lines, ensures that strings and parameters line up"
         Top             =   1410
         Width           =   3135
      End
      Begin VB.CheckBox cbIndentCase 
         Caption         =   "Indent within ""Select Case"" lines"
         Height          =   255
         Left            =   7920
         TabIndex        =   9
         ToolTipText     =   "Indents everything with a Select Case statement."
         Top             =   2010
         Width           =   3015
      End
      Begin VB.CheckBox cbDebugCol1 
         Caption         =   "Force Stop and Debug to column 1"
         Height          =   255
         Left            =   7920
         TabIndex        =   10
         ToolTipText     =   "Ensures that all 'Stop' and 'Debug' lines remain in column 1"
         Top             =   2310
         Width           =   3015
      End
      Begin VB.CheckBox cbIndentFirst 
         Caption         =   "Indent first comment block"
         Height          =   255
         Left            =   8160
         TabIndex        =   5
         ToolTipText     =   "Sets whether the first comment block in a procedure is indented."
         Top             =   510
         Width           =   2415
      End
      Begin VB.ListBox lbCodeExample 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6360
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7575
      End
      Begin VB.Label lblURL 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.oaltd.co.uk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   9165
         MouseIcon       =   "frmOptions.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   6360
         Width           =   1980
      End
      Begin VB.Image imgOALogo 
         Height          =   540
         Left            =   8145
         MouseIcon       =   "frmOptions.frx":0152
         MousePointer    =   99  'Custom
         Picture         =   "frmOptions.frx":02A4
         Top             =   6000
         Width           =   900
      End
      Begin VB.Label lblOA 
         Caption         =   "© 1998-2004 by Office Automation Ltd"
         Height          =   405
         Left            =   9165
         TabIndex        =   26
         Top             =   5970
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10080
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   6960
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'* THIS MODULE:     Handles the Indenting Options userform, reading and writing the
'*                  options from/to the registry.
'*
'* PROCEDURES:
'*   UserForm_Initialize  Sets up the example procedure and get the options from the registry
'*   cbIndentProc_Click   Handle clicking on the "Indent all in procedure" check box
'*   cbIndentFirst_Click  Handle clicking on the "Indent first comment block" check box
'*   cbIndentDim_Click    Handle clicking on the "Indent declarations" check box
'*   cbIndentCmt_Click    Handle clicking on the "Indent comments" check box
'*   cbIndentCase_Click   Handle clicking on the "Indent Select Case" check box
'*   cbAlignCont_Click    Handle clicking on the "Align continued lines" check box
'*   cbAlignDim_Click     Handle clicking on the "Align Dim lines" check box
'*   cbDebugCol1_Click    Handle clicking on the "Force Debug and Stop to column 1" check box
'*   obEOLAbsolute_Click  Handle clicking on the "Keep comments in absolute position" option button
'*   obEOLSameGap_Click   Handle clicking on the "Keep same gap" option button
'*   obEOLStdGap_Click    Handle clicking on the "Force standard gap" option button
'*   obEOLAlignCol_Click  Handle clicking on the "Align in column" option button
'*   tbAlignDimCol_Change    Handle changing the Dim alignment text box
'*   spnAlignDimCol_Change   Handle changing the Dim alignment spinner
'*   tbEOLAlignCol_Change    Handle changing the EOL Comment alignment text box
'*   spnEOLAlignCol_Change   Handle changing the EOL Comment alignment spinner
'*   cmdOK_Click          Handle clicking on the OK button.  Stores the options in the registry
'*   cmdCancel_Click      Handle clicking on the Cancel button.  Restores the original options
'*   SaveOptions          Save the options to the registry
'*   UpdateCodeListBox    Common routine to work out the indenting in the example procedure
'*   UserForm_QueryClose  Handle the use of the 'x' to close the form - treat same as Cancel
'*
'***************************************************************************
'*
'* CHANGE HISTORY
'*
'*  DATE        NAME                DESCRIPTION
'*  14/07/1999  Stephen Bullen      Initial version
'*  14/04/2000  Stephen Bullen      Added more options
'*  03/05/2000  Stephen Bullen      Added option to not indent Dims
'*  10/05/2000  Stephen Bullen      Added options for procedure and module hot-keys
'*
'***************************************************************************

Option Explicit
Option Compare Text
Option Base 1

'API call to launch a web page in the default browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                                      ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
                                      ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim mbUpdating As Boolean
Attribute mbUpdating.VB_VarUserMemId = 1073938432

'
' Initialises the userform, by reading the options from the registry
'
Private Sub Form_Load()

    mbUpdating = False

    On Error Resume Next

    lblOA.Caption = "© 1998-2004 by" & vbLf & "Office Automation Ltd"

    cbIndentProc = Abs(CInt(GetSetting(psREG_SECTION, psREG_KEY, "IndentProc", "Y") = "Y"))
    cbIndentFirst = Abs(CInt(GetSetting(psREG_SECTION, psREG_KEY, "IndentFirst", "N") = "Y"))
    cbIndentDim = Abs(CInt(GetSetting(psREG_SECTION, psREG_KEY, "IndentDim", "Y") = "Y"))
    cbIndentCmt = Abs(CInt(GetSetting(psREG_SECTION, psREG_KEY, "IndentCmt", "Y") = "Y"))
    cbIndentCase = Abs(CInt(GetSetting(psREG_SECTION, psREG_KEY, "IndentCase", "N") = "Y"))
    cbAlignCont = Abs(CInt(GetSetting(psREG_SECTION, psREG_KEY, "AlignContinued", "Y") = "Y"))
    cbAlignIgnoreOps = Abs(GetSetting(psREG_SECTION, psREG_KEY, "AlignIgnoreOps", "N") = "Y")
    cbDebugCol1 = Abs(CInt(GetSetting(psREG_SECTION, psREG_KEY, "DebugCol1", "N") = "Y"))
    cbCompilerCol1 = Abs(GetSetting(psREG_SECTION, psREG_KEY, "CompilerCol1", "N") = "Y")
    cbIndentCompilerStuff = Abs(GetSetting(psREG_SECTION, psREG_KEY, "IndentCompiler", "Y") = "Y")
    cbAlignDim = Abs(CInt(GetSetting(psREG_SECTION, psREG_KEY, "AlignDim", "N") = "Y"))
    tbAlignDimCol = GetSetting(psREG_SECTION, psREG_KEY, "AlignDimCol", "15")
    cbEnableUndo = Abs(CInt(GetSetting(psREG_SECTION, psREG_KEY, "EnableUndo", "Y") = "Y"))
    cbHotKeyProc = Abs(CInt(GetSetting(psREG_SECTION, psREG_KEY, "EnableProcHotKey", "Y") = "Y"))
    ebHotKeyProc = GetSetting(psREG_SECTION, psREG_KEY, "ProcHotKeyCode", "+^P")
    cbHotKeyMod = Abs(CInt(GetSetting(psREG_SECTION, psREG_KEY, "EnableModHotKey", "Y") = "Y"))
    ebHotKeyMod = GetSetting(psREG_SECTION, psREG_KEY, "ModHotKeyCode", "+^M")

    ebHotKeyProc_Change
    ebHotKeyMod_Change
    cbHotKeyProc_Click
    cbHotKeyMod_Click

    Select Case GetSetting(psREG_SECTION, psREG_KEY, "EOLComments", "SameGap")
    Case "Absolute"
        obEOLAbsolute = True
    Case "SameGap"
        obEOLSameGap = True
    Case "StandardGap"
        obEOLStdGap = True
    Case "AlignInCol"
        obEOLAlignCol = True
    End Select

    tbEOLAlignCol = GetSetting(psREG_SECTION, psREG_KEY, "EOLAlignCol", "55")

    'Update the spinners
    tbAlignDimCol_Change
    tbEOLAlignCol_Change

    'Store inital values in the controls' Tag property
    cbIndentProc.Tag = cbIndentProc
    cbIndentFirst.Tag = cbIndentFirst
    cbIndentDim.Tag = cbIndentDim
    cbIndentCmt.Tag = cbIndentCmt
    cbIndentCase.Tag = cbIndentCase
    cbAlignCont.Tag = cbAlignCont
    cbAlignIgnoreOps.Tag = cbAlignIgnoreOps
    cbDebugCol1.Tag = cbDebugCol1
    cbCompilerCol1.Tag = cbCompilerCol1
    cbIndentCompilerStuff.Tag = cbIndentCompilerStuff
    cbAlignDim.Tag = cbAlignDim
    tbAlignDimCol.Tag = tbAlignDimCol
    cbEnableUndo.Tag = cbEnableUndo
    cbHotKeyProc.Tag = cbHotKeyProc
    ebHotKeyProc.Tag = ebHotKeyProc
    cbHotKeyMod.Tag = cbHotKeyMod
    ebHotKeyMod.Tag = ebHotKeyMod
    obEOLAbsolute.Tag = obEOLAbsolute
    obEOLSameGap.Tag = obEOLSameGap
    obEOLStdGap.Tag = obEOLStdGap
    obEOLAlignCol.Tag = obEOLAlignCol
    tbEOLAlignCol.Tag = tbEOLAlignCol

    mbUpdating = True

    'Update the code box
    UpdateCodeListBox

    'Make us a child of the VBE window
    SetParentToVBE Me

End Sub

'When changing the check boxes, update the example indenting
Private Sub cbIndentProc_Click(): UpdateCodeListBox: End Sub
Private Sub cbIndentFirst_Click(): UpdateCodeListBox: End Sub
Private Sub cbIndentDim_Click(): UpdateCodeListBox: End Sub
Private Sub cbIndentCmt_Click(): UpdateCodeListBox: End Sub
Private Sub cbIndentCase_Click(): UpdateCodeListBox: End Sub
Private Sub cbAlignDim_Click(): UpdateCodeListBox: End Sub
Private Sub cbAlignIgnoreOps_Click(): UpdateCodeListBox: End Sub
Private Sub cbDebugCol1_Click(): UpdateCodeListBox: End Sub
Private Sub cbIndentCompilerStuff_Click(): UpdateCodeListBox: End Sub
Private Sub obEOLAbsolute_Click(): UpdateCodeListBox: End Sub
Private Sub obEOLSameGap_Click(): UpdateCodeListBox: End Sub
Private Sub obEOLStdGap_Click(): UpdateCodeListBox: End Sub
Private Sub obEOLAlignCol_Click(): UpdateCodeListBox: End Sub

'
'   When changing the 'Align continued strings and parameters',
'   enable/disable the 'Ignore operators' check box
'
Private Sub cbAlignCont_Click()

    If Not mbUpdating Then Exit Sub

    cbAlignIgnoreOps.Enabled = (cbAlignCont.Value = 0)

    If cbAlignIgnoreOps.Enabled Then
        cbAlignIgnoreOps.Value = cbAlignIgnoreOps.Tag
    Else
        cbAlignIgnoreOps.Value = False
    End If

    UpdateCodeListBox

End Sub

'
'   When changing the 'Force compiler directives to col1',
'   enable/disable the 'Indent within compiler' check box
'
Private Sub cbCompilerCol1_Click()

    If Not mbUpdating Then Exit Sub

    cbIndentCompilerStuff.Enabled = (cbCompilerCol1.Value = 0)

    If cbIndentCompilerStuff.Enabled Then
        cbIndentCompilerStuff.Value = cbIndentCompilerStuff.Tag
    Else
        cbIndentCompilerStuff.Value = False
    End If

    UpdateCodeListBox

End Sub

Private Sub Form_Activate()
    SetParentToVBE Me
End Sub

Private Sub imgOALogo_Click()
    lblURL_Click
End Sub

'
'   When changing the Alignment column text box, validate and update the spinner
'
Private Sub tbAlignDimCol_Change()

    'Validate the item in the edit box
    If tbAlignDimCol = "" Then
        spnAlignDimCol = 100
    ElseIf Not IsNumeric(tbAlignDimCol) Then
        tbAlignDimCol = ""
        spnAlignDimCol = 100
    Else
        tbAlignDimCol = CStr(Abs(CInt(tbAlignDimCol)))

        'Only allow entries between 0 and 100
        If CInt(tbAlignDimCol) < 0 Or CInt(tbAlignDimCol) > 100 Then
            tbAlignDimCol = ""
            spnAlignDimCol = 100
        Else
            spnAlignDimCol.Value = 100 - CInt(tbAlignDimCol)
        End If
    End If

    UpdateCodeListBox

End Sub

'
'   When changing the Alignment column spinner, update the text box
'
Private Sub spnAlignDimCol_Change()
    tbAlignDimCol.Text = CStr(100 - spnAlignDimCol.Value)
End Sub

Private Sub cbHotKeyProc_Click()
    ebHotKeyProc.Enabled = cbHotKeyProc
End Sub

Private Sub ebHotKeyProc_Change()
    ebHotKeyProc.Text = UCase(ebHotKeyProc.Text)
    ebHotKeyProc.ToolTipText = fnGetTip(ebHotKeyProc.Text) & ".  See SendKeys for valid codes."
End Sub

Private Sub cbHotKeyMod_Click()
    ebHotKeyMod.Enabled = cbHotKeyMod
End Sub

Private Sub ebHotKeyMod_Change()
    ebHotKeyMod.Text = UCase(ebHotKeyMod.Text)
    ebHotKeyMod.ToolTipText = fnGetTip(ebHotKeyMod.Text) & ".  See SendKeys for valid codes."
End Sub

'
'   When changing the Alignment column text box, validate and update the spinner
'
Private Sub tbEOLAlignCol_Change()

    'Validate the item in the edit box
    If tbEOLAlignCol = "" Then
        spnEOLAlignCol = 100
    ElseIf Not IsNumeric(tbEOLAlignCol) Then
        tbEOLAlignCol = ""
        spnEOLAlignCol = 100
    Else
        tbEOLAlignCol = CStr(Abs(CInt(tbEOLAlignCol)))

        'Only allow entries between 0 and 100
        If CInt(tbEOLAlignCol) < 0 Or CInt(tbEOLAlignCol) > 100 Then
            tbEOLAlignCol = ""
            spnEOLAlignCol = 100
        Else
            spnEOLAlignCol.Value = 100 - CInt(tbEOLAlignCol)
        End If
    End If

    UpdateCodeListBox

End Sub

'
'   When changing the Alignment column spinner, update the text box
'
Private Sub spnEOLAlignCol_Change()
    tbEOLAlignCol.Text = CStr(100 - spnEOLAlignCol.Value)
End Sub

Private Sub lblURL_Click()
    'Launch the default browser to show the BMS web site.
    ShellExecute 0&, vbNullString, "www.oaltd.co.uk", vbNullString, vbNullString, vbNormalFocus
End Sub

'
' Store the current options in the registry and unload the form
'
Private Sub cmdOK_Click()

    Dim oCtls As CommandBarControls, oCtl As CommandBarControl

    'Save the current options to the registry
    SaveOptions

    'Unhook previous key combinations
    VBEOnKey ebHotKeyProc.Tag
    VBEOnKey ebHotKeyMod.Tag

    'Hook new ones
    If cbHotKeyProc Then VBEOnKey ebHotKeyProc, "Indent2KProc"
    If cbHotKeyMod Then VBEOnKey ebHotKeyMod, "Indent2KMod"

    'Save new key hot-key combination options
    SaveSetting psREG_SECTION, psREG_KEY, "EnableProcHotKey", IIf(cbHotKeyProc, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "ProcHotKeyCode", ebHotKeyProc
    SaveSetting psREG_SECTION, psREG_KEY, "EnableModHotKey", IIf(cbHotKeyMod, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "ModHotKeyCode", ebHotKeyMod

    'Have we turned off Undo?
    If Not cbEnableUndo Then

        'Clear out the old undo information
        Erase pauUndo

        'Disable all of our Undo buttons
        For Each oCtl In poBtnUndo
            oCtl.Enabled = False
        Next
    End If

    'Unload the userform
    Unload Me

End Sub

'
' Unload the form
'
Private Sub cmdCancel_Click()

    'Reset the registry entries to their original values (stored in their .Tag property)
    SaveSetting psREG_SECTION, psREG_KEY, "IndentProc", IIf(cbIndentProc.Tag, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "IndentFirst", IIf(cbIndentFirst.Tag, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "IndentDim", IIf(cbIndentDim.Tag, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "IndentCmt", IIf(cbIndentCmt.Tag, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "IndentCase", IIf(cbIndentCase.Tag, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "AlignContinued", IIf(cbAlignCont.Tag, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "AlignIgnoreOps", IIf(cbAlignIgnoreOps.Tag, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "DebugCol1", IIf(cbDebugCol1.Tag, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "CompilerCol1", IIf(cbCompilerCol1.Tag, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "IndentCompiler", IIf(cbIndentCompilerStuff.Tag, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "AlignDim", IIf(cbAlignDim.Tag, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "AlignDimCol", tbAlignDimCol.Tag
    SaveSetting psREG_SECTION, psREG_KEY, "EnableUndo", IIf(cbEnableUndo.Tag, "Y", "N")

    Select Case True
    Case obEOLAbsolute.Tag
        SaveSetting psREG_SECTION, psREG_KEY, "EOLComments", "Absolute"
    Case obEOLSameGap.Tag
        SaveSetting psREG_SECTION, psREG_KEY, "EOLComments", "SameGap"
    Case obEOLStdGap.Tag
        SaveSetting psREG_SECTION, psREG_KEY, "EOLComments", "StandardGap"
    Case obEOLAlignCol.Tag
        SaveSetting psREG_SECTION, psREG_KEY, "EOLComments", "AlignInCol"
    End Select

    SaveSetting psREG_SECTION, psREG_KEY, "EOLAlignCol", tbEOLAlignCol.Tag

    Unload Me
End Sub

'
' Save the form's options to the registry
'
Private Sub SaveOptions()

    If Not mbUpdating Then Exit Sub

    'Store the current options in the registry
    SaveSetting psREG_SECTION, psREG_KEY, "IndentProc", IIf(cbIndentProc, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "IndentFirst", IIf(cbIndentFirst, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "IndentDim", IIf(cbIndentDim, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "IndentCmt", IIf(cbIndentCmt, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "IndentCase", IIf(cbIndentCase, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "AlignContinued", IIf(cbAlignCont, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "AlignIgnoreOps", IIf(cbAlignIgnoreOps, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "DebugCol1", IIf(cbDebugCol1, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "CompilerCol1", IIf(cbCompilerCol1, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "IndentCompiler", IIf(cbIndentCompilerStuff, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "AlignDim", IIf(cbAlignDim, "Y", "N")
    SaveSetting psREG_SECTION, psREG_KEY, "AlignDimCol", tbAlignDimCol
    SaveSetting psREG_SECTION, psREG_KEY, "EnableUndo", IIf(cbEnableUndo, "Y", "N")

    Select Case True
    Case obEOLAbsolute
        SaveSetting psREG_SECTION, psREG_KEY, "EOLComments", "Absolute"
    Case obEOLSameGap
        SaveSetting psREG_SECTION, psREG_KEY, "EOLComments", "SameGap"
    Case obEOLStdGap
        SaveSetting psREG_SECTION, psREG_KEY, "EOLComments", "StandardGap"
    Case obEOLAlignCol
        SaveSetting psREG_SECTION, psREG_KEY, "EOLComments", "AlignInCol"
    End Select

    SaveSetting psREG_SECTION, psREG_KEY, "EOLAlignCol", tbEOLAlignCol

End Sub
'
' Routine to work out the indenting in the example code and put it in the list box.
'
Private Sub UpdateCodeListBox()

    Dim asCodeLines(1 To 30) As String, i As Integer

    If Not mbUpdating Then Exit Sub

    'Save the current options to the registry
    SaveOptions

    'Disable/enable certain options dependant on other
    cbIndentFirst.Enabled = cbIndentProc
    cbIndentDim.Enabled = cbIndentProc
    tbAlignDimCol.Enabled = cbAlignDim
    spnAlignDimCol.Enabled = cbAlignDim
    tbEOLAlignCol.Enabled = obEOLAlignCol
    spnEOLAlignCol.Enabled = obEOLAlignCol

    'Define the example procedure code lines
    asCodeLines(1) = "' Example Procedure"
    asCodeLines(2) = "Sub ExampleProc()"
    asCodeLines(3) = ""
    asCodeLines(4) = "'SMART INDENTER"
    asCodeLines(5) = "'© 1998-2004 by Office Automation Ltd."
    asCodeLines(6) = ""
    asCodeLines(7) = "Dim iCount As Integer"
    asCodeLines(8) = "Static sName As String"
    asCodeLines(9) = ""
    asCodeLines(10) = "If YouWantMoreExamplesAndTools Then"
    asCodeLines(11) = "' Visit http://www.oaltd.co.uk"
    asCodeLines(12) = ""
    asCodeLines(13) = "Select Case X"
    asCodeLines(14) = "Case ""A"""
    asCodeLines(15) = "' If you have any comments or suggestions, _"
    asCodeLines(16) = " or find valid VBA code that isn't indented correctly,"
    asCodeLines(17) = ""
    asCodeLines(18) = "#If VBA6 Then"
    asCodeLines(19) = "MsgBox ""Contact stephen@oaltd.co.uk"""
    asCodeLines(20) = "#End If"
    asCodeLines(21) = ""
    asCodeLines(22) = "Case ""Continued strings and parameters can be"" _"
    asCodeLines(23) = "& ""lined up for easier reading, optionally ignoring"" _"
    asCodeLines(24) = ", ""any operators (&+, etc) at the start of the line."""
    asCodeLines(25) = ""
    asCodeLines(26) = "Debug.Print ""X<>1"""
    asCodeLines(27) = "End Select           'Case X"
    asCodeLines(28) = "End If               'More Tools?"
    asCodeLines(29) = ""
    asCodeLines(30) = "End Sub"

    'Run the array through the indenting code
    RebuildCodeArray asCodeLines, "", 0, False

    'Put the procedure code in the list box.
    With lbCodeExample
        .Clear
        For i = LBound(asCodeLines) To UBound(asCodeLines)
            .AddItem asCodeLines(i)
        Next
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, CloseMode As Integer)

    'Treat same as cancelled
    If CloseMode = vbFormControlMenu Then
        cmdCancel_Click
    End If

End Sub

Function fnGetTip(ByVal sKey As String) As String
Attribute fnGetTip.VB_UserMemId = 1610809373

    Dim sString As String

    sKey = UCase(sKey)

    'Work out if Ctrl/Alt/Shift included
    If InStr(1, sKey, "^") > 0 Then sKey = Replace$(sKey, "^", ""): sString = sString & "Ctrl+"
    If InStr(1, sKey, "+") > 0 Then sKey = Replace$(sKey, "+", ""): sString = sString & "Shift+"
    If InStr(1, sKey, "%") > 0 Then sKey = Replace$(sKey, "%", ""): sString = sString & "Alt+"

    'Return the key's description
    If sKey = "" Then
        fnGetTip = "No Hot-key"

    ElseIf GetKeyCode(sKey) = 0 Then
        fnGetTip = "Unknown Hot-key"
    Else
        fnGetTip = sString & LCase$(sKey)
    End If

End Function
