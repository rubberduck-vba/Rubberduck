Attribute VB_Name = "modIndentRoutine"
'***************************************************************************
'*
'* PROJECT NAME:    SMART INDENTER
'* AUTHOR:          STEPHEN BULLEN, Office Automation Ltd.
'*
'*                  COPYRIGHT © 1999-2004 BY OFFICE AUTOMATION LTD
'*
'* CONTACT:         stephen@oaltd.co.uk
'* WEB SITE:        http://www.oaltd.co.uk
'*
'* DESCRIPTION:     Adds items to the VBE environment to recreate the indenting
'*                  for the current procedure, module or project.
'*
'* THIS MODULE:     Contains the main procedure to rebuild the code's indenting
'*
'* PROCEDURES:
'*   RebuildModule      Copies code module into an array for rebuilding and back again
'*   RebuildCodeArray   Main routine to rebuild the indenting in the code
'*   fnFindFirstItem    Check whether a line of code contains any special treatment items
'*   CheckLine          Routine to identify whether to indent or outdent a line
'*   ArrayFromVariant   Convert a variant array to a string array, for faster processing
'*   fnAlignFunction    Find where to align to for continued lines
'*
'***************************************************************************
'*
'* CHANGE HISTORY
'*
'*  DATE        NAME                DESCRIPTION
'*  14/07/1999  Stephen Bullen      Initial version
'*  14/04/2000  Stephen Bullen      Improved algorithm, added options and split out module handling
'*  03/05/2000  Stephen Bullen      Added option to not indent Dims and handle line numbers
'*  24/05/2000  Stephen Bullen      Improved routine for aligning continued lines
'*  27/05/2000  Stephen Bullen      Fix comments with Type/Enum, Rem handling and brackets in strings
'*  04/07/2000  Stephen Bullen      Fix handling of aligned 'As' items and continued lines
'*  24/11/2000  Stephen Bullen      Added maintenance of Members' attributes for VB5 and 6
'*  07/10/2004  Stephen Bullen      Changed to Office Automation
'*  09/10/2004  Stephen Bullen      Bug fixes and more options
'*
'***************************************************************************

Option Explicit
Option Compare Binary
Option Base 1

Const miTAB As Integer = 9

'Variant arrays to hold the code items to look for
Dim masInProc() As String, masInCode() As String, masOutProc() As String, masOutCode() As String
Attribute masInCode.VB_VarUserMemId = 1073741824
Attribute masOutProc.VB_VarUserMemId = 1073741824
Attribute masOutCode.VB_VarUserMemId = 1073741824
Dim masDeclares() As String, masLookFor() As String, masFnAlign() As String
Attribute masDeclares.VB_VarUserMemId = 1073741828
Attribute masLookFor.VB_VarUserMemId = 1073741828
Attribute masFnAlign.VB_VarUserMemId = 1073741828

'Variables to hold our indenting options
Dim mbIndentProc As Boolean, mbIndentCmt As Boolean, mbIndentCase As Boolean, mbAlignCont As Boolean, mbIndentDim As Boolean
Attribute mbIndentProc.VB_VarUserMemId = 1073741831
Attribute mbIndentCmt.VB_VarUserMemId = 1073741831
Attribute mbIndentCase.VB_VarUserMemId = 1073741831
Attribute mbAlignCont.VB_VarUserMemId = 1073741831
Attribute mbIndentDim.VB_VarUserMemId = 1073741831
Dim mbIndentFirst As Boolean, mbAlignEOL As Boolean, mbAlignDim As Boolean, mbDebugCol1 As Boolean, mbEnableUndo As Boolean
Attribute mbIndentFirst.VB_VarUserMemId = 1073741836
Attribute mbAlignEOL.VB_VarUserMemId = 1073741836
Attribute mbAlignDim.VB_VarUserMemId = 1073741836
Attribute mbDebugCol1.VB_VarUserMemId = 1073741836
Attribute mbEnableUndo.VB_VarUserMemId = 1073741836
Dim miIndentSpaces As Integer, miEOLAlignCol As Integer, miAlignDimCol As Integer, mbCompilerStuffCol1 As Boolean
Attribute miIndentSpaces.VB_VarUserMemId = 1073741841
Attribute miEOLAlignCol.VB_VarUserMemId = 1073741841
Attribute miAlignDimCol.VB_VarUserMemId = 1073741841
Attribute mbCompilerStuffCol1.VB_VarUserMemId = 1073741841
Dim mbIndentCompilerStuff As Boolean, mbAlignIgnoreOps As Boolean
Attribute mbIndentCompilerStuff.VB_VarUserMemId = 1073741845
Attribute mbAlignIgnoreOps.VB_VarUserMemId = 1073741845

'Variables to hold operational information
Dim mbInitialised As Boolean, mbContinued As Boolean, mbInIf As Boolean, mbNoIndent As Boolean, mbFirstProcLine As Boolean
Attribute mbInitialised.VB_VarUserMemId = 1073741847
Attribute mbContinued.VB_VarUserMemId = 1073741847
Attribute mbInIf.VB_VarUserMemId = 1073741847
Attribute mbNoIndent.VB_VarUserMemId = 1073741847
Attribute mbFirstProcLine.VB_VarUserMemId = 1073741847
Dim miParamStart As Long
Attribute miParamStart.VB_VarUserMemId = 1073741852
Dim msEOLComment As String
Attribute msEOLComment.VB_VarUserMemId = 1073741853

''''''''''''''''''''''''''''''''''
' Function:   RebuildModule
'
' Comments:   This procedure goes through the lines in a module,
'             rebuilding the code's indenting.
'
' Arguments:  modCode    - The code module to indent
'             sName      - The display name of the item being indented
'             iStartLine - Value giving the line to start indenting from
'             iEndLine   - Value giving the line to end indenting at
'             iProgDone  - Value giving how much indenting has been done in total
'
Public Sub RebuildModule(modCode As CodeModule, sName As String, iStartLine As Long, iEndline As Long, iProgDone As Long)

    Dim asCode() As String, asOriginal() As String, i As Long, oCM As Object
    Dim colMemInfo As New Collection, oMemInfo As CMemberInfo, iMembers As Long

    On Error Resume Next

    ReDim asCode(0 To iEndline - iStartLine)
    ReDim asOriginal(0 To iEndline - iStartLine)

    frmProgress.MessageText = "Rebuilding '" & sName & "'"

    mbEnableUndo = (GetSetting(psREG_SECTION, psREG_KEY, "EnableUndo", "Y") = "Y")

    'Is Undo enabled?  If so, set up our storage
    If mbEnableUndo Then
        piUndoCount = piUndoCount + 1

        'Make some space in our undo array
        If piUndoCount = 1 Then
            ReDim pauUndo(1 To 1)
        Else
            ReDim Preserve pauUndo(1 To piUndoCount)
        End If

        'Store the undo information
        With pauUndo(piUndoCount)
            Set .oMod = modCode
            .sName = sName
            .lStartLine = iStartLine
            .lEndLine = iEndline
            ReDim .asIndented(0 To iEndline - iStartLine)
            ReDim .asOriginal(0 To iEndline - iStartLine)
        End With
    End If

    'Store all the procedure attributes in the module
    Set oCM = modCode
    iMembers = oCM.Members.Count

    If iMembers > 0 Then
        For i = 1 To iMembers
            If oCM.Members(i).CodeLocation >= iStartLine And oCM.Members(i).CodeLocation <= iEndline Then
                Set oMemInfo = New CMemberInfo
                CopyMemberInfo oCM.Members(i), oMemInfo
                colMemInfo.Add oMemInfo, CStr(oCM.Members(i).CodeLocation)
            End If
        Next
    End If

    'Read code module into an array and store the original code in our undo array
    For i = 0 To iEndline - iStartLine
        asCode(i) = modCode.Lines(iStartLine + i, 1)
        asOriginal(i) = asCode(i)

        If mbEnableUndo Then pauUndo(piUndoCount).asOriginal(i) = asCode(i)
    Next

    'Indent the array, showing the progress
    RebuildCodeArray asCode, sName, iProgDone, True

    'Copy the changed code back into the module and store in our undo array
    For i = 0 To iEndline - iStartLine

        If asOriginal(i) <> asCode(i) Then
            modCode.ReplaceLine iStartLine + i, asCode(i)
        End If

        If mbEnableUndo Then pauUndo(piUndoCount).asIndented(i) = asCode(i)
    Next

    'Write back our members' properties (for VB6)
    If colMemInfo.Count > 0 Then
        For i = 1 To iMembers
            CopyMemberInfo colMemInfo(CStr(oCM.Members(i).CodeLocation)), oCM.Members(i)
        Next
    End If

End Sub

Public Sub RebuildCodeArray(asCodeLines() As String, sName As String, iProgDone As Long, Optional bShowProgress As Variant)

    'Variables used for the indenting code
    Dim X As Integer, i As Integer, j As Integer, k As Integer, iGap As Integer, iLineAdjust As Integer
    Dim lLineCount As Long, iCommentStart As Long, iStart As Long, iScan As Long, iDebugAdjust As Integer
    Dim iIndents As Integer, iIndentNext As Integer, iIn As Integer, iOut As Integer
    Dim iFunctionStart As Long, iParamStart As Long
    Dim bInCmt As Boolean, bProcStart As Boolean, bAlign As Boolean, bFirstCont As Boolean
    Dim bAlreadyPadded As Boolean, bFirstDim As Boolean
    Dim sLine As String, sLeft As String, sRight As String, sMatch As String, sItem As String
    Dim vaScope As Variant, vaStatic As Variant, vaType As Variant, vaInProc As Variant
    Dim iCodeLineNum As Long, sCodeLineNum As String, sOrigLine As String

    On Error Resume Next

    mbNoIndent = False
    mbInIf = False

    If IsMissing(bShowProgress) Then bShowProgress = True

    'Read the indenting options from the registry
    miIndentSpaces = GetTabWidth                 'Read VB's own setting for tab width
    mbIndentProc = (GetSetting(psREG_SECTION, psREG_KEY, "IndentProc", "Y") = "Y")
    mbIndentFirst = (GetSetting(psREG_SECTION, psREG_KEY, "IndentFirst", "N") = "Y")
    mbIndentDim = (GetSetting(psREG_SECTION, psREG_KEY, "IndentDim", "Y") = "Y")
    mbIndentCmt = (GetSetting(psREG_SECTION, psREG_KEY, "IndentCmt", "Y") = "Y")
    mbIndentCase = (GetSetting(psREG_SECTION, psREG_KEY, "IndentCase", "N") = "Y")
    mbAlignCont = (GetSetting(psREG_SECTION, psREG_KEY, "AlignContinued", "Y") = "Y")
    mbAlignIgnoreOps = (GetSetting(psREG_SECTION, psREG_KEY, "AlignIgnoreOps", "Y") = "Y")
    mbDebugCol1 = (GetSetting(psREG_SECTION, psREG_KEY, "DebugCol1", "N") = "Y")
    mbAlignDim = (GetSetting(psREG_SECTION, psREG_KEY, "AlignDim", "N") = "Y")
    miAlignDimCol = Val(GetSetting(psREG_SECTION, psREG_KEY, "AlignDimCol", "15"))

    If mbCompilerStuffCol1 <> (GetSetting(psREG_SECTION, psREG_KEY, "CompilerCol1", "N") = "Y") Or _
       mbIndentCompilerStuff <> (GetSetting(psREG_SECTION, psREG_KEY, "IndentCompiler", "Y") = "Y") Then
        mbInitialised = False
    End If

    mbCompilerStuffCol1 = (GetSetting(psREG_SECTION, psREG_KEY, "CompilerCol1", "N") = "Y")
    mbIndentCompilerStuff = (GetSetting(psREG_SECTION, psREG_KEY, "IndentCompiler", "Y") = "Y")

    msEOLComment = GetSetting(psREG_SECTION, psREG_KEY, "EOLComments", "SameGap")
    miEOLAlignCol = GetSetting(psREG_SECTION, psREG_KEY, "EOLAlignCol", "55")

    ' Create the list of items to match for the indenting at procedure level
    If Not mbInitialised Then
        vaScope = Array("", "Public ", "Private ", "Friend ")
        vaStatic = Array("", "Static ")
        vaType = Array("Sub", "Function", "Property Let", "Property Get", "Property Set", "Type", "Enum")

        X = 1
        ReDim vaInProc(1)
        For i = 1 To UBound(vaScope)
            For j = 1 To UBound(vaStatic)
                For k = 1 To UBound(vaType)
                    ReDim Preserve vaInProc(X)
                    vaInProc(X) = vaScope(i) & vaStatic(j) & vaType(k)
                    X = X + 1
                Next
            Next
        Next

        ArrayFromVariant masInProc, vaInProc

        'Items to match when outdenting at procedure level
        ArrayFromVariant masOutProc, Array("End Sub", "End Function", "End Property", "End Type", "End Enum")

        If mbIndentCompilerStuff Then
            'Items to match when indenting within a procedure
            ArrayFromVariant masInCode, Array("If", "ElseIf", "Else", "#If", "#ElseIf", "#Else", "Select Case", "Case", "With", "For", "Do", "While")

            'Items to match when outdenting within a procedure
            ArrayFromVariant masOutCode, Array("ElseIf", "Else", "End If", "#ElseIf", "#Else", "#End If", "Case", "End Select", "End With", "Next", "Loop", "Wend")
        Else
            'Items to match when indenting within a procedure
            ArrayFromVariant masInCode, Array("If", "ElseIf", "Else", "Select Case", "Case", "With", "For", "Do", "While")

            'Items to match when outdenting within a procedure
            ArrayFromVariant masOutCode, Array("ElseIf", "Else", "End If", "Case", "End Select", "End With", "Next", "Loop", "Wend")
        End If

        'Items to match for declarations
        ArrayFromVariant masDeclares, Array("Dim", "Const", "Static", "Public", "Private", "#Const")

        'Things to look for within a line of code for special handling
        ArrayFromVariant masLookFor, Array("""", ": ", " As ", "'", "Rem ", "Stop ", "Debug.Print ", "Debug.Assert ", "#If ", "#ElseIf ", "#Else ", "#End If ", "#Const ")

        mbInitialised = True
    End If

    'Things to skip when finding the function start of a line
    ArrayFromVariant masFnAlign, Array("Set ", "Let ", "LSet ", "RSet ", "Declare Function", "Declare Sub", "Private Declare Function", "Private Declare Sub", "Public Declare Function", "Public Declare Sub")

    If masInCode(UBound(masInCode)) <> "Select Case" And mbIndentCase Then
        'If extra-indenting within Select Case, ensure that we have two items in the arrays
        ReDim Preserve masInCode(UBound(masInCode) + 1)
        masInCode(UBound(masInCode)) = "Select Case"

        ReDim Preserve masOutCode(UBound(masOutCode) + 1)
        masOutCode(UBound(masOutCode)) = "End Select"

    ElseIf masInCode(UBound(masInCode)) = "Select Case" And Not mbIndentCase Then
        'If not extra-indenting within Select Case, ensure that we have one item in the arrays
        ReDim Preserve masInCode(UBound(masInCode) - 1)
        ReDim Preserve masOutCode(UBound(masOutCode) - 1)
    End If

    'Flag if the lines are at the top of a procedure
    bProcStart = False
    bFirstDim = False
    bFirstCont = True

    'Loop through all the lines to indent
    For lLineCount = LBound(asCodeLines) To UBound(asCodeLines)
        iLineAdjust = 0
        bAlreadyPadded = False
        iCodeLineNum = -1

        sOrigLine = asCodeLines(lLineCount)

        'At each 0.5% interval, update the progress bar itself
        If bShowProgress Then
            If (iProgDone + lLineCount - LBound(asCodeLines)) Mod Int(frmProgress.Max / 200 + 1) = 0 Then
                frmProgress.Progress = iProgDone + lLineCount - LBound(asCodeLines)
            End If
        End If

        'Read the line of code to indent
        sLine = Trim$(asCodeLines(lLineCount))

        'If we're not in a continued line, initialise some variables
        If Not (mbContinued Or bInCmt) Then
            mbFirstProcLine = False
            iIndentNext = 0
            iCommentStart = 0
            iIndents = iIndents + iDebugAdjust
            iDebugAdjust = 0
            iFunctionStart = 0
            iParamStart = 0

            i = InStr(1, sLine, " ")
            If i > 0 Then
                If IsNumeric(Left$(sLine, i - 1)) Then
                    iCodeLineNum = Val(Left$(sLine, i - 1))
                    sLine = Trim$(Mid$(sLine, i + 1))
                    sOrigLine = Space(i) & Mid(sOrigLine, i + 1)
                End If
            End If
        End If

        'Is there anything on the line?
        If Len(sLine) > 0 Then

            ' Remove leading Tabs
            Do Until Left$(sLine, 1) <> Chr$(miTAB)
                sLine = Mid$(sLine, 2)
            Loop

            ' Add an extra space on the end
            sLine = sLine & " "

            If bInCmt Then
                'Within a multi-line comment - indent to line up the comment text
                sLine = Space$(iCommentStart) & sLine

                'Remember if we're in a continued comment line
                bInCmt = Right$(Trim$(sLine), 2) = " _"
                GoTo PTR_REPLACE_LINE
            End If

            'Remember the position of the line segment
            iStart = 1
            iScan = 0

            If mbContinued And mbAlignCont Then
                If mbAlignIgnoreOps And Left$(sLine, 2) = ", " Then iParamStart = iFunctionStart - 2

                If mbAlignIgnoreOps And (Mid$(sLine, 2, 1) = " " Or Left$(sLine, 2) = ":=") And Left$(sLine, 2) <> ", " Then
                    sLine = Space$(iParamStart - 3) & sLine
                    iLineAdjust = iLineAdjust + iParamStart - 3
                    iScan = iScan + iParamStart - 3
                Else
                    sLine = Space$(iParamStart - 1) & sLine
                    iLineAdjust = iLineAdjust + iParamStart - 1
                    iScan = iScan + iParamStart - 1
                End If

                bAlreadyPadded = True
            End If

            'Scan through the line, character by character, checking for
            'strings, multi-statement lines and comments
            Do
                iScan = iScan + 1

                sItem = fnFindFirstItem(sLine, iScan)

                Select Case sItem
                Case ""
                    iScan = iScan + 1
                    'Nothing found => Skip the rest of the line
                    GoTo PTR_NEXT_PART

                Case """"
                    'Start of a string => Jump to the end of it
                    iScan = InStr(iScan + 1, sLine, """")
                    If iScan = 0 Then iScan = Len(sLine) + 1

                Case ": "
                    'A multi-statement line separator => Tidy up and continue

                    If Right$(Left$(sLine, iScan), 6) <> " Then:" Then
                        sLine = Left$(sLine, iScan + 1) & Trim$(Mid$(sLine, iScan + 2))

                        'And check the indenting for the line segment
                        CheckLine Mid$(sLine, iStart, iScan - 1), iIn, iOut, bProcStart
                        If bProcStart Then bFirstDim = True

                        If iStart = 1 Then
                            iIndents = iIndents - iOut
                            If iIndents < 0 Then iIndents = 0
                            iIndentNext = iIndentNext + iIn
                        Else
                            iIndentNext = iIndentNext + iIn - iOut
                        End If
                    End If

                    'Update the pointer and continue
                    iStart = iScan + 2

                Case " As "
                    'An " As " in a declaration => Line up to required column

                    If mbAlignDim Then
                        bAlign = mbNoIndent      'Don't need to check within Type
                        If Not bAlign Then
                            ' Check if we start with a declaration item
                            For i = LBound(masDeclares) To UBound(masDeclares)
                                sMatch = masDeclares(i) & " "
                                If Left$(sLine, Len(sMatch)) = sMatch Then
                                    bAlign = True
                                    Exit For
                                End If
                            Next
                        End If

                        If bAlign Then
                            i = InStr(iScan + 3, sLine, " As ")
                            If i = 0 Then
                                'OK to indent
                                If mbIndentProc And bFirstDim And Not mbIndentDim And Not mbNoIndent Then
                                    iGap = miAlignDimCol - Len(RTrim$(Left$(sLine, iScan)))

                                    'Adjust for a line number at the start of the line
                                    If iCodeLineNum > -1 Then iGap = iGap - Len(CStr(iCodeLineNum)) - 1
                                Else
                                    iGap = miAlignDimCol - Len(RTrim$(Left$(sLine, iScan))) - iIndents * miIndentSpaces

                                    'Adjust for a line number at the start of the line
                                    If iCodeLineNum > -1 Then
                                        If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
                                            iGap = iGap - (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) - 1
                                        End If
                                    End If
                                End If

                                If iGap < 1 Then iGap = 1
                            Else
                                'Multiple declarations on the line, so don't space out
                                iGap = 1
                            End If

                            'Work out the new spacing
                            sLeft = RTrim$(Left$(sLine, iScan))
                            sLine = sLeft & Space$(iGap) & Mid$(sLine, iScan + 1)

                            'Update the counters
                            iLineAdjust = iLineAdjust + iGap + Len(sLeft) - iScan
                            iScan = Len(sLeft) + iGap + 3
                        End If
                    Else
                        'Not aligning Dims, so remove any existing spacing
                        iScan = Len(RTrim$(Left$(sLine, iScan)))
                        sLine = RTrim$(Left$(sLine, iScan)) & " " & Trim$(Mid(sLine, iScan + 1))
                        iScan = iScan + 3
                    End If

                Case "'", "Rem "
                    'The start of a comment => Handle end-of-line comments properly

                    If iScan = 1 Then
                        'New comment at start of line
                        If bProcStart And Not mbIndentFirst And Not mbNoIndent Then
                            'No indenting

                        ElseIf mbIndentCmt Or bProcStart Or mbNoIndent Then
                            'Inside the procedure, so indent to align with code
                            sLine = Space$(iIndents * miIndentSpaces) & sLine
                            iCommentStart = iScan + iIndents * miIndentSpaces

                        ElseIf iIndents > 0 And mbIndentProc And Not bProcStart Then
                            'At the top of the procedure, so indent once if required
                            sLine = Space$(miIndentSpaces) & sLine
                            iCommentStart = iScan + miIndentSpaces
                        End If

                    Else
                        'New comment at the end of a line

                        'Make sure it's a proper 'Rem'
                        If sItem = "Rem " And Mid$(sLine, iScan - 1, 1) <> " " And Mid$(sLine, iScan - 1, 1) <> ":" Then GoTo PTR_NEXT_PART

                        'Check the indenting of the previous code segment
                        CheckLine Mid$(sLine, iStart, iScan - 1), iIn, iOut, bProcStart
                        If bProcStart Then bFirstDim = True

                        If iStart = 1 Then
                            iIndents = iIndents - iOut
                            If iIndents < 0 Then iIndents = 0
                            iIndentNext = iIndentNext + iIn
                        Else
                            iIndentNext = iIndentNext + iIn - iOut
                        End If

                        'Get the text before the comment, and the comment text
                        sLeft = Trim$(Left$(sLine, iScan - 1))
                        sRight = Trim$(Mid$(sLine, iScan))

                        'Indent the code part of the line
                        If bAlreadyPadded Then
                            sLine = RTrim$(Left$(sLine, iScan - 1))
                        Else
                            If mbContinued Then
                                sLine = Space$((iIndents + 2) * miIndentSpaces) & sLeft
                            Else
                                If mbIndentProc And bFirstDim And Not mbIndentDim Then
                                    sLine = sLeft
                                Else
                                    sLine = Space$(iIndents * miIndentSpaces) & sLeft
                                End If
                            End If
                        End If

                        mbContinued = (Right$(Trim$(sLine), 2) = " _")

                        'How do we handle end-of-line comments?
                        Select Case msEOLComment
                        Case "Absolute"
                            iScan = iScan - iLineAdjust + Len(sOrigLine) - Len(LTrim$(sOrigLine))
                            iGap = iScan - Len(sLine) - 1

                        Case "SameGap"
                            iScan = iScan - iLineAdjust + Len(sOrigLine) - Len(LTrim$(sOrigLine))
                            iGap = iScan - Len(RTrim$(Left$(sOrigLine, iScan - 1))) - 1

                        Case "StandardGap"
                            iGap = miIndentSpaces * 2

                        Case "AlignInCol"
                            iGap = miEOLAlignCol - Len(sLine) - 1
                        End Select

                        'Adjust for a line number at the start of the line
                        If iCodeLineNum > -1 Then
                            Select Case msEOLComment
                            Case "Absolute", "AlignInCol"
                                If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
                                    iGap = iGap - (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) - 1
                                End If
                            End Select
                        End If

                        If iGap < 2 Then iGap = miIndentSpaces

                        iCommentStart = Len(sLine) + iGap

                        'Put the comment in the required column
                        sLine = sLine & Space$(iGap) & sRight
                    End If

                    'Work out where the text of the comment starts, to align the next line
                    If Mid$(sLine, iCommentStart, 4) = "Rem " Then iCommentStart = iCommentStart + 3
                    If Mid$(sLine, iCommentStart, 1) = "'" Then iCommentStart = iCommentStart + 1
                    Do Until Mid$(sLine, iCommentStart, 1) <> " "
                        iCommentStart = iCommentStart + 1
                    Loop
                    iCommentStart = iCommentStart - 1

                    'Adjust for a line number at the start of the line
                    If iCodeLineNum > -1 Then
                        If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
                            iCommentStart = iCommentStart + (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) + 1
                        End If
                    End If

                    'Remember if we're in a continued comment line
                    bInCmt = Right$(Trim$(sLine), 2) = " _"

                    'Rest of line is comment, so no need to check any more
                    GoTo PTR_REPLACE_LINE

                Case "Stop ", "Debug.Print ", "Debug.Assert "

                    'A debugging statement - do we want to force to column 1?
                    If mbDebugCol1 And iStart = 1 And iScan = 1 Then
                        iLineAdjust = iLineAdjust - (Len(sOrigLine) - LTrim$(Len(sOrigLine)))
                        iDebugAdjust = iIndents
                        iIndents = 0
                    End If

                Case "#If ", "#ElseIf ", "#Else ", "#End If ", "#Const "
                    'Do we want to force compiler directives to column 1?
                    If mbCompilerStuffCol1 And iStart = 1 And iScan = 1 Then
                        iLineAdjust = iLineAdjust - (Len(sOrigLine) - LTrim$(Len(sOrigLine)))
                        iDebugAdjust = iIndents
                        iIndents = 0
                    End If
                End Select

PTR_NEXT_PART:
            Loop Until iScan > Len(sLine)        'Part of the line

            'Do we have some code left to check?
            '(i.e. a line without a comment or the last segment of a multi-statement line)
            If iStart < Len(sLine) Then

                If Not mbContinued Then bProcStart = False

                'Check the indenting of the remaining code segment
                CheckLine Mid$(sLine, iStart), iIn, iOut, bProcStart
                If bProcStart Then bFirstDim = True

                If iStart = 1 Then
                    iIndents = iIndents - iOut
                    If iIndents < 0 Then iIndents = 0
                    iIndentNext = iIndentNext + iIn
                Else
                    iIndentNext = iIndentNext + iIn - iOut
                End If
            End If

            'Start from the left at each procedure start
            If mbFirstProcLine Then iIndents = 0

            ' What about line continuations?  Here, I indent the continued line by
            ' two indents, and check for the end of the continuations.  Note
            ' that Excel won't allow comments in the middle of line continuations
            ' and that comments are treated differently above.
            If mbContinued Then
                If Not mbAlignCont Then
                    sLine = Space$((iIndents + 2) * miIndentSpaces) & sLine
                End If
            Else

                ' Check if we start with a declaration item
                bAlign = False
                If mbIndentProc And bFirstDim And Not mbIndentDim And Not bProcStart Then
                    For i = LBound(masDeclares) To UBound(masDeclares)
                        sMatch = masDeclares(i) & " "
                        If Left$(sLine, Len(sMatch)) = sMatch Then
                            bAlign = True
                            Exit For
                        End If
                    Next
                End If

                'Not a declaration item to left-align, so pad it out
                If Not bAlign Then
                    If Not bProcStart Then bFirstDim = False
                    sLine = Space$(iIndents * miIndentSpaces) & sLine
                End If

            End If

            mbContinued = (Right$(Trim$(sLine), 2) = " _")

        End If                                   'Anything there?

PTR_REPLACE_LINE:

        'Add the code line number back in
        If iCodeLineNum > -1 Then
            sCodeLineNum = CStr(iCodeLineNum)

            If Len(Trim$(Left$(sLine, Len(sCodeLineNum) + 1))) = 0 Then
                sLine = sCodeLineNum & Mid$(sLine, Len(sCodeLineNum) + 1)
            Else
                sLine = sCodeLineNum & " " & Trim$(sLine)
            End If
        End If

        asCodeLines(lLineCount) = RTrim$(sLine)

        'If it's not a continued line, update the indenting for the following lines
        If Not mbContinued Then
            iIndents = iIndents + iIndentNext
            iIndentNext = 0
            If iIndents < 0 Then iIndents = 0
        Else
            'A continued line, so if we're not in a comment and we want smart continuing,
            'work out which to continue from
            If mbAlignCont And Not bInCmt Then
                If Left$(Trim$(sLine), 2) = "& " Or Left$(Trim$(sLine), 2) = "+ " Then sLine = "  " & sLine

                iFunctionStart = fnAlignFunction(sLine, bFirstCont, iParamStart)
                If iFunctionStart = 0 Then
                    iFunctionStart = (iIndents + 2) * miIndentSpaces
                    iParamStart = iFunctionStart
                End If
            End If
        End If

        bFirstCont = Not mbContinued
    Next

End Sub

'
'  Find the first occurrence of one of our key items in the list
'
'    Returns the text of the item found
'    Updates the iFrom parameter to point to the location of the found item
'
Function fnFindFirstItem(ByRef sLine As String, ByRef iFrom As Long) As String

    Dim sItem As String, iFirst As Long, iFound As Long, iItem As Integer

    On Error Resume Next

    'Assume we don't find anything
    iFirst = Len(sLine)

    'Loop through the items to find within the line
    For iItem = LBound(masLookFor) To UBound(masLookFor)

        'What to find?
        sItem = masLookFor(iItem)

        'Is it there?
        iFound = InStr(iFrom, sLine, sItem)

        'Is it before any other items?
        If iFound > 0 And iFound < iFirst Then
            iFirst = iFound
            fnFindFirstItem = sItem
        End If
    Next

    'Update the location of the found item
    iFrom = iFirst

End Function


'
'  Check the line (segment) to see if it needs in- or out-denting
'
Function CheckLine(ByVal sLine As String, ByRef iIndentNext As Integer, ByRef iOutdentThis As Integer, ByRef bProcStart As Boolean)

    Dim i As Integer, j As Integer, sMatch As String

    On Error Resume Next

    'Assume we don't indent or outdent the code
    iIndentNext = 0
    iOutdentThis = 0

    'Tidy up the line
    sLine = Trim$(sLine) & " "

    'We don't check within Type and Enums
    If Not mbNoIndent Then

        ' Check for indenting within the code
        For i = LBound(masInCode) To UBound(masInCode)
            sMatch = masInCode(i)
            If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":")) Then
                iIndentNext = iIndentNext + 1
            End If
        Next

        ' Check for out-denting within the code
        For i = LBound(masOutCode) To UBound(masOutCode)
            sMatch = masOutCode(i)
            'Check at start of line for 'real' outdenting
            If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
                iOutdentThis = iOutdentThis + 1
            End If
        Next
    End If

    'Check procedure-level indenting
    For i = LBound(masInProc) To UBound(masInProc)
        sMatch = masInProc(i)
        If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then

            bProcStart = True
            mbFirstProcLine = True

            'Don't indent within Type or Enum constructs
            If Right$(sMatch, 4) = "Type" Or Right$(sMatch, 4) = "Enum" Then
                iIndentNext = iIndentNext + 1
                mbNoIndent = True

            ElseIf mbIndentProc And Not mbNoIndent Then
                iIndentNext = iIndentNext + 1
            End If
            Exit For
        End If
    Next

    'Check procedure-level outdenting
    For i = LBound(masOutProc) To UBound(masOutProc)
        sMatch = masOutProc(i)
        If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then

            'Don't indent within Type or Enum constructs
            If Right$(sMatch, 4) = "Type" Or Right$(sMatch, 4) = "Enum" Or mbIndentProc Then
                iOutdentThis = iOutdentThis + 1
                mbNoIndent = False
            End If
            Exit For
        End If
    Next

    'If we're not indenting, no need to consider the special cases
    If mbNoIndent Then Exit Function

    ' Treat If as a special case.  If anything other than a comment follows
    ' the Then, we don't indent
    If Left$(sLine, 3) = "If " Or Left$(sLine, 4) = "#If " Or mbInIf Then

        If mbInIf Then iIndentNext = 1

        'Strip any strings from the line
        i = InStr(1, sLine, """")
        Do Until i = 0
            j = InStr(i + 1, sLine, """")
            If j = 0 Then j = Len(sLine)
            sLine = Left$(sLine, i - 1) & Mid$(sLine, j + 1)
            i = InStr(1, sLine, """")
        Loop

        'And strip comments
        i = InStr(1, sLine, "'")
        If i > 0 Then sLine = Left$(sLine, i - 1)

        ' Do we have a Then statement in the line.  Adding a space on the
        ' end of the test means we can test for Then being both within or
        ' at the end of the line
        sLine = " " & sLine & " "
        i = InStr(1, sLine, " Then ")

        ' Allow for line continuations within the If statement
        mbInIf = (Right$(Trim$(sLine), 2) = " _")

        If i > 0 Then
            ' If there's something after the Then, we don't indent the If
            If Trim$(Mid$(sLine, i + 5)) <> "" Then iIndentNext = 0

            ' No need to check next time around
            mbInIf = False
        End If

        If mbInIf Then iIndentNext = 0
    End If

End Function

'
' Convert a Variant array to a string array for faster comparisons
'
Sub ArrayFromVariant(asString() As String, vaVariant As Variant)

    Dim iLow As Integer, iHigh As Integer, i As Integer

    On Error Resume Next

    iLow = LBound(vaVariant)
    iHigh = UBound(vaVariant)

    ReDim asString(iLow To iHigh)

    For i = iLow To iHigh
        asString(i) = vaVariant(i)
    Next

End Sub

'
' Locate the start of the first parameter on the line
'
Function fnAlignFunction(ByVal sLine As String, bFirstLine As Boolean, iParamStart As Long) As Long

    Dim iLPad As Integer, iCheck As Long, iBrackets As Long, iChar As Long, sMatch As String, iSpace As Integer
    Dim vAlign As Variant, bFound As Boolean, iAlign As Integer
    Dim iFirstThisLine As Integer

    Static coBrackets As Collection

    On Error Resume Next

    ReDim vAlign(1 To 2)

    If bFirstLine Then Set coBrackets = New Collection

    'Convert and numbers at the start of the line to spaces
    iChar = InStr(1, sLine, " ")
    If iChar > 1 Then
        If IsNumeric(Left$(sLine, iChar - 1)) Then
            sLine = Mid$(sLine, iChar + 1)
            iLPad = iChar
        End If
    End If

    iLPad = iLPad + Len(sLine) - Len(LTrim$(sLine))

    iFirstThisLine = coBrackets.Count

    sLine = Trim$(sLine)

    iCheck = 1

    'Skip over stuff that we don't want to locate the start off
    For iChar = LBound(masFnAlign) To UBound(masFnAlign)
        sMatch = masFnAlign(iChar)
        If Left$(sLine, Len(sMatch)) = sMatch Then
            iCheck = iCheck + Len(sMatch) + 1
            Exit For
        End If
    Next

    iBrackets = 0
    iSpace = 999
    For iChar = iCheck To Len(sLine)
        Select Case Mid$(sLine, iChar, 1)
        Case """"
            'A String => jump to the end of it
            iChar = InStr(iChar + 1, sLine, """")

        Case "("
            'Start of another function => remember this position
            vAlign(1) = "("
            vAlign(2) = iChar + iLPad
            coBrackets.Add vAlign

            vAlign(1) = ","
            vAlign(2) = iChar + iLPad + 1
            coBrackets.Add vAlign

        Case ")"
            'Function finished => Remove back to the previous open bracket
            vAlign = coBrackets(coBrackets.Count)
            Do Until vAlign(1) = "(" Or coBrackets.Count = iFirstThisLine
                coBrackets.Remove coBrackets.Count
                vAlign = coBrackets(coBrackets.Count)
            Loop
            If coBrackets.Count > iFirstThisLine Then coBrackets.Remove coBrackets.Count

        Case " "
            If Mid$(sLine, iChar, 3) = " = " Then
                'Space before an = sign => remember it to align to later
                bFound = False
                For iAlign = 1 To coBrackets.Count
                    vAlign = coBrackets(iAlign)
                    If vAlign(1) = "=" Or vAlign(1) = " " Then
                        bFound = True
                        Exit For
                    End If
                Next

                If Not bFound Then
                    vAlign(1) = "="
                    vAlign(2) = iChar + iLPad + 2
                    coBrackets.Add vAlign
                End If

            ElseIf coBrackets.Count = 0 And iChar < Len(sLine) - 2 Then
                'Space after a name before the end of the line => remember it for later
                vAlign(1) = " "
                vAlign(2) = iChar + iLPad
                coBrackets.Add vAlign

            ElseIf iChar > 5 Then
                'Clear the collection if we find a Then in an If...Then and set the
                'indenting to align with the bit after the "If "
                If Mid$(sLine, iChar - 5, 6) = " Then " Then
                    Do Until coBrackets.Count <= 1
                        coBrackets.Remove coBrackets.Count
                    Loop
                End If
            End If

        Case ","
            'Start of a new parameter => remember it to align to
            vAlign(1) = ","
            vAlign(2) = iChar + iLPad + 2
            coBrackets.Add vAlign

        Case ":"
            If Mid$(sLine, iChar, 2) = ":=" Then
                'A named paremeter => remember to align to after the name
                vAlign(1) = ","
                vAlign(2) = iChar + iLPad + 2
                coBrackets.Add vAlign

            ElseIf Mid$(sLine, iChar, 2) = ": " Then
                'A new line section, so clear the brackets
                Set coBrackets = New Collection
                iChar = iChar + 1
            End If
        End Select
    Next

    'If we end with a comma or a named parameter, get rid of all other comma alignments
    If Right$(Trim$(sLine), 3) = ", _" Or Right$(Trim$(sLine), 4) = ":= _" Then
        For iAlign = coBrackets.Count To 1 Step -1
            vAlign = coBrackets(iAlign)
            If vAlign(1) = "," Then
                coBrackets.Remove iAlign
            Else
                Exit For
            End If
        Next
    End If

    'If we end with a "( _", remove it and the space alignment after it
    If Right$(Trim$(sLine), 3) = "( _" Then
        coBrackets.Remove coBrackets.Count
        coBrackets.Remove coBrackets.Count
    End If

    iParamStart = 0

    'Get the position of the unmatched bracket and align to that
    For iAlign = 1 To coBrackets.Count
        vAlign = coBrackets(iAlign)
        If vAlign(1) = "," Then
            iParamStart = vAlign(2)
        ElseIf vAlign(1) = "(" Then
            iParamStart = vAlign(2) + 1
        Else
            iCheck = vAlign(2)
        End If
    Next

    If iCheck = 1 Or iCheck >= Len(sLine) + iLPad - 1 Then
        If coBrackets.Count = 0 And bFirstLine Then
            iCheck = miIndentSpaces * 2 + iLPad
        Else
            iCheck = iLPad
        End If
    End If

    If iParamStart = 0 Then iParamStart = iCheck + 1

    fnAlignFunction = iCheck + 1

End Function


'
' Look backwards through a string to find another string
'
Function RevInStr(sLookIn As String, sLookFor As String, iStart As Long) As Long

    Dim i As Long

    i = InStr(1, sLookIn, sLookFor)
    Do Until i > iStart Or i = 0
        RevInStr = i
        i = InStr(i + 1, sLookIn, sLookFor)
    Loop

End Function


