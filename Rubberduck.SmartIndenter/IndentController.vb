Option Explicit On
Option Strict On

Imports Microsoft.Vbe.Interop

Public Interface IIndentController
    Sub IndentProcedure(codeModule As CodeModule, procedureName As String, startLine As Long, endLine As Long, Optional ByRef progress As Long? = Nothing)
    Sub RebuildCodeArray(asCodeLines() As String, procedureName As String, Optional ByRef progress As Long = 0, Optional showProgress As Boolean = False)
End Interface

Public Class IndentController : Implements IIndentController

    Private mbNoIndent As Boolean
    Private mbInIf As Boolean

    Private settings As ISmartIndenterSettings

    Public Sub New(settings As ISmartIndenterSettings)
        Me.settings = settings
    End Sub


    Public Sub IndentProcedure(codeModule As CodeModule, 
                               procedureName As String, 
                               startLine As Long, 
                               endLine As Long, 
                               Optional ByRef progress As Long? = Nothing) _
    Implements IIndentController.IndentProcedure
        
        'Variables used for the indenting code
        Dim X As Long, i As Long, j As Long, k As Long, iGap As Long, iLineAdjust As Long
        Dim lLineCount As Long, iCommentStart As Long, iStart As Long, iScan As Long, iDebugAdjust As Long
        Dim iIndents As Long, iIndentNext As Long, iIn As Long, iOut As Long
        Dim iFunctionStart As Long, iParamStart As Long
        Dim bInCmt As Boolean, bProcStart As Boolean, bAlign As Boolean, bFirstCont As Boolean
        Dim bAlreadyPadded As Boolean, bFirstDim As Boolean
        Dim sLine As String, sLeft As String, sRight As String, sMatch As String, sItem As String
        Dim vaScope As Object, vaStatic As Object, vaType As Object, vaInProc As Object
        Dim iCodeLineNum As Long, sCodeLineNum As String, sOrigLine As String

        On Error Resume Next

        mbNoIndent = False
        mbInIf = False

        Dim bShowProgress as Boolean = progress.HasValue

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

    Public Sub RebuildCodeArray(asCodeLines() As String, procedureName As String, Optional ByRef progress As Long = 0, Optional showProgress As Boolean = False) Implements IIndentController.RebuildCodeArray

    End Sub

End Class
