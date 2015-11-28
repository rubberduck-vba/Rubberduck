Option Explicit On
Option Strict On

Public Interface ISmartIndenterSettings
    Property AlignContinued As Boolean
    Property AlignDim As Boolean
    Property AlignDimCol As Long
    Property AlignIgnoreOps As Boolean
    Property CompilerCol1 As Boolean
    Property DebugCol1 As Boolean
    Property EnableUndo As Boolean
    Property EOLAlignCol As Long
    Property EOLComments As String
    Property IndentCase As Boolean 
    Property IndentCmt As Boolean
    Property IndentCompiler As Boolean 
    Property IndentDim As Boolean
    Property IndentFirst As Boolean
    Property IndentProc As Boolean
End Interface