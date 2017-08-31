Attribute VB_Name = "ErrorUtilities"
Option Explicit

Public Sub ReraiseError(ByVal Err As ErrObject, _
    ByVal MethodName As String, _
    Optional ByVal Details As String _
)
    If Not IsMissing(Details) Then MethodName = MethodName & "(" & Details & ")"
    Err.Raise Err.Number, "    " & MethodName & vbNewLine & Err.Source
End Sub

Public Sub DisplayError(ByVal MyError As ErrObject, _
    ByVal MethodName As String, _
    Optional ByVal Details As String _
)
    If Not IsMissing(Details) Then MethodName = MethodName & "(" & Details & ")"
    MsgBox "Error #" & Err.Number & ": " & Err.Description & _
            vbNewLine & "From: " & _
            vbNewLine & "    " & MethodName & _
            vbNewLine & "    " & Err.Source, _
            vbOKOnly Or vbCritical, MethodName
End Sub
