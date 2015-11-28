Option Explicit On
Option Strict On

Imports Microsoft.Vbe.Interop

Public Interface IIndenter
    Sub IndentProcedure()
    Sub IndentModule()
    Sub IndentProject()
    Sub IndentFromProjectWindow()
    Sub UndoIndenting()
End Interface

Public Class Indenter
    
End Class
