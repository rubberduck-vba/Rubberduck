Option Explicit On

Imports System.Windows.Forms;

Public Class frmVbeOnKey
    Private Sub Form_Closing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        UnHookAll
    End Sub
End Class