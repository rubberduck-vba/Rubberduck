VERSION 5.00
Begin VB.Form frmVBEOnKey 
   Caption         =   "VBEOnKey-ID"
   ClientHeight    =   45
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   45
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmVBEOnKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Terminate()
    UnHookAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHookAll
End Sub
