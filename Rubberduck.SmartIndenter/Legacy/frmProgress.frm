VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indenting Progress"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fInside 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   150
      TabIndex        =   3
      Top             =   510
      Width           =   1950
      Begin VB.Label lblFront 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "50%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   3930
      End
   End
   Begin VB.Label lblBack 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "50%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   630
      Width           =   3930
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3990
   End
   Begin VB.Label lblMessage 
      Caption         =   "Indenting modAppMain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmProgress"
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
'* THIS MODULE:     Shows a (modeless) progress bar.
'*
'* PROCEDURES:
'*   UserForm_Initialize  Sets up the initial parameters
'*   Min                  Property to set/get the userform progress minimum value
'*   Max                  Property to set/get the userform progress maximum value
'*   Action               Property to set the action to perform when the form is shown
'*   UserForm_Activate    When the form is shown, calls a common routine to perform the action
'*   MesssageText         Property to set the message text
'*   Progress             Property to set the form's progress
'*
'***************************************************************************
'*
'* CHANGE HISTORY
'*
'*  DATE        NAME                DESCRIPTION
'*  14/07/1999  Stephen Bullen      Initial version
'*  17/04/2000  Stephen Bullen      Added call to set parent to the VBE
'*  04/05/2000  Stephen Bullen      Modified to use modal form - for Project 2000 and XL97
'*
'***************************************************************************

Option Explicit
Option Compare Text
Option Base 1

Dim mdMin As Double
Attribute mdMin.VB_VarUserMemId = 1073938432
Dim mdMax As Double
Attribute mdMax.VB_VarUserMemId = 1073938433
Dim msAction As String
Attribute msAction.VB_VarUserMemId = 1073938434

'
' Initialise the form to show a blank text and 0% complete
'
Private Sub Form_Load()

    On Error Resume Next

    MessageText = ""

    'Make us a child of the VBE window
    SetParentToVBE Me

End Sub

' Let the calling routine set/get the Minimum scale for the progress bar
Public Property Let Min(dNewMin As Double): mdMin = dNewMin: End Property
Public Property Get Min() As Double: Min = mdMin: End Property
Attribute Min.VB_UserMemId = 1745027076

' Let the calling routine set the Maximum scale for the progress bar
Public Property Let Max(dNewMax As Double): mdMax = dNewMax: End Property
Public Property Get Max() As Double: Max = mdMax: End Property
Attribute Max.VB_UserMemId = 1745027075

' Let the calling routine set the action to perform
Public Property Let Action(sNew As String): msAction = sNew: End Property
Public Property Get Action() As String: Action = msAction: End Property
Attribute Action.VB_UserMemId = 1745027074

'
' When the form is shown, call the actioning routine
'
Private Sub Form_Activate()

    Dim d As Double

    'Check max is greater than min
    If mdMin > mdMax Then
        d = mdMax
        mdMax = mdMin
        mdMax = d
    End If

    Progress = mdMin

    'Allow form to be displayed
    DoEvents

    'Perform the indenting action
    DoAction msAction

    'Unload the form
    Unload Me

End Sub

'
' Handle updating the userform's text
'
Property Let MessageText(sText As String)
Attribute MessageText.VB_UserMemId = 1745027073

    lblMessage.Caption = sText
    Me.Refresh

End Property

'
' Handle updating the userform progress bar
'
Property Let Progress(dAmt As Double)
Attribute Progress.VB_UserMemId = 1745027072

    Dim dPerc As Double

    If mdMax = mdMin Then
        dPerc = 0
    Else
        dPerc = Abs((dAmt - mdMin) / (mdMax - mdMin))
    End If

    'Set the wide of the inside frame, rouding to the nearest pixel
    fInside.Width = Int(lblBack.Width * dPerc / 0.75 + 1) * 0.75

    'Set the captions for the blue-on-white and white-on-blue texts.
    lblBack.Caption = Format(dPerc, "0%")
    lblFront.Caption = Format(dPerc, "0%")

    Me.Refresh

End Property

