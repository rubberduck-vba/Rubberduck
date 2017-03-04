VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Smart Indenter"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmAbout.frx":0000
   ScaleHeight     =   5160
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   4560
      Width           =   1575
   End
   Begin VB.PictureBox picMenu 
      AutoSize        =   -1  'True
      Height          =   2205
      Left            =   120
      Picture         =   "frmAbout.frx":0152
      ScaleHeight     =   2145
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   120
      MouseIcon       =   "frmAbout.frx":207F4
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":20946
      ScaleHeight     =   540
      ScaleWidth      =   900
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   1440
      MouseIcon       =   "frmAbout.frx":20D85
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label lblMore 
      Caption         =   "For more utilities and examples (primarily for Microsoft Excel), visit:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Label lblInstruct 
      Caption         =   $"frmAbout.frx":20ED7
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label lblCopyright 
      Caption         =   "© 1998-2005 by Office Automation Ltd"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label lblTitle 
      Caption         =   "Smart Indenter v3.5.2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    SetParentToVBE Me
End Sub

Private Sub btnOK_Click()

    On Error Resume Next

    Unload Me

End Sub

Private Sub lblURL_Click()

    On Error Resume Next

    ShellExecute 0&, vbNullString, "www.oaltd.co.uk", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub picLogo_Click()
    lblURL_Click
End Sub
