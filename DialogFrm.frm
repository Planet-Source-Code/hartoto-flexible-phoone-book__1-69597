VERSION 5.00
Begin VB.Form DialogFrm 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1995
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4200
   Icon            =   "DialogFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DialogFrm.frx":164A
   ScaleHeight     =   1995
   ScaleWidth      =   4200
   Begin VB.CommandButton cmdTidak 
      Caption         =   "&Tidak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2880
      Picture         =   "DialogFrm.frx":6F59
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1200
   End
   Begin VB.CommandButton cmdYa 
      Caption         =   "&Ya"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1560
      Picture         =   "DialogFrm.frx":7115
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   75
      Picture         =   "DialogFrm.frx":730D
      Top             =   480
      Width           =   240
   End
   Begin VB.Image imgVbar 
      Height          =   10770
      Left            =   0
      Picture         =   "DialogFrm.frx":7534
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   0
      Picture         =   "DialogFrm.frx":7EFC
      Top             =   0
      Width           =   11370
   End
   Begin VB.Label lblKomentar 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "DialogFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bYa As Boolean

Private Sub cmdTidak_Click()
 bYa = False
 Unload Me
End Sub

Private Sub cmdTidak_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
      cmdTidak_Click
   
  End Select
End Sub

Private Sub cmdYa_Click()
  bYa = True
  Unload Me
End Sub

Private Sub cmdYa_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
      cmdTidak_Click
   
  End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyEscape
      cmdTidak_Click
   
  End Select
End Sub

Private Sub Form_Resize()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

End Sub

