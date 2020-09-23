VERSION 5.00
Begin VB.Form frmDialog 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1575
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5685
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Dialog.frx":164A
   ScaleHeight     =   1575
   ScaleWidth      =   5685
   Begin VB.CommandButton cmdTidak 
      Caption         =   "&Tidak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   4320
      TabIndex        =   1
      Top             =   960
      Width           =   1200
   End
   Begin VB.CommandButton cmdYa 
      Caption         =   "&Ya"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   4320
      TabIndex        =   0
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   0
      Picture         =   "Dialog.frx":97D6
      Top             =   0
      Width           =   11370
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Dialog.frx":A85D
      Stretch         =   -1  'True
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblKomentar 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "frmDialog"
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

Private Sub cmdYa_Click()
  bYa = True
  Unload Me
End Sub

Private Sub Form_Resize()
Me.Left = (Screen.width - Me.width) / 2
Me.Top = (Screen.height - Me.height) / 2

End Sub
