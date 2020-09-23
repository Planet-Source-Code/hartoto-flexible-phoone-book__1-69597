VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form HelpFrm 
   Caption         =   "Help"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8505
   Icon            =   "HelpFrm.frx":0000
   Picture         =   "HelpFrm.frx":164A
   ScaleHeight     =   5685
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4695
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   7815
      ExtentX         =   13785
      ExtentY         =   8281
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Picture         =   "HelpFrm.frx":6F59
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image imgHelp 
      Height          =   225
      Left            =   120
      Picture         =   "HelpFrm.frx":7115
      Top             =   600
      Width           =   225
   End
   Begin VB.Image imgVBar 
      Height          =   10770
      Left            =   0
      Picture         =   "HelpFrm.frx":732C
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "HelpFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub Form_Resize()
cmdTutup.Left = Me.Width - cmdTutup.Width - 200
WebBrowser1.Width = Me.Width - imgVBar.Width - 200
WebBrowser1.Height = Me.Height - cmdTutup.Height - 200
WebBrowser1.Top = cmdTutup.Top + cmdTutup.Height + 100
WebBrowser1.Left = imgVBar.Width + 200
End Sub
