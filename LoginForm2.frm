VERSION 5.00
Begin VB.Form LoginForm2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   Icon            =   "LoginForm2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LoginForm2.frx":164A
   ScaleHeight     =   2145
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   615
      Left            =   2040
      Picture         =   "LoginForm2.frx":6F59
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   480
      Picture         =   "LoginForm2.frx":7170
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   3480
      Picture         =   "LoginForm2.frx":7368
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2040
      MaxLength       =   47
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   2040
      MaxLength       =   47
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image imgVbar 
      Height          =   10770
      Left            =   0
      Picture         =   "LoginForm2.frx":7524
      Top             =   0
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Shape shptxt 
      BorderColor     =   &H8000000D&
      Height          =   435
      Index           =   1
      Left            =   2160
      Top             =   720
      Width           =   1695
   End
   Begin VB.Shape shptxt 
      BorderColor     =   &H8000000D&
      Height          =   435
      Index           =   0
      Left            =   2160
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "LoginForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oUsrSrv As New CUsersServices

Private Enum itxt
  User = 0
  pass = 1
End Enum

Private Sub cmdBatal_Click()
   End
End Sub

Private Sub cmdHelp_Click()
  MsgBox "User: super" & vbCrLf & "Password: super"
End Sub

Private Sub cmdOK_Click()
 If oUsrSrv.GetUserValidation(txt(itxt.User), txt(itxt.pass)) Then
       'sUserName = txt(itxt.User).Text
       Unload Me
       'MsgBox "oUsrSrv.GroupName = " & oUsrSrv.GroupName
       sUserGroupName = oUsrSrv.GroupName
       MainFrm.LoadMenus
       
     Else
        MsgBox "Name Or Password Error, try again!", vbCritical, "Login"
    End If
    
    Set oUsrSrv = Nothing
End Sub

Private Sub Form_Load()
 shptxt(itxt.User).Width = txt(itxt.User).Width + 30
 shptxt(itxt.User).Height = txt(itxt.User).Height + 30
 shptxt(itxt.User).Top = txt(itxt.User).Top - 15
 shptxt(itxt.User).Left = txt(itxt.User).Left - 15
 
 shptxt(itxt.pass).Width = txt(itxt.pass).Width + 30
 shptxt(itxt.pass).Height = txt(itxt.pass).Height + 30
 shptxt(itxt.pass).Top = txt(itxt.pass).Top - 15
 shptxt(itxt.pass).Left = txt(itxt.pass).Left - 15
 
 Me.Top = (Screen.Height - Me.Height) / 2
 Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = itxt.User Then
    If KeyAscii = 13 Then
      txt(itxt.pass).SetFocus
    End If
  End If
  If Index = itxt.pass Then
    If KeyAscii = 13 Then
      cmdOK.SetFocus
    End If
  End If
End Sub
