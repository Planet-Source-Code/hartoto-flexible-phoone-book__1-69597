VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4320
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":164A
   ScaleHeight     =   2190
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1080
      Width           =   2085
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   2085
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   4290
      TabIndex        =   0
      Top             =   1740
      Width           =   4320
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1820
         TabIndex        =   2
         Tag             =   "OK"
         Top             =   80
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3050
         TabIndex        =   1
         Tag             =   "Cancel"
         Top             =   80
         Width           =   1200
      End
   End
   Begin VB.Shape shpBorder 
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Shape shpTop 
      BorderColor     =   &H00E0E0E0&
      DrawMode        =   2  'Blackness
      Height          =   1095
      Left            =   120
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Tag             =   "&Password:"
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Tag             =   "&User Name:"
      Top             =   720
      Width           =   1080
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   420
      Left            =   0
      Picture         =   "frmLogin.frx":97D6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11370
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oUsrSrv As New CUsersServices
Public OK As Boolean
'

Private Sub cmdCancel_Click()
    OK = False
    End
End Sub
Private Sub cmdOK_Click()
    
    If oUsrSrv.GetUserValidation(txtUserName, txtPassword) Then
       'MsgBox "OK"
       OK = True
       'frmMain.sUserName = txtUserName.Text
       sUserName = txtUserName.Text
       Unload Me
       'MsgBox "oUsrSrv.GroupName = " & oUsrSrv.GroupName
       sUserGroupName = oUsrSrv.GroupName
       frmMain.LoadMenus
       
     Else
        MsgBox "Nama atau Password salah, coba lagi!", vbCritical, "Login"
        'txtPassword.SelStart = 0
        'txtPassword.SelLength = Len(txtPassword.Text)
    End If
    
    Set oUsrSrv = Nothing
     
End Sub

Private Sub Form_Activate()
 txtUserName.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Me.OK = False Then
  Cancel = 1
End If
End Sub

Private Sub Form_Resize()
 shpBorder.Height = Me.Height
 shpBorder.Width = Me.Width
 shpBorder.Move 0, 0
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 cmdOK.SetFocus
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   txtPassword.SetFocus
 End If
End Sub
