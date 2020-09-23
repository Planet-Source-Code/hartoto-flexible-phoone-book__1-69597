VERSION 5.00
Begin VB.Form EditFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1845
   ClientLeft      =   1020
   ClientTop       =   2400
   ClientWidth     =   4080
   Icon            =   "Editfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4080
   Begin XBSuperData.Edit Edit1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3201
   End
End
Attribute VB_Name = "EditFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_oCaller As Object

Public sCari As String
Public sKeyFd As String
Public sXMLFile As String
'
Private Sub Edit1_onCancel()
  Unload Me
End Sub

Private Sub Edit1_onDataRefresh()
  For Each ctl In oCaller.Controls
    With ctl
        ctl.InitData
        ctl.EnhanceGrid
        Unload oCaller
        oCaller.Show
     End With
  Next
End Sub

Private Sub Form_Activate()
  If Edit1.bIsGetLookup Then
     Edit1.getLookup
  End If
  Me.Height = Edit1.Height + 450
  Me.Width = Edit1.Width
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
      Call Edit1_onCancel
   
  End Select
End Sub

Private Sub Form_Resize()
  Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Public Property Get oCaller() As Object
  Set oCaller = m_oCaller
End Property

Public Property Set oCaller(ByVal vNewValue As Object)
  Set m_oCaller = vNewValue
End Property



