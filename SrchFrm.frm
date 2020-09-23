VERSION 5.00
Begin VB.Form SrchFrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Jelajah"
   ClientHeight    =   6660
   ClientLeft      =   1470
   ClientTop       =   3240
   ClientWidth     =   7200
   Icon            =   "SrchFrm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   7200
   WindowState     =   2  'Maximized
   Begin XBSuperData.Srch Srch1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11880
   End
End
Attribute VB_Name = "SrchFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public sXMLFile As String

Public sCari As String
Public sKeyFd As String
Public sMasterFd As String
'
'

Private Sub Form_Resize()
 For Each ctl In Me.Controls
    With ctl
      .Width = Me.Width
      .Height = Me.Height
    End With
 Next
End Sub

Private Sub Srch1_addNew()

  EditMode = lEditMode.Add
  For Each ctl In Me.Controls
    With ctl
        With EditFrm
           Set .oCaller = Me
           .Caption = "Tambah"
           .Edit1.XMLFile = ctl.XMLFile
           .Edit1.KeyFd = ctl.KeyFd
           .Edit1.sCari = ctl.sCari
           .Edit1.InitData
           'MsgBox Srch1.KeyFd & " " & Srch1.sCari
           .Show vbModal
         End With
    End With
  Next
  
End Sub

Private Sub Srch1_editRec()
  
  EditMode = lEditMode.Edit
  For Each ctl In Me.Controls
    With ctl
        With EditFrm
          Set .oCaller = Me
          .Caption = "Edit"
          .Edit1.XMLFile = ctl.XMLFile
          .Edit1.KeyFd = ctl.KeyFd
          .Edit1.sCari = ctl.sCari
          .Edit1.InitData
          .Show vbModal
        End With
    End With
  Next
End Sub

Private Sub Srch1_grdDblClick()
  For Each ctl In Me.Controls
    With ctl
        frmEdit.sKeyFd = ctl.KeyFd
        frmEdit.sCari = ctl.sCari
        frmEdit.Show
      End With
  Next
  'MsgBox sCari
  
End Sub
Private Sub Srch1_onClose()
 Unload Me
End Sub

Private Sub Form_Activate()
 For Each ctl In Me.Controls
    With ctl
      .XMLFile = App.Path & "\" & sXMLFile
      .InitData
      .EnhanceGrid
    End With
 Next
End Sub

