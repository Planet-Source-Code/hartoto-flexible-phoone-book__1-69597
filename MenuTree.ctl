VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl MenuTree 
   Alignable       =   -1  'True
   BackColor       =   &H80000013&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3120
      ScaleHeight     =   1215
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   1065
      Visible         =   0   'False
      Width           =   75
   End
   Begin MSComctlLib.TreeView tvMenus 
      Height          =   3135
      Left            =   60
      TabIndex        =   1
      Top             =   330
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   5530
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Height          =   195
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   50
      Width           =   195
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   2040
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   2040
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Image imgSplitter 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   4200
      MousePointer    =   99  'Custom
      Top             =   1170
      Width           =   75
   End
End
Attribute VB_Name = "MenuTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private miSplitterLeftOffset As Integer
'Private mbResizing As Boolean
'Private m_MinHorizontalSize As Integer

Event CloseMe()
Event NodeClick(ByVal Node As MSComctlLib.Node)
'

Public Property Get MenuTreeView() As TreeView
    Set MenuTreeView = tvMenus
End Property

Private Sub cmdClose_Click()
    RaiseEvent CloseMe
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iNewPos As Integer
    
    'As the mouse moves we need to also move the
    'picSplitter control.  We need to contain
    'the splitter to the area of its parent control
    'and only allow it to show in valid areas.
'    If mbResizing Then
'        iNewPos = UserControl.Extender.Left + imgSplitter.Left + x - miSplitterLeftOffset
'        If iNewPos < UserControl.Extender.Left + m_MinHorizontalSize Then
'            picSplitter.Left = UserControl.Extender.Left + m_MinHorizontalSize
'        ElseIf iNewPos > UserControl.Parent.width - 240 Then
'            picSplitter.Left = UserControl.Parent.width - 240
'        Else
'            picSplitter.Left = iNewPos
'        End If
'    End If
End Sub

Private Sub tvMenus_NodeClick(ByVal Node As MSComctlLib.Node)
    RaiseEvent NodeClick(Node)
End Sub

Private Sub UserControl_Initialize()
    'ThinBorder UserControl.hwnd, False
End Sub

Private Sub UserControl_Resize()
    'On Error Resume Next
    Dim lWidth As Long, lHeight As Long
    lWidth = UserControl.ScaleWidth
    lHeight = UserControl.ScaleHeight
    
    With imgSplitter
        .Move lWidth - .Width, 0, .Width, lHeight
        picSplitter.Move .Left, 0, .Width, lHeight
    End With
    With cmdClose
        .Move lWidth - .Width - 30 - imgSplitter.Width, 60, .Width, .Height
        tvMenus.Move 60, .Top + .Height + 30, _
            lWidth - 80 - imgSplitter.Width, _
            lHeight - (.Top + .Height + 100)
    End With
    
    picSplitter.ZOrder 0
End Sub
