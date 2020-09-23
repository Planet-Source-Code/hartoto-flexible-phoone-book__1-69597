VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.UserControl Srch 
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   Picture         =   "Srch.ctx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   8265
   Begin VB.CommandButton cmd 
      BackColor       =   &H80000016&
      Caption         =   "&Delete"
      Height          =   555
      Index           =   2
      Left            =   4560
      Picture         =   "Srch.ctx":590F
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Edit"
      Height          =   555
      Index           =   1
      Left            =   3240
      Picture         =   "Srch.ctx":5B73
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Add"
      Height          =   555
      Index           =   0
      Left            =   1920
      Picture         =   "Srch.ctx":5D9F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H80000016&
      Caption         =   "&Close"
      Height          =   555
      Index           =   3
      Left            =   5880
      Picture         =   "Srch.ctx":5F55
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1200
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   320
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
   Begin VB.ComboBox cboCol 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Cari"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "&Refresh All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6360
      Picture         =   "Srch.ctx":6111
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd 
      Height          =   2220
      Left            =   960
      TabIndex        =   8
      Top             =   2640
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3916
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   -2147483629
      ForeColorFixed  =   -2147483647
      GridColor       =   12632256
      ScrollTrack     =   -1  'True
      FillStyle       =   1
      GridLinesFixed  =   1
      GridLinesUnpopulated=   1
      MergeCells      =   2
      AllowUserResizing=   1
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
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imgFind 
      Height          =   195
      Left            =   70
      Picture         =   "Srch.ctx":6328
      Stretch         =   -1  'True
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgVBar 
      Height          =   10770
      Left            =   0
      Picture         =   "Srch.ctx":6507
      Top             =   0
      Width           =   345
   End
   Begin VB.Shape shpTop 
      DrawMode        =   12  'Nop
      Height          =   1215
      Left            =   480
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Column"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblJudul 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   0
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Picture         =   "Srch.ctx":6ECF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11715
   End
End
Attribute VB_Name = "Srch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_sRptName As String
Private m_sgolPersh As String
Private m_sKodeTrans As String
Private m_sJudul As String

Private m_sTblCounter As String
Private m_sKodeTransComplete As String

Private m_bCariHPP As Boolean
Private m_bLookup As Boolean

Event grdDblClick()
Event AddNew()
Event editRec()
Event Delete()
Event onClose()
Event onCetak()

Private m_sCari As String
Private m_sKeyFd As String
Private m_sOrderFd As String
Private m_bMDetail As Boolean
Private m_sMasterFd As String
Private m_sMasterKeyFd As String

Private m_sXMLfile As String
Private m_sConn As String
Private m_sSQL As String
'
Private oGrd As clsGrid

'06 des 2004
Private Enum icmd
 iNew = 0
 iModify = 1
 iDelete = 2
 iClose = 3
 iCetak = 4
End Enum
Private Sub ActionKey()

End Sub

Public Property Get MasterFd() As String
MasterFd = m_sMasterFd
End Property

Public Property Let MasterFd(ByVal vNewValue As String)
m_sMasterFd = vNewValue
End Property

Private Sub cboCol_Click()
 txtSearch.SetFocus
End Sub

Private Sub cboCol_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
      cmd(icmd.iClose).Value = True
    Case vbKeyInsert
      cmd(icmd.iNew).Value = True
  End Select
End Sub

Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
      cmd(icmd.iClose).Value = True
    Case vbKeyInsert
      cmd(icmd.iNew).Value = True
  End Select
End Sub

Public Sub cmdAll_Click()
  m_sKeyFd = ""
  m_sCari = ""
  InitData
  EnhanceGrid
End Sub

Private Sub cmdAll_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
      cmd(icmd.iClose).Value = True
    Case vbKeyInsert
      cmd(icmd.iNew).Value = True
  End Select
End Sub

Private Sub cmdSearch_Click()
 If cboCol.Text <> "" And txtSearch.Text <> "" Then
   m_sSQL = getSQLXML(m_sXMLfile, cboCol.Text, txtSearch.Text)
   'MsgBox m_sSQL & " " & "SearchForm.cmdSearch_click"
   Dim oRs As New ADODB.Recordset
   
   oRs.open m_sSQL, conn, adOpenKeyset
   Set grd.DataSource = oRs
   grd.Refresh
   Set oRs = Nothing
   EnhanceGrid
 End If
End Sub


Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyEscape
      cmd_Click (icmd.iClose)
    
    Case vbKeyEnd
      'cmdLast_Click
      
    Case vbKeyHome
      'cmdFirst_Click
      
    Case vbKeyInsert
      cmd(icmd.iNew).Value = True
      
    Case vbKeyDelete
      'MsgBox "delete"
      cmd_Click (icmd.iDelete)
  End Select
End Sub

Private Sub txtSearch_Change()
  cmdSearch.Value = True
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
      cmd(icmd.iClose).Value = True
    Case vbKeyInsert
      cmd(icmd.iNew).Value = True
  End Select
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   grd.SetFocus
 End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
     ' cmdClose_Click
    
    Case vbKeyEnd
     ' cmdLast_Click
      
    Case vbKeyHome
     ' cmdFirst_Click
      
    Case vbKeyUp, vbKeyPageUp
     ' If Shift = vbCtrlMask Then
     '   cmdFirst_Click
     ' Else
     '   cmdPrevious_Click
     ' End If
      
    Case vbKeyDown, vbKeyPageDown
     ' If Shift = vbCtrlMask Then
     '   cmdLast_Click
     ' Else
     '   cmdNext_Click
     ' End If
      
    Case vbKeyDelete
      MsgBox "delete"
  End Select
End Sub

Private Sub UserControl_Resize()
  
  Image1.Width = UserControl.Width
  grd.Width = UserControl.Width - 300 - imgVBar.Width
  grd.Height = UserControl.Height - 3150
  shpTop.Width = grd.Width
  'Frame1.width = UserControl.width - 300
  grd.Left = 100 + imgVBar.Width
  'Frame1.Left = 100
  txtSearch.Width = UserControl.Width - 3500
  cboCol.Width = UserControl.Width - 3500
  'Picture1.Top = UserControl.Height - 1100
'  Picture1.Width = UserControl.Width - 350
  cmdSearch.Left = UserControl.Width - 2400
  cmdAll.Left = UserControl.Width - 1400

 For i = 0 To cmd.Count - 1
    Jrk = Jrk + 50
    cmd(cmd.Count - 1 - i).Left = UserControl.Width - cmd(i).Width * (i + 1) - 180 - Jrk
  Next i
  
End Sub

Private Sub UserControl_Show()
  UserControl.BackColor = Ambient.BackColor
'  Frame1.BackColor = Ambient.BackColor
 
End Sub

Sub InitData()
  
  Set oDom = New MSXML.DOMDocument
  oDom.Load m_sXMLfile
      
  Do
   DoEvents
  Loop Until oDom.readyState = XML_LOAD_COMPLETE
  'MsgBox m_sXMLfile & " " & oDom.xml

  Set conn = New ADODB.Connection
  Set RS = New ADODB.Recordset
  
  'MsgBox sConn
  
  conn.ConnectionString = sConn
  conn.open
  
  If m_sCari = "" Then
    m_sSQL = getSQLXML(m_sXMLfile, "", "")
    'MsgBox m_sSQL
  Else
    m_sSQL = getSQLXML(m_sXMLfile, m_sKeyFd, m_sCari)
    'MsgBox m_sSQL
  End If
  
  'MsgBox m_sSQL
  RS.open m_sSQL, conn, adOpenDynamic, adLockPessimistic
  grd.Clear
  Set grd.DataSource = RS
  
  Set oConn = Nothing
  Set RS = Nothing
  
  Fillcombo
  'cmd(icmd.iNew).SetFocus
  cboCol.ListIndex = 1
  
End Sub
Private Sub Fillcombo()
  'On Error Resume Next
  
  Dim i As Integer
  Dim oNodeTemp As IXMLDOMNode
  
  
  Set oNodeTemp = oDom.selectSingleNode("//search[@name=""" & "grid" & """]")
  'MsgBox oNodeTemp.xml
  Set oNodeTemp = oNodeTemp.childNodes(0).childNodes(0)
 
  cboCol.Clear
  For i = 0 To oNodeTemp.childNodes.length - 1
    cboCol.AddItem oNodeTemp.childNodes(i).Attributes(0).Text
  Next i
  
  Set oNodeTemp = Nothing
  
End Sub
Function getSQLXML(m_sXMLfile, sKeyFd, sSearch) As String
  
  Dim s As String, i As Integer, pj As Integer
  
  pj = Len(sSearch)
    
  'MsgBox m_sXMLfile
  'MsgBox oDom.xml
  
  Set oNode = oDom.selectSingleNode("//editform[@name=""" & "myform" & """]")
  'MsgBox oNode.xml
  m_sOrderFd = oNode.Attributes(2).Text
  m_sTblCounter = oNode.Attributes(4).Text
  m_sKodeTrans = oNode.Attributes(5).Text
  'MsgBox "m_sKodeTrans = " & m_sKodeTrans

  s = "SELECT "
  For i = 0 To oNode.childNodes.length - 1
   s = s & oNode.childNodes(i).Attributes(2).Text
   s = s & ","
  Next i
  s = Left(s, Len(s) - 1)
  s = s & " FROM "
  s = s & oNode.Attributes(1).Text
  
  If sKeyFd <> "" Then
      s = s & " WHERE " & sKeyFd & " like '%" & sSearch & "%'"
  End If
  s = s & " ORDER BY " & m_sOrderFd & " ASC"
  'MsgBox IsNumeric(sSearch)
  'MsgBox s & " " & "SearchForm_getSQLXML"
  getSQLXML = s
End Function
Function getSQLXMLNode(m_sXMLfile, sKeyFd, sSearch, _
  oNode As IXMLDOMNode) As String
  
  Dim s As String, i As Integer, pj As Integer
  pj = Len(sSearch)
  
  'Set oNode = oDom.selectSingleNode("//fields[@name=""" & "myfields" & """]")
  MsgBox oNode.xml
  m_sOrderFd = oNode.Attributes(2).Text
  m_sTblCounter = oNode.Attributes(4).Text

  s = "SELECT "
  For i = 0 To oNode.childNodes.length - 1
   s = s & oNode.childNodes(i).Attributes(0).Text
   s = s & ","
  Next i
  s = Left(s, Len(s) - 1)
  s = s & " FROM "
  s = s & oNode.Attributes(1).Text
  
  If sKeyFd <> "" Then
    If IsNumeric(sSearch) Then
      s = s & " WHERE " & sKeyFd & " = " & sSearch
    Else
      's = s & " WHERE LEFT(" & sKeyFd & "," & pj & ")" & " LIKE '" & sSearch & "%'"
      s = s & " WHERE " & sKeyFd & " LIKE '%" & sSearch & "%'"
    End If
  End If
  s = s & " ORDER BY " & m_sOrderFd & " DESC"
  'MsgBox IsNumeric(sSearch)
  'MsgBox s & " " & "SearchForm_getSQLXML"
  getSQLXMLNode = s
End Function
Sub EnhanceGrid()
  
Dim i As Integer

Set oGrd = New clsGrid

Dim oNodeGrid As IXMLDOMNode

Set oNodeGrid = oDom.selectSingleNode("//search[@name=""" & "grid" & """]")
'MsgBox oNodeGrid.xml
'm_bMDetail = CBool(oNodeGrid.Attributes(2).Text)
lblJudul.Caption = oNodeGrid.Attributes(3).Text
m_sJudul = oNodeGrid.Attributes(3).Text

Set oNodeGrid = oNodeGrid.childNodes(0).childNodes(0)

ReDim grdColName(0 To oNodeGrid.childNodes.length - 1) As String
For i = 0 To oNodeGrid.childNodes.length - 1
  'judul
  grd.TextMatrix(0, i) = oNodeGrid.childNodes(i).Attributes(1).Text
  grd.ColWidth(i) = CLng(oNodeGrid.childNodes(i).Attributes(2).Text)
  oGrd.colName = oNodeGrid.childNodes(i).Attributes(0).Text
  oGrd.AddCol
Next i
    
Exit Sub
  
localErr:

MsgBox "'" & m_sXMLfile & "'" & " not Loaded!", vbCritical, "Error"
End Sub

Sub cmd_Click(Index As Integer)

Dim sTblName As String, sSQL As String
Dim s As String, sSQLTemp As String
Dim rsTemp As New ADODB.Recordset
Dim connTemp As New ADODB.Connection

Select Case Index
 
 Case icmd.iNew
   m_sCari = ""
   m_sKeyFd = ""
   m_sMasterKeyFd = oGrd.oColNames(1)
   m_sMasterFd = grd.TextMatrix(grd.Row, 0)
   'm_sMasterFd = grd.TextMatrix(
   EditMode = lEditMode.Add
   RaiseEvent AddNew
 
 Case icmd.iModify
   m_sCari = grd.TextMatrix(grd.Row, 0)
   m_sKeyFd = oGrd.oColNames(1)
   
   'MsgBox m_sKeyFd & " " & m_sCari, , "searchform.cmd_click"
   'Stop
   'm_sMasterKeyFd = oGrd.oColNames(2)
   'm_sMasterFd = grd.TextMatrix(grd.row, 1)
   'MsgBox oGrd.oColNames(1)
   'm_sKodeTransComplete = grd.TextMatrix(grd.row, 1)
   'MsgBox m_sKodeTransComplete
   EditMode = lEditMode.Edit
   
   RaiseEvent editRec
 
 Case icmd.iDelete
  ShowDialog ("Hapus Data Ini?")
  If DialogFrm.bYa = True Then
    m_sCari = grd.TextMatrix(grd.Row, 0)
    m_sKeyFd = oGrd.oColNames(1)
    
    Set ocmd = New ADODB.Command
    sTblName = oNode.Attributes(1).Text
    
    'If IsNumeric(m_sCari) Then
    '  sSQL = "DELETE FROM " & sTblName & _
    '    " WHERE " & m_sKeyFd & "=" & m_sCari
    'Else
      sSQL = "DELETE FROM " & sTblName & _
        " WHERE " & m_sKeyFd & "='" & m_sCari & "'"
    'End If
    
    'MsgBox sSQL
    With ocmd
      .ActiveConnection = conn
      .CommandText = sSQL
      .Execute
    End With
     
    RS.Requery
    m_sKeyFd = ""
    m_sCari = ""
    InitData
    EnhanceGrid
  End If
  
 Case icmd.iClose
   RaiseEvent onClose
   
 
   
   RaiseEvent onCetak
   
End Select
End Sub

Private Sub Form_Load()
 

End Sub

Private Sub Form_Resize()
  'On Error Resume Next
  'This will resize the grid when the form is resized
  grd.Width = UserControl.Width
  grd.Height = UserControl.Height - grdDataGrid.Top - 30 - picButtons.Height - picStatBox.Height
  lblStatus.Left = UserControl.Width - lblStatus.Width + 400
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  'PrimaryCLS.AddNew
  'lblStatus.Caption = "Add record"
  'mbAddNewFlag = True
  SetButtons False

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  CmdAdd.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub grd_DblClick()
  cmd(icmd.iModify).Value = True
End Sub
Private Sub grd_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  grd_DblClick
End If

End Sub

Private Sub txtFields_KeyPress(Index As Integer, _
  KeyAscii As Integer)
  
  'make kapital letter
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  
End Sub
Public Property Get XMLFile() As String
  XMLFile = m_sXMLfile
  
End Property

Public Property Let XMLFile(ByVal vNewValue As String)
  m_sXMLfile = vNewValue
  
End Property
Public Property Get OrderFd() As Variant
  OrderFd = m_OrderFd
  
End Property
Public Property Get bMDetail() As Boolean
  bMDetail = m_bMDetail
  
End Property
Public Property Get bCariHPP() As Boolean
  bCariHPP = m_bCariHPP
  
End Property
Public Property Let bCariHPP(bNewCariHPP As Boolean)
  m_bCariHPP = bNewCariHPP
  
End Property


Public Property Get sCari() As String
  sCari = m_sCari
  
End Property
Public Property Let sCari(sNew As String)
  m_sCari = sNew
  
End Property
Public Property Get KeyFd() As String
 KeyFd = m_sKeyFd
End Property

Public Property Let KeyFd(newKeyFd As String)
 m_sKeyFd = newKeyFd
End Property

Public Property Get bLookup() As Boolean
 bLookup = m_bLookup
 
End Property

Public Property Let bLookup(ByVal vNewValue As Boolean)
 m_bLookup = vNewValue
 
End Property
Public Property Get MasterKeyFd() As String
MasterKeyFd = m_sMasterKeyFd
End Property

Public Property Let MasterKeyFd(ByVal vNewValue As String)
m_sMasterKeyFd = vNewValue
End Property

Public Property Let KodeTransComplete(ByVal vNewValue As String)
 m_sKodeTransComplete = vNewValue
End Property
Public Property Get KodeTransComplete() As String
 KodeTransComplete = m_sKodeTransComplete
End Property

Public Property Get Judul() As String
Judul = m_sJudul
End Property

Public Property Let Judul(ByVal vNewValue As String)
 m_sJudul = vNewValue
End Property
Public Property Let KodeTrans(ByVal vNewValue As String)
 m_sKodeTrans = vNewValue
End Property
Public Property Get KodeTrans() As String
 KodeTrans = m_sKodeTrans
End Property

Public Property Get SQL() As String
  SQL = m_sSQL
End Property

Public Property Let SQL(ByVal vNewValue As String)
  m_sSQL = vNewValue
End Property

Public Property Get golPersh() As String
golPersh = m_sgolPersh
End Property

Public Property Let golPersh(ByVal vNewValue As String)
m_sgolPersh = vNewValue
End Property

Public Property Get RptName() As String
RptName = m_sRptName
End Property

Public Property Let RptName(ByVal vNewValue As String)
m_sRptName = vNewValue
End Property
