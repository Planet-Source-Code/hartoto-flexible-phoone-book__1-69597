VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Edit 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   Picture         =   "Edit.ctx":0000
   ScaleHeight     =   2235
   ScaleWidth      =   4575
   Begin VB.CommandButton cmd 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   1965
      Picture         =   "Edit.ctx":590F
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1200
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   1
      Left            =   3240
      Picture         =   "Edit.ctx":5AFE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1200
   End
   Begin VB.CommandButton cmdLkp 
      Caption         =   "&..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   0
      Left            =   3960
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   320
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   16761024
      Format          =   143589377
      CurrentDate     =   38542
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   2400
      MaxLength       =   47
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgEdit 
      Height          =   285
      Left            =   60
      Picture         =   "Edit.ctx":5CBA
      Top             =   480
      Width           =   285
   End
   Begin VB.Shape shptxt 
      BorderColor     =   &H8000000D&
      Height          =   320
      Index           =   0
      Left            =   1680
      Top             =   480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image imgVbar 
      Height          =   10770
      Left            =   0
      Picture         =   "Edit.ctx":5F14
      Top             =   0
      Width           =   345
   End
   Begin VB.Label lblJudul 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "Edit.ctx":68DC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11370
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'sep 06
Private m_sTblName As String

Private lHeight As Long
Private m_sJudul   As String
Private m_sTblCounter As String
Private m_sLookupMD As String

Private m_sLookupNameFd As String
Private m_sLookupKeyFd As String
Private m_bIsGetLookup As Boolean
Private m_bLookup As Boolean

Event onCancel()
Event onDataRefresh()
Event onLookup()

Private Enum icmd
 save = 0
 Cancel = 1
End Enum

Public EdMode As Long

Private m_sCari As String
Private m_KeyFd As String

Private m_sXMLfile As String
'
Sub getLookup()
  'MsgBox m_sLookupKeyFd & " " & m_sLookupNameFd, vbInformation, "EditFormDetails_getLookup"
  txt(3).Text = m_sLookupKeyFd
  txt(4).Text = m_sLookupNameFd
  'Label1(0).Caption = m_sLookupKeyFd
  'Label1(1).Caption = m_sLookupNameFd
  
  txt(5).SetFocus
End Sub
Function getObjIndex(ByVal sTag As String) As Integer
 For Each ctl In UserControl.Controls
  If ctl.Tag <> "" Then
   If ctl.Tag = sTag Then getObjIndex = ctl.Index
   'MsgBox sTag & " " & ctl.Tag
  End If
'   MsgBox ctl.Index
 Next
End Function
Function getSQLXML(ByVal oNode As IXMLDOMNode, sKeyFd, sSearch) As String
  
    Dim s As String, i As Integer
  
    s = "SELECT "
    For i = 0 To oNode.childNodes.length - 1
     s = s & oNode.childNodes(i).Attributes(2).Text
     s = s & ","
    Next i
    s = Left(s, Len(s) - 1)
    s = s & " FROM "
    s = s & oNode.Attributes(1).Text
    
    'If m_KeyFd <> "" Then
    If EditMode = lEditMode.Edit Then
    'MsgBox sSearch & " " & IsNumeric(sSearch)
      'If IsNumeric(sSearch) Then
      '  s = s & " WHERE " & m_KeyFd & " = " & sSearch
      'Else
        s = s & " WHERE " & m_KeyFd & " = '" & sSearch & "'"
      'End If
    End If
    
getSQLXML = s
End Function

Public Property Let XMLFile(ByVal vNewValue As String)
  m_sXMLfile = vNewValue
  
End Property
Public Property Get XMLFile() As String
  XMLFile = m_sXMLfile
  
End Property

Sub InitData()
  
  'On Error Resume Next
  Dim i As Integer
  Dim rsTemp As New ADODB.Recordset
  Dim rsKadis As New ADODB.Recordset
  
  Set oDom = New MSXML.DOMDocument
  oDom.Load m_sXMLfile

  Do
   DoEvents
  Loop Until oDom.readyState = XML_LOAD_COMPLETE
  'MsgBox m_sXMLfile & " " & oDom.xml
    
  Set oNode = oDom.selectSingleNode("//editform[@name=""" & "myform" & """]")
  'm_sTblCounter = oNode.Attributes(2).Text
  
  drawForm
      
  Set conn = New ADODB.Connection
  Set RS = New ADODB.Recordset
  
  'MsgBox oNode.xml
  m_sTblName = oNode.Attributes(1).Text
  
  conn.open sConn
  'If m_sCari = "" Then
  If EditMode = lEditMode.Add Then
    m_sSQL = getSQLXML(oNode, "", "")
  Else
   m_sSQL = getSQLXML(oNode, m_sKeyFd, m_sCari)
  End If
  
 RS.open m_sSQL, conn, adOpenDynamic, adLockOptimistic
 
 Dim OldCount As Long, NewCount As Long
 Dim oConn As New ADODB.Connection
 'MsgBox getReadOnlyRS("SELECT COUNT(*) as RecCount FROM " & m_sTblName).Fields("RecCount")
 Dim oRsTmp As New ADODB.Recordset
 oConn.open sConn
 oRsTmp.open "Select FCounter as RecCount from TblCounter", oConn, adOpenKeyset
 OldCount = oRsTmp.Fields("RecCount")
 NewCount = OldCount + 1
 Set oRsTmp = Nothing
  
'MsgBox m_sTblCounter
If EditMode = lEditMode.Add Then
    'txt(1).Text = getAutoNum(getCounterRS(m_sTblCounter))
    'txt(1).Enabled = False
    'txt(2).Text = getAutoTranCode(m_sKodeTrans, getAutoNum(getCounterRS(m_sTblCounter)))
    'm_sKodeTransComplete = getAutoTranCode(m_sKodeTrans, getAutoNum(getCounterRS(m_sTblCounter)))
    'dtp(3).Value = Date
    'MsgBox m_sSQL, vbInformation, "MasterDetails_InitData"
    'cbo(getObjIndex("cboGol")).ListIndex = 0
    txt(1).Text = NewCount
    
  Else
     'edit mode
      For Each ctl In UserControl.Controls
        i = 1
        If ctl.Tag <> "" Then
          If Left(ctl.Tag, 3) = "txt" Then
            If Not IsNull(RS.Fields(txt(getObjIndex(ctl.Tag)).DataField)) Then
              txt(getObjIndex(ctl.Tag)).Text = RS.Fields(txt(getObjIndex(ctl.Tag)).DataField)
              'MsgBox RS.Fields(txt(getObjIndex(ctl.Tag)).DataField)
            End If
          End If
          'If Left(ctl.Tag, 3) = "dbc" Then
            'dbc(getObjIndex(ctl.Tag)).Text = rs.Fields(dbc(getObjIndex(ctl.Tag)).DataField)
          'End If
          If Left(ctl.Tag, 3) = "dtp" Then
            dtp(getObjIndex(ctl.Tag)).Value = RS.Fields(dtp(getObjIndex(ctl.Tag)).DataField)
          End If
          If Left(ctl.Tag, 3) = "chk" Then
            chk(getObjIndex(ctl.Tag)).Value = Abs(RS.Fields(chk(getObjIndex(ctl.Tag)).DataField))
          End If
          If Left(ctl.Tag, 3) = "cbo" Then
            'Set rsTemp = getReadOnlyRS("SELECT KODE_GOL, GOLONGAN FROM GOL_PERUSAHAAN WHERE KODE_GOL = " & cbo(getObjIndex(ctl.Tag)).Text)
            'cbo(getObjIndex(ctl.Tag)).Text = rs.Fields(cbo(getObjIndex(ctl.Tag)).DataField)
            'cbo(getObjIndex(ctl.Tag)).Text = rs.Fields("KODE_GOL") & " " & rs.Fields("GOLONGAN")
          End If
        End If
      Next
      
  End If
   
  txt(1).Enabled = False
  
   
End Sub
Sub drawForm()
  
  Dim lwidth2 As Long, lWidth As Long
  'Dim ctl As Object
  Dim i As Integer, sObj As String
  m_sTblCounter = oNode.Attributes(4).Text
  For i = 0 To oNode.childNodes.length - 1
    sObj = oNode.childNodes(i).Attributes(0).Text
    'MsgBox sObj & " " & i & oNode.childNodes(i).Attributes(5).Text
    
    If sObj = "textbox" Then
        Load txt(i + 1)
        txt(i + 1).Tag = oNode.childNodes(i).Attributes(1).Text
        If txt(i + 1).Tag = "txtPassword" Then
          txt(i + 1).PasswordChar = "*"
        End If
        txt(i + 1).DataField = oNode.childNodes(i).Attributes(2).Text
        txt(i + 1).Left = oNode.childNodes(i).Attributes(3).Text
        txt(i + 1).Top = oNode.childNodes(i).Attributes(4).Text
        txt(i + 1).Width = oNode.childNodes(i).Attributes(5).Text
        txt(i + 1).Visible = True
        txt(i + 1).TabIndex = i
        
        Load shptxt(i + 1)
        shptxt(i + 1).Height = txt(i + 1).Height + 30
        shptxt(i + 1).Left = txt(i + 1).Left - 15
        shptxt(i + 1).Top = txt(i + 1).Top - 15
        shptxt(i + 1).Width = txt(i + 1).Width + 35
        shptxt(i + 1).Visible = True
          
        Load lbl(i + 1)
        lbl(i + 1).Top = oNode.childNodes(i).Attributes(4).Text
        lbl(i + 1).Caption = oNode.childNodes(i).Attributes(6).Text
        lbl(i + 1).Left = txt(i + 1).Left - lbl(i + 1).Width - 150
        lbl(i + 1).Visible = True
        
        
        lHeight = lHeight + txt(i + 1).Height
        lWidth = lbl(i + 1).Width + txt(i + 1).Width
        'MsgBox txt(i + 1).Width
     End If
    
    If sObj = "dtp" Then
        Load dtp(i + 1)
        dtp(i + 1).Tag = oNode.childNodes(i).Attributes(1).Text
        dtp(i + 1).DataField = oNode.childNodes(i).Attributes(2).Text
        dtp(i + 1).Left = oNode.childNodes(i).Attributes(3).Text
        dtp(i + 1).Top = oNode.childNodes(i).Attributes(4).Text
        dtp(i + 1).Width = oNode.childNodes(i).Attributes(5).Text
        'dtp(i + 1).Index = i + 1
        dtp(i + 1).Visible = True
        Load lbl(i + 1)
        lbl(i + 1).Top = oNode.childNodes(i).Attributes(4).Text
        lbl(i + 1).Caption = oNode.childNodes(i).Attributes(6).Text
        lbl(i + 1).Left = dtp(i + 1).Left - lbl(i + 1).Width
        lbl(i + 1).Visible = True
        lHeight = lHeight + dtp(i + 1).Height
    End If
    
    If sObj = "combo" Then
        Load cbo(i + 1)
        cbo(i + 1).Tag = oNode.childNodes(i).Attributes(1).Text
        cbo(i + 1).DataField = oNode.childNodes(i).Attributes(2).Text
        cbo(i + 1).Left = oNode.childNodes(i).Attributes(3).Text
        cbo(i + 1).Top = oNode.childNodes(i).Attributes(4).Text
        cbo(i + 1).Width = oNode.childNodes(i).Attributes(5).Text
        'cbo(i + 1).Index = i + 1
        cbo(i + 1).Visible = True
        
        Load lbl(i + 1)
        lbl(i + 1).Top = oNode.childNodes(i).Attributes(4).Text
        lbl(i + 1).Caption = oNode.childNodes(i).Attributes(6).Text
        lbl(i + 1).Left = cbo(i + 1).Left - lbl(i + 1).Width
        lbl(i + 1).Visible = True
        lHeight = lHeight + cbo(i + 1).Height

    End If
  Next i
  
  lblJudul.Caption = Me.Judul
  
  
  UserControl.Height = (lHeight) + cmd(0).Height + 1200
  UserControl.Width = lWidth + 50 + imgVbar.Width + 500
  imgEdit.Top = Image1.Height
  
  'MsgBox lWidth
 
End Sub
Private Sub cmd_Click(Index As Integer)
 
 Dim sKeyFd As String, sTblName As String
 Dim oNode As IXMLDOMNode
 Dim oConn As New ADODB.Connection
 
 Set oNode = oDom.selectSingleNode("//editform[@name=""" & "myform" & """]")
 
 sKeyFd = oNode.Attributes(6).Text
 sTblName = oNode.Attributes(1).Text
 Select Case Index
  Case icmd.save
    Select Case EditMode
      Case 0 'add
        
        UpdateData lEditMode.Add, "Attribute"
        Dim OldCount As Long, NewCount As Long
        Dim oRsTmp As New ADODB.Recordset
        oConn.open sConn
        oRsTmp.open "Select FCounter as RecCount from TblCounter", oConn, adOpenKeyset
        OldCount = oRsTmp.Fields("RecCount")
        'MsgBox "oldCount = " & oldCount
        NewCount = OldCount + 1
        'MsgBox "NewCount = " & NewCount
        Set oRsTmp = Nothing
        
        Dim oRsUpdateCount As New ADODB.Recordset
        oRsUpdateCount.open "Update TblCounter set FCounter = " & NewCount, oConn, adOpenDynamic
        Set oRsUpdateCount = Nothing
        Set oConn = Nothing
        RaiseEvent onDataRefresh
       
      Case 1 'edit
        'Dim oConn As New ADODB.Connection
        Dim oRs As New ADODB.Recordset
        Dim s As String
        s = "DELETE FROM " & sTblName & " WHERE " & sKeyFd & "='" & txt(1).Text & "'"
        oConn.open sConn
        oRs.open s, oConn
        UpdateData lEditMode.Add, "Attribute"
        RaiseEvent onDataRefresh
    End Select
  
  Case icmd.Cancel
    RaiseEvent onCancel
   
  End Select
 
End Sub
Private Sub cmd_ClickX(Index As Integer)

Dim i As Integer, rsCounter As New ADODB.Recordset
Dim sSQLCounter As String
Dim rsTemp As New ADODB.Recordset
Dim connTemp As New ADODB.Connection

'On Error GoTo localError

Select Case Index
 Case icmd.save
   Select Case EditMode
'     Case lEditMode.Add
'       rs.AddNew
'       For Each ctl In UserControl.Controls
'        If ctl.Tag <> "" Then
'           If Left(ctl.Tag, 3) = "txt" Then
'             If txt(getObjIndex(ctl.Tag)).Text = "" Then
'               If txt(getObjIndex(ctl.Tag)).DataField = "MODAL" Or _
'                txt(getObjIndex(ctl.Tag)).DataField = "NO_URUT" _
'               Then
'                  txt(getObjIndex(ctl.Tag)).Text = 0
'                 Else
'                  txt(getObjIndex(ctl.Tag)).Text = Space(1)
'               End If
'             End If
'             rs.Fields(txt(getObjIndex(ctl.Tag)).DataField) = txt(getObjIndex(ctl.Tag)).Text
'           End If
'        End If
'       Next
       
       'Select Case plEditMode
   Case 0 'add record
      sSQL = getSQLXMLNode("", "", oNode, "Attribute", EditMode)
      'MsgBox sSQL, , "UpdateData_rsTemp"
     'Stop
      connTemp.open sConn
      rsTemp.open sSQL, connTemp, adOpenDynamic, adLockPessimistic
      rsTemp.AddNew
     
     For Each ctl In UserControl.Controls
       Select Case sXMLNodeType
         Case "Element"
            If ctl.Tag <> "" Then
               If Left(ctl.Tag, 3) = "txt" Then
                 rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).baseName) = txt(getObjIndex(ctl.Tag)).Text
               End If
               'If Left(ctl.Tag, 3) = "dbc" Then
               '  rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).baseName) = dbc(getObjIndex(ctl.Tag)).Text
               'End If
               If Left(ctl.Tag, 3) = "cbo" Then
                 rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).baseName) = cbo(getObjIndex(ctl.Tag)).Text
               End If
               If Left(ctl.Tag, 3) = "dtp" Then
                  rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).baseName) = dtp(getObjIndex(ctl.Tag)).Value
               End If
               If Left(ctl.Tag, 3) = "chk" Then
                  rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).baseName) = chk(getObjIndex(ctl.Tag)).Value
               End If
            End If
         
         Case "Attribute"
            If ctl.Tag <> "" Then
               If Left(ctl.Tag, 3) = "txt" Then
                 rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text) = txt(getObjIndex(ctl.Tag)).Text
               End If
               'If Left(ctl.Tag, 3) = "dbc" Then
               '  rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text) = dbc(getObjIndex(ctl.Tag)).Text
               'End If
               If Left(ctl.Tag, 3) = "cbo" Then
                 'MsgBox oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text & " " & ctl.Tag & " " & getObjIndex(ctl.Tag)
                 rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text) = cbo(getObjIndex(ctl.Tag)).Text
               End If
               If Left(ctl.Tag, 3) = "dtp" Then
                  'MsgBox oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text & " " & ctl.Tag & " " & getObjIndex(ctl.Tag)
                  rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text) = dtp(getObjIndex(ctl.Tag)).Value
                  'MsgBox "rsTemp.Fields(" & Chr(34) & oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text & Chr(34) & ") = " & dtp(getObjIndex(ctl.Tag)).Value
               End If
               If Left(ctl.Tag, 3) = "chk" Then
                 rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text) = chk(getObjIndex(ctl.Tag)).Value
               End If
            End If
         Case Else
           MsgBox " sXMLNodeType '" & sXMLNodeType & "' tidak diizinkan...!"
         End Select
     Next
     rsTemp.Update
 
       'For i = 0 To txt.Count - 1
         'If IsNull(txt(i + 1).Text) Then
         ' txt(i + 1).Text = 0
         'End If
       '  rs.Fields(i) = txt(i + 1).Text
       'Next i
       RS.Update
       
       sSQLCounter = "SELECT NO_URUT FROM " & m_sTblCounter
       'MsgBox sSQLCounter
       rsCounter.open sSQLCounter, conn, adOpenDynamic, adLockPessimistic
       rsCounter.AddNew
       rsCounter.Fields("NO_URUT") = CLng(getAutoNum(getCounterRS(m_sTblCounter)))
       rsCounter.Update
       Set rsCounter = Nothing
       RaiseEvent onDataRefresh
     
     Case lEditMode.Edit
       'MsgBox m_sCari, vbInformation, "EditForm.cmd_click"
       RS.Delete adAffectCurrent
       RS.AddNew
       'For i = 0 To txt.Count - 1
         'MsgBox txt(i + 1).Text
      '   rs.Fields(i) = txt(i + 1).Text & " "
       'Next i
'       For Each ctl In UserControl.Controls
'        If ctl.Tag <> "" Then
'           If Left(ctl.Tag, 3) = "txt" Then
'             If txt(getObjIndex(ctl.Tag)).Text = "" Then
'               If txt(getObjIndex(ctl.Tag)).DataField = "MODAL" Or _
'               txt(getObjIndex(ctl.Tag)).DataField = "NO_URUT" _
'               Then
'                 txt(getObjIndex(ctl.Tag)).Text = 0
'               Else
'                 txt(getObjIndex(ctl.Tag)).Text = Space(1)
'               End If
'             End If
'             rs.Fields(txt(getObjIndex(ctl.Tag)).DataField) = txt(getObjIndex(ctl.Tag)).Text
'             'MsgBox rs.Fields(txt(getObjIndex(ctl.Tag)).DataField)
'           End If
'        End If
'       Next
       
    Case 1 'edit records
      sSQL = getSQLXMLNode(oNode.Attributes(6).Text, txt(2).Text, oNode, "Attribute", plEditMode)
      'MsgBox sSQL
      rsTemp.open sSQL, connTemp
   
       
       RS.Update
       RaiseEvent onDataRefresh
       
   End Select
 Case icmd.Cancel
   RaiseEvent onCancel
 End Select
 
 Exit Sub
 
localError:
  'MsgBox Err.Number & " " & Err.Description, vbCritical
  Resume Next
End Sub
Private Sub BlankForm()
        
'On Error Resume Next
For i = 0 To txt.Count - 1
  'MsgBox
  txt(i + 1).Text = ""
Next i
 
End Sub

Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case vbKeyEscape
      cmd(icmd.Cancel).Value = True
   
  End Select
End Sub

Private Sub txt_Change(Index As Integer)
  On Error Resume Next
  
  If Index = getObjIndex("txtModal") Then
    txt(getObjIndex("txtTerbilang")).Text = ToTerbilang(CDbl(txt(getObjIndex("txtModal")).Text)) & " rupiah"
    txt(getObjIndex("txtTerbilang")).Enabled = False
    
  End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case vbKeyEscape
      cmd(icmd.Cancel).Value = True
   
  End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
 
 Dim ctl As Control, i As Integer
 For Each ctl In UserControl.Controls
  If ctl.Tag <> "" Then
    i = i + 1
  End If
 Next
 
 If KeyAscii = 13 And txt(Index).Text = "?" Or txt(Index).Text = "*" Then
  m_sLookupMD = "NoMasterDetails"
  RaiseEvent onLookup
 End If
 
' MsgBox i & " " & Index
 
 If Index = i And KeyAscii = 13 Then
   cmd(icmd.save).SetFocus
 Else
    If KeyAscii = 13 Then
      SendKeys "{TAB}"
    End If
End If
 
'--- buat jadi huruf besar
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
 
 
 
End Sub

Private Sub UserControl_Resize()
  'Image1.Width = UserControl.Width
  cmd(0).Top = UserControl.Height - cmd(0).Height - 180
  cmd(1).Top = UserControl.Height - cmd(1).Height - 180
  cmd(0).Left = UserControl.Width - (2 * cmd(0).Width) - 100 - 100 - 50
  cmd(1).Left = UserControl.Width - (cmd(0).Width) - 150 - 50
End Sub

Private Sub UserControl_Show()
 UserControl.BackColor = Ambient.BackColor
 'MsgBox Me.Judul
 lblJudul.Caption = Me.Judul
End Sub

Public Property Let sCari(vNewValue As String)
 m_sCari = vNewValue
 
End Property
Public Property Let KeyFd(ByVal vNewValue As String)
 m_KeyFd = vNewValue
 
End Property

Public Property Get sCari() As String
 sCari = m_sCari
 
End Property
Public Property Get KeyFd() As String
 KeyFd = m_KeyFd
 
End Property
Public Property Get bLookup() As Boolean
 bLookup = m_bLookup
 
End Property

Public Property Let bLookup(ByVal vNewValue As Boolean)
 m_bLookup = vNewValue
End Property
Public Property Get LookupKeyFd() As String
 LookupKeyFd = m_sLookupKeyFd
End Property

Public Property Let LookupKeyFd(ByVal vNewValue As String)
m_sLookupKeyFd = vNewValue
End Property
Public Property Get LookupNameFd() As String
LookupNameFd = m_sLookupNameFd
End Property

Public Property Let LookupNameFd(ByVal vNewValue As String)
m_sLookupNameFd = vNewValue
End Property

Public Property Get bIsGetLookup() As Boolean
   bIsGetLookup = m_bIsGetLookup
End Property

Public Property Let bIsGetLookup(ByVal vNewValue As Boolean)
  m_bIsGetLookup = vNewValue
End Property

Public Property Get LookupMD() As String
 LookupMD = m_sLookupMD
 
End Property

Public Property Let LookupMD(ByVal vNewValue As String)
  m_sLookupMD = vNewValue
End Property
Public Property Get Judul() As String
 Judul = m_sJudul
End Property

Public Property Let Judul(ByVal vNewValue As String)
 m_sJudul = vNewValue
End Property


Function getSQLXMLNode(sKeyFd, sSearch, _
  oNode As IXMLDOMNode, sXMLNodeType As String, _
  ByVal lEditMode As Long) As String
  
  Dim s As String, i As Integer, pj As Integer
  pj = Len(sSearch)
  Dim sDataType As String
      
  'MsgBox oNode.xml, , "getSQLXMLNode"
  'MsgBox "lEditMode = " & lEditMode
'  m_sOrderFd = oNode.Attributes(2).Text
'  m_sTblCounter = oNode.Attributes(4).Text

  'MsgBox "oNode.childNodes.length - 1 = " & oNode.childNodes.length - 1
  sDataType = oNode.childNodes(i).Attributes(7).Text
  Select Case lEditMode
    Case 0 'for add record
        s = "SELECT "
        For i = 0 To oNode.childNodes.length - 1
         Select Case sXMLNodeType
           Case Is = "Element"
             s = s & oNode.childNodes(i).nodeName
           Case Is = "Attribute"
             If sDataType = "t_string" And oNode.childNodes(i).Attributes(2).Text = "" Then
                  oNode.childNodes(i).Attributes(2).Text = Space(1)
                ElseIf (sDataType = "t_int" Or sDataType = "t_long") And oNode.childNodes(i).Attributes(2).Text = "" Then
                  oNode.childNodes(i).Attributes(2).Text = 0
                ElseIf sDataType = "t_date" And oNode.childNodes(i).Attributes(2).Text = "" Then
                  oNode.childNodes(i).Attributes(2).Text = Date
                Else
                  'MsgBox "Type data '" & sDataType & "' belum didaftarkan!"
             End If
             s = s & oNode.childNodes(i).Attributes(2).Text
         Case Else
         End Select
         s = s & ","
        Next i
        s = Left(s, Len(s) - 1)
        s = s & " FROM "
        s = s & oNode.Attributes(1).Text
        
    Case 1 'for edit
       s = "UPDATE " & oNode.Attributes(1).Text
       s = s & " SET "
       For Each ctl In UserControl.Controls
        Select Case sXMLNodeType
          Case "Element"
            If ctl.Tag <> "" Then
               If Left(ctl.Tag, 3) = "txt" Then
                 s = s & oNode.childNodes(i).nodeName & " = '" & txt(getObjIndex(ctl.Tag)).Text & "'"
               End If
               'If Left(ctl.Tag, 3) = "dbc" Then
               '  S = S & oNode.childNodes(I).nodeName & " = '" & dbc(getObjIndex(ctl.Tag)).Text & "'"
               'End If
               If Left(ctl.Tag, 3) = "cbo" Then
                 s = s & oNode.childNodes(i).nodeName & " = '" & Left(cbo(getObjIndex(ctl.Tag)).Text, 1) & "'"
               End If
               If Left(ctl.Tag, 3) = "dtp" Then
                  s = s & oNode.childNodes(i).nodeName & " = '" & dtp(getObjIndex(ctl.Tag)).Value & "'"
               End If
               If Left(ctl.Tag, 3) = "chk" Then
                 s = s & oNode.childNodes(i).nodeName & " = '" & chk(getObjIndex(ctl.Tag)).Value & "'"
               End If
               s = s & ","
            End If
         
          Case "Attribute"
              If ctl.Tag <> "" Then
                sDataType = oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(7).Text
                If Left(ctl.Tag, 3) = "txt" Then
                  If getObjIndex(ctl.Tag) <> 1 Then 'primary key
                    If sDataType = "t_string" And txt(getObjIndex(ctl.Tag)).Text = "" Then
                        txt(getObjIndex(ctl.Tag)).Text = Space(1)
                      ElseIf (sDataType = "t_int" Or sDataType = "t_long") And txt(getObjIndex(ctl.Tag)).Text = "" Then
                        txt(getObjIndex(ctl.Tag)).Text = 0
                      ElseIf sDataType = "t_date" And txt(getObjIndex(ctl.Tag)).Text = "" Then
                        txt(getObjIndex(ctl.Tag)).Text = Date
                      Else
                        'MsgBox "Type data '" & sDataType & "' belum didaftarkan!"
                    End If
                    s = s & oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(2).Text & "= '" & txt(getObjIndex(ctl.Tag)).Text & "'"
                  End If
               End If
               'If Left(ctl.Tag, 3) = "dbc" Then
                 'S = S & oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text & "= '" & dbc(getObjIndex(ctl.Tag)).Text & "'"
               'End If
               If Left(ctl.Tag, 3) = "cbo" Then
                 s = s & oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(2).Text & "= '" & cbo(getObjIndex(ctl.Tag)).Text & "'"
               End If
               If Left(ctl.Tag, 3) = "dtp" Then
                  s = s & oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(2).Text & "= '" & dtp(getObjIndex(ctl.Tag)).Value & "'"
               End If
               If Left(ctl.Tag, 3) = "chk" Then
                  s = s & oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(2).Text & "= '" & chk(getObjIndex(ctl.Tag)).Value & "'"
               End If
                If getObjIndex(ctl.Tag) <> 1 Then 'primary key
                    s = s & ","
                End If
            End If
         
          Case Else
         
           MsgBox " sXMLNodeType '" & sXMLNodeType & "' tidak diizinkan...!"
         End Select
       Next
       s = Left(s, Len(s) - 1)

  End Select
        
  If sKeyFd <> "" Then
    'If IsNumeric(sSearch) Then
    '  s = s & " WHERE " & sKeyFd & " = " & sSearch
    'Else
    's = s & " WHERE LEFT(" & sKeyFd & "," & pj & ")" & " LIKE '" & sSearch & "%'"
      s = s & " WHERE " & sKeyFd & "= '" & sSearch & "'"
    'End If
  End If
  's = s & " ORDER BY " & m_sOrderFd & " DESC"
  'MsgBox IsNumeric(sSearch)
  'MsgBox s & " " & "SearchForm_getSQLXML"
  getSQLXMLNode = s
End Function


Sub UpdateData(plEditMode As Long, sXMLNodeType As String)
 
 'On Error Resume Next
 'On Error GoTo localErr
 
 Dim sDataType As String
 Dim connTemp As New ADODB.Connection
 Dim rsTemp As New ADODB.Recordset
 
 Dim s As String, sSQL As String
 Dim oXML As New MSXML.DOMDocument
 Dim oNode As IXMLDOMNode
 Dim i As Integer, vRecCount As Variant
 
 's = getXMLUpdateMaster(plEditMode)
 'MsgBox s, , "getXMLUpdateMaster"
 'Stop
     
 Dim h As Integer, sTemp As String
      
 '---- setting gridRows to update
 'handle grid content if "" then exit
 'Set oNode = oXML.loadXML(s)
 Set oNode = oDom.selectSingleNode("//editform[@name=""" & "myform" & """]")
 connTemp.CursorLocation = adUseServer
 connTemp.open sConn
  
 'MsgBox "plEditMode = " & plEditMode
 Select Case plEditMode
   Case 0 'add record
      sSQL = getSQLXMLNode("", "", oNode, "Attribute", plEditMode)
      'MsgBox sSQL, , "UpdateData_rsTemp"
      'Debug.Print sSQL
      
     'Stop
     rsTemp.open sSQL, connTemp, adOpenDynamic, adLockPessimistic
     rsTemp.AddNew
       
     For Each ctl In UserControl.Controls
       'MsgBox oNode.xml
       'sDataType = oNode.childNodes(7).Text
       'MsgBox oNode.childNodes(7).Text
       
       Select Case sXMLNodeType
         Case "Element"
            If ctl.Tag <> "" Then
               If Left(ctl.Tag, 3) = "txt" Then
                 rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).baseName) = txt(getObjIndex(ctl.Tag)).Text
               End If
               'If Left(ctl.Tag, 3) = "dbc" Then
               '  rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).baseName) = dbc(getObjIndex(ctl.Tag)).Text
               'End If
               If Left(ctl.Tag, 3) = "cbo" Then
                 rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).baseName) = cbo(getObjIndex(ctl.Tag)).Text
               End If
               If Left(ctl.Tag, 3) = "dtp" Then
                  rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).baseName) = dtp(getObjIndex(ctl.Tag)).Value
               End If
               If Left(ctl.Tag, 3) = "chk" Then
                  rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).baseName) = chk(getObjIndex(ctl.Tag)).Value
               End If
            End If
         
         Case "Attribute"
            If ctl.Tag <> "" Then
               sDataType = oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(7).Text
               'MsgBox oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(7).Text
               If Left(ctl.Tag, 3) = "txt" Then
                    If sDataType = "t_string" And txt(getObjIndex(ctl.Tag)).Text = "" Then
                      txt(getObjIndex(ctl.Tag)).Text = Space(1)
                    ElseIf (sDataType = "t_int" Or sDataType = "t_long") And txt(getObjIndex(ctl.Tag)).Text = "" Then
                      txt(getObjIndex(ctl.Tag)).Text = 0
                    ElseIf sDataType = "t_date" And txt(getObjIndex(ctl.Tag)).Text = "" Then
                      txt(getObjIndex(ctl.Tag)).Text = Date
                    Else
                      'MsgBox "Type data '" & sDataType & "' belum didaftarkan!"
                    End If
                    rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(2).Text) = txt(getObjIndex(ctl.Tag)).Text
               End If
               'If Left(ctl.Tag, 3) = "dbc" Then
                 'rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text) = dbc(getObjIndex(ctl.Tag)).Text
               'End If
               If Left(ctl.Tag, 3) = "cbo" Then
                 'MsgBox oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text & " " & ctl.Tag & " " & getObjIndex(ctl.Tag)
                 rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(2).Text) = cbo(getObjIndex(ctl.Tag)).Text
               End If
               If Left(ctl.Tag, 3) = "dtp" Then
                  'MsgBox oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text & " " & ctl.Tag & " " & getObjIndex(ctl.Tag)
                  rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(2).Text) = dtp(getObjIndex(ctl.Tag)).Value
                  'MsgBox "rsTemp.Fields(" & Chr(34) & oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(0).Text & Chr(34) & ") = " & dtp(getObjIndex(ctl.Tag)).Value
               End If
               If Left(ctl.Tag, 3) = "chk" Then
                 rsTemp.Fields(oNode.childNodes(getObjIndex(ctl.Tag) - 1).Attributes(2).Text) = chk(getObjIndex(ctl.Tag)).Value
               End If
            End If
         Case Else
           MsgBox " sXMLNodeType '" & sXMLNodeType & "' tidak diizinkan...!"
         End Select
     Next
     rsTemp.Update
          
                
   Case 1 'edit records
      Dim rsEdit As New ADODB.Recordset
      Dim oConn As New ADODB.Connection
      oConn.open sConn
      
      sSQL = getSQLXMLNode(oNode.Attributes(6).Text, txt(1).Text, oNode, "Attribute", plEditMode)
      rsEdit.open sSQL, oConn, adOpenDynamic, adLockPessimistic
      
 End Select
 
 
 
 Set rsTemp = Nothing
 Set connTemp = Nothing
'MsgBox sSQL
 
 Exit Sub
localErr:
  
 MsgBox Err.Description
 Resume Next
 
End Sub

