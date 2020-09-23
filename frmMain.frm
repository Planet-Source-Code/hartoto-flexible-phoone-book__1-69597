VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFC0C0&
   Caption         =   "XB-Stok"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8250
   Icon            =   "frmMain.frx":0000
   Picture         =   "frmMain.frx":164A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6135
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMenu 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   5475
      ScaleHeight     =   6105
      ScaleWidth      =   2745
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin Lensa.MenuTree mtMenus 
         Height          =   6255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2775
         _extentx        =   4895
         _extenty        =   11033
      End
   End
   Begin MSComctlLib.ImageList img 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16B30
            Key             =   "g1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1740A
            Key             =   "kasbank"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19114
            Key             =   "hp"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19F66
            Key             =   "laporan"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ADB8
            Key             =   "hutang"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BC0A
            Key             =   "g3"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CA5C
            Key             =   "g4"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D8AE
            Key             =   "g5"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E700
            Key             =   "gl"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2040A
            Key             =   "BookOpen"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2085C
            Key             =   "pembelian"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":216AE
            Key             =   "g8"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":233B8
            Key             =   "g10"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2352A
            Key             =   "BookPink"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2397C
            Key             =   "g2"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":247CE
            Key             =   "ProsesHPP"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuKeluar 
      Caption         =   "&Keluar"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Tampilan"
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuVMenu 
         Caption         =   "Menu"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPemakai 
      Caption         =   "&Pemakai"
      Begin VB.Menu mnuGantiPemakai 
         Caption         =   "&Ganti Pemakai"
      End
      Begin VB.Menu mnuGantiPassword 
         Caption         =   "&Ganti Password"
      End
   End
   Begin VB.Menu mnuLtkMenu 
      Caption         =   "&Letak Menu"
      Begin VB.Menu mnuLtkMKanan 
         Caption         =   "&Kanan"
      End
      Begin VB.Menu mnuLtkMKiri 
         Caption         =   "K&iri"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Private m_EnableAttr As ToolBarItems
Private Const mcstrMod$ = "frmMain"
'Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private oDom As MSXML.DOMDocument
Private m_xmlFile As String
'
Sub cekLogin()
  'If frmLogin.OK = True Then
  '  Unload frmLogin
  'Else
  
  'End If
End Sub


Public Sub SetStatus(Optional ByVal StatusText As String = vbNullString)
    'On Error Resume Next
    If StatusText = vbNullString Then
        Me.sbStatusBar.Panels(1).Text = "Ready"
    Else
        Me.sbStatusBar.Panels(1).Text = StatusText
    End If
End Sub

Private Sub iForm_AddNew()
    'n.a
End Sub


Private Sub iForm_Cancel()
    'n.a
End Sub

Private Sub iForm_CloseMe()
    End
End Sub

Private Sub iForm_delete()
    'n.a
End Sub

Private Sub iForm_DeleteRow()
    'n/a
End Sub


Private Sub iForm_Find(ByVal Key As String)
    'n.a
End Sub


Private Sub iForm_MainMenu()
    mtMenus.Visible = Not mtMenus.Visible
End Sub

Private Function iForm_OpenDB() As Boolean
    'n/a
End Function


Private Sub iForm_PrintOut()
    'n/a
End Sub

Private Sub iForm_Refresh()
    'n.a
End Sub

Private Function iForm_Save() As Boolean
    'n.a
End Function

Private Sub iForm_ShowFormView()
    'n.a
End Sub


Private Sub MDIForm_Load()
    
    'On Error GoTo Err_MDIForm_Load
    
 '   Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
 '   Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
 '   Me.width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
 '   Me.height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    sbStatusBar.Font.Name = "Tahoma"
    sbStatusBar.Font.Bold = True
    'sbStatusBar.Panels(1).Text = Format(Date, "medium date")
    Dim sSQL As String
    'LoadMenus
    
    sbStatusBar.Panels(1).Text = "(C)XBasicPro 2005 - " & Year(Date)
 '   Exit Sub

'Err_MDIForm_Load:
'        ErrorMsg Err.Number, Err.Description, "MDIForm_Load", mcstrMod
'    Exit Sub
End Sub


Private Sub MDIForm_Resize()
  On Error Resume Next
  mtMenus.Height = Me.Height - 1200
  
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
        'On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub


Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable To display Help Contents. There is no Help associated With this project.", vbInformation, Me.Caption
    Else
        'On Error Resume Next
        'nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable To display Help Contents. There is no Help associated With this project.", vbInformation, Me.Caption
    Else
        'On Error Resume Next
        'nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuBLR_Click()

End Sub

Private Sub mnuBLLabaRugi_Click()
  
  Dim oConn As New ADODB.Connection
  Dim ocmd As New ADODB.Command
  
  oConn.CursorLocation = adUseServer
  oConn.open sConn
  ocmd.ActiveConnection = oConn
  ocmd.CommandType = adCmdStoredProc
  ocmd.CommandText = "USP_BUAT_LAP_LABA_RUGI"
  ocmd.Execute
  
  Set oConn = Nothing
  Set ocmd = Nothing
      
  'Load frmLabaRugi
  'frmLabaRugi.Show
  
End Sub

Private Sub mnuBLNeraca_Click()
  'frmNeraca.Show
End Sub

Private Sub mnuBukaLaci_Click()
  MSComm1.CommPort = 1
  'Sets the Com Port number (this can be changed here)
  MSComm1.Settings = "9600,n,8,1"
  'Sets the Baud rate (9600 in this case)
  MSComm1.PortOpen = True
  'Sends a pulse to the Com Port
  MSComm1.Output = Chr$(27) + "p" + "0" + "zz"
  MSComm1.PortOpen = False
End Sub

Private Sub mnuHelpAbout_Click()
  
End Sub



Private Sub mnuGantiPassword_Click()
 frmGantiPassword.Show vbModal
End Sub

Private Sub mnuGantiPemakai_Click()
 frmLogin.Show vbModal
End Sub

Private Sub mnuHelp_Click()
Dim sPsn As String
  
  sPsn = "(C) XBasicPro 2005 - " & Year(Date) & vbCrLf & vbCrLf
  sPsn = sPsn & "Jl. Denai Gg. Muslimin 35 Medan" & vbCrLf
  sPsn = sPsn & "Phone: (061)7348086 - HP:08136 222 0708" & vbCrLf
  sPsn = sPsn & "E-mail: hartoto_d@yahoo.com" & vbCrLf
  sPsn = sPsn & "Website: http://x-basicpro.cjb.net" & vbCrLf
  sPsn = sPsn & "M E D A N  -  S U M A T E R A  U T A R A"
  MsgBox sPsn, vbInformation, "Stock"
End Sub

Private Sub mnuKeluar_Click()
End

End Sub



Private Sub mnuLtkMKanan_Click()
 picMenu.Align = vbAlignRight
End Sub

Private Sub mnuLtkMKiri_Click()
 picMenu.Align = vbAlignLeft
End Sub

Private Sub mnuVMenu_Click()
  'mtMenus.Visible = True
  'picMenu.Visible = True
  
  mnuVMenu.Checked = Not mnuVMenu.Checked
  mtMenus.Visible = mnuVMenu.Checked
  picMenu.Visible = mnuVMenu.Checked
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Sub LoadMenus()
  
  'On Error Resume Next
  
  Dim i As Integer, oNode As IXMLDOMNode, j As Integer
  Dim nodX As Node
  Dim sTeks As String, sImage As String, sTag As String
  Dim sKey As String
  Dim sTeksSub As String, sImageSub As String, sTagSub As String
  Dim sKeySub As String
  
  'MsgBox "sUserGroupName = " & sUserGroupName
  If sUserGroupName = "Administrators" Then
   m_xmlFile = App.Path & "\menu_01.xml"
  ElseIf sUserGroupName = "Supervisors" Then
    m_xmlFile = App.Path & "\menu_02.xml"
  ElseIf sUserGroupName = "Operators" Then
    m_xmlFile = App.Path & "\menu_03.xml"
  Else
    MsgBox "No group name '" & sUserGroupName & "'", vbCritical, "Error"
  End If
  
  Set oDom = New MSXML.DOMDocument
  oDom.Load m_xmlFile
  If Trim(oDom.xml) <> "" Then
    Do
     DoEvents
    Loop Until oDom.readyState = XML_LOAD_COMPLETE
  Else
   'GoTo XMLError
  End If
  
  mtMenus.MenuTreeView.Nodes.Clear
  'MsgBox oDOM.childNodes(0).childNodes.Length
  For i = 0 To oDom.childNodes(0).childNodes.length - 1
    sTeks = oDom.childNodes(0).childNodes(i).Attributes(0).Text
    sTag = oDom.childNodes(0).childNodes(i).Attributes(1).Text
    sImage = oDom.childNodes(0).childNodes(i).Attributes(2).Text
    sKey = oDom.childNodes(0).childNodes(i).Attributes(3).Text
    
    'MsgBox sTag & " " & sKey
      
  ' Configure TreeView
    With mtMenus.MenuTreeView
        .ImageList = img
        .Sorted = False
        .LabelEdit = False
        .LineStyle = tvwRootLines
        .Style = tvwTreelinesPlusMinusPictureText
        .Indentation = 10
         'Set mNode = .Nodes.Add(, , sKey, sTeks, sImage)
         .Nodes.Add , , sKey, sTeks, sImage
        .Nodes.Item(i + 1).Tag = sTag
    End With
    
    'configure menu item
    If oDom.childNodes(0).childNodes(i).hasChildNodes = True Then
      For j = 0 To oDom.childNodes(0).childNodes(i).childNodes.length - 1
        sTeksSub = oDom.childNodes(0).childNodes(i).childNodes(j).Attributes(0).Text
        sTagSub = oDom.childNodes(0).childNodes(i).childNodes(j).Attributes(1).Text
        sImageSub = oDom.childNodes(0).childNodes(i).childNodes(j).Attributes(2).Text
        sKeySub = oDom.childNodes(0).childNodes(i).childNodes(j).Attributes(3).Text

        'Set nodX = tvw.Nodes.Add(sKey, tvwChild)
        Set nodX = mtMenus.MenuTreeView.Nodes.Add(sKey, tvwChild)
        nodX.Key = sKeySub
        nodX.Text = sTeksSub
        nodX.Image = sImageSub
        nodX.Tag = sTagSub
      Next j
    End If
  Next i
          
          
sbStatusBar.Panels(1).Text = "(C)XBasicPro 2005 - " & Year(Date) & " |  " & _
  sUserName & " : " & sUserGroupName

Exit Sub

XMLError:
  MsgBox "File '" & m_xmlFile & "' tidak ada atau tdk valid", vbCritical, "LoadMenus"
  
End Sub

Private Sub mtMenus_CloseMe()
    mtMenus.Visible = False
    picMenu.Visible = False
    'With Me.tlBar.Buttons("Main")
    '    .Value = tbrUnpressed
    'End With
End Sub

Private Sub mtMenus_NodeClick(ByVal Node As MSComctlLib.Node)
        
'On Error Resume Next
Screen.MousePointer = vbHourglass

If Not (Node Is Nothing) Then
 Select Case Node.Key
  Case "P1.1"
     GoSub UnloadAllForm
     'frmBrgBidSrch.sXMLFile = "BrgBid.xml"
     'frmBrgBidSrch.Show
     'frmBrgBidSrch.ZOrder 0
    
  Case "P1.2" 'stok
     GoSub UnloadAllForm
               
   Case "P3.8"
     GoSub UnloadAllForm
     ShowReport "USP_STOCKS_SALDO_REPORT", "rptStocksSaldo.rpt"

   
   Case "P6.2"
        
        
   Case "P8.4"
        Dim kt As String, kt2 As String, s As String
        Dim oXML As DOMDocument
        
        kt = Chr(34)
        kt2 = Chr(34) + Chr(34)
        
        'buat data XML kosong
        s = "<M TglAwal=" & kt & Date & kt & Space(1)
        s = s & "TglAkhir=" & kt & Date & kt & Space(1)
        s = s & "Terima=" & kt2 & Space(1)
        s = s & "StatusCair=" & kt2 & ">"
        s = s & "<D TipeAlatBayar=" & kt2 + Space(1)
        s = s & "NoWarkat=" & kt2 & Space(1)
        s = s & "JthTempo=" & kt2 & Space(1)
        s = s & "Nominal=" & kt2 & Space(1)
        s = s & "StatusCair =" & kt2 & Space(1)
        s = s & "TglCair=" & kt2 & Space(1) & "/>"
        s = s & "</M>"
        
        Set oXML = New DOMDocument
        oXML.loadXML s
        
        'GoSub UnloadAllForm
        'frmWarkatCair.WarkatCair1.XMLFile = "WarkatCair.xml"
        'Set frmWarkatCair.WarkatCair1.oXMLData = oXML
        Set oXML = Nothing
        'frmWarkatCair.WarkatCair1.InitData
        'frmWarkatCair.Show
     
  Case "P10.1"
     'GoSub UnloadAllForm
     ShowReport "USP_REPORT_SUPPLIERS", "rptSuppliers.rpt"
     GoSub UnloadAllForm
     frmRptXML.sXMLFile = App.Path & "\CBTransPerCust_reports.xml"
     frmRptXML.Show
          
   Case "P10.3"
     'GoSub UnloadAllForm
     frmRpt.sXMLFile = App.Path & "\STOCKS_REPORTS.xml"
     frmRpt.Show
        
  Case "P10.4"
     'GoSub UnloadAllForm
     ShowReport "USP_REPORT_MUTASI_STOCKS", "rptMutasiStocks.rpt"
     GoSub UnloadAllForm
     frmRptXML.sXMLFile = App.Path & "\CBTransPerAcc_reports.xml"
     frmRptXML.Show
     
 End Select
    
    'MsgBox Node.Key & " " & Node.Tag
    
End If
Screen.MousePointer = vbDefault

Exit Sub

UnloadAllForm:
     
     Unload frmCabangSrch
     Unload frmRuangSrch
     
     'Unload frmBrgSubBidSrch
     Unload frmBrgSubBid
     'Unload frmBrgSubBidSrch
     
     'Unload frmBrgSubSubBidSrch
     Unload frmBrgSubSubBid
     'Unload frmBrgSubSubBidSrch
     
     Unload frmBrgBidEdit
     Unload frmBrgBidSrch
     
     Unload frmStocksTrOutSearch
     Unload frmStocksTrOut
     Unload frmStocksTrOutEditD
     
     Unload frmStocksTrInSearch
     'Unload frmStocksTrIn
     Unload frmStocksTrInEditD
     
     'Unload frmLokasiSearch
     
     Unload frmStocksSearch
     'Unload frmStocks
     
'     Unload frmRpt
'     Unload frmRptXML
'
       
Return




End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

