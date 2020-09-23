VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.MDIForm MainFrm 
   BackColor       =   &H00FFC0C0&
   Caption         =   "XB-Super Data"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8250
   Icon            =   "MainFrm.frx":0000
   Picture         =   "MainFrm.frx":164A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSCommLib.MSComm MSComm1 
      Left            =   720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
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
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
      EndProperty
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
      Begin XBSuperData.MenuTree mtMenus 
         Height          =   6255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   11033
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
            Picture         =   "MainFrm.frx":6F59
            Key             =   "g1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":7833
            Key             =   "kasbank"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":953D
            Key             =   "hp"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":A38F
            Key             =   "laporan"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":B1E1
            Key             =   "hutang"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":C033
            Key             =   "g3"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":CE85
            Key             =   "g4"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":DCD7
            Key             =   "g5"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":EB29
            Key             =   "gl"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":10833
            Key             =   "BookOpen"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":10C85
            Key             =   "pembelian"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":11AD7
            Key             =   "g8"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":137E1
            Key             =   "g10"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":13953
            Key             =   "BookPink"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":13DA5
            Key             =   "g2"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":14BF7
            Key             =   "ProsesHPP"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuKeluar 
      Caption         =   "&Exit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Tampilan"
      Visible         =   0   'False
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuVMenu 
         Caption         =   "Menu"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuLtkMenu 
      Caption         =   "&View Menu"
      Begin VB.Menu mnuLtkMKanan 
         Caption         =   "&Right"
      End
      Begin VB.Menu mnuLtkMKiri 
         Caption         =   "&Left"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHPetunjuk 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuHAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Private m_EnableAttr As ToolBarItems
'Private Const mcstrMod$ = "frmMain"
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
    
  
    'sbStatusBar.Panels(1).Text = Format(Date, "medium date")
    Dim sSQL As String
    'LoadMenus
    
    sbStatusBar.Panels(1).Text = "(C)XBasicPro 2005 - " & Year(Date)
    sbStatusBar.Panels(2).Text = Format(Date, "dd/mm/yy")
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



Private Sub mnuAbout_Click()

End Sub


Private Sub mnuGantiPemakai_Click()
 LoginForm2.Show vbModal
End Sub

Private Sub mnuHAbout_Click()
Dim sPsn As String
  
  sPsn = "(C) XBasicPro 2005 - " & Year(Date) & vbCrLf & vbCrLf & vbCrLf
  sPsn = sPsn & "IT Solutios " & vbCrLf
  sPsn = sPsn & "Program: Stock, GL, Website dll." & vbCrLf
  sPsn = sPsn & "Private Less: "
  sPsn = sPsn & "Visual Basic, VB.Net, ASP, ASP.NET" & vbCrLf & vbCrLf
  sPsn = sPsn & "Jl. Denai No.92 Medan" & vbCrLf
  sPsn = sPsn & "Phone: 77374922 - HP:0813 7677 2820" & vbCrLf
  sPsn = sPsn & "E-mail: hartoto_d@yahoo.com" & vbCrLf
  sPsn = sPsn & "M E D A N  -  S U M A T E R A  U T A R A"
  MsgBox sPsn, vbInformation
End Sub



Private Sub mnuHPetunjuk_Click()
Load HelpFrm
HelpFrm.WebBrowser1.Navigate App.Path & "\help.htm"
HelpFrm.Show
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
       'MsgBox oDom.xml
       SrchFrm.sXMLFile = "config.xml"
       SrchFrm.Show
       SrchFrm.ZOrder 0
     
     
  Case "P10.1"
     'GoSub UnloadAllForm
 '    ShowReport "USP_REPORT_SUPPLIERS", "rptSuppliers.rpt"
 '    GoSub UnloadAllForm
     'frmRptXML.sXMLFile = App.Path & "\CBTransPerCust_reports.xml"
     'frmRptXML.Show
          
     
 End Select
    
    'MsgBox Node.Key & " " & Node.Tag
    
End If
Screen.MousePointer = vbDefault

Exit Sub

UnloadAllForm:
    Unload SrchFrm
     
    
Return




End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

