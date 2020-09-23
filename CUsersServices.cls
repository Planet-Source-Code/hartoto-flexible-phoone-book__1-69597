VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUsersServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"431CF9190109"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=431CF946038A
Private m_sUserName As String

'##ModelId=431CF9650290
Private m_sPassword As String

'##ModelId=431CFA080000
Private bEdit As Boolean

'##ModelId=431CFA18036B
Private bAdd As Boolean

'##ModelId=431CFA2800CB
Private bDelete As Boolean

'##ModelId=431CFE9802FD
Private bLaunch As Boolean

'##ModelId=431CFF6402DE
Private bAdmin As Boolean

'##ModelId=431CFFAB005D
Private bValidUser As Boolean

'##ModelId=431CFFD60119
Private ModulName As String

'##ModelId=431D021B0196
Private m_iGroupID As Integer

Private m_iPersonnelID As Integer
Private m_sGroupName As String
'


'memeriksa keabsahan pemakai user atau password
'##ModelId=431D000D03D8
Public Function GetUserValidation(m_sUserName, m_sPassword) As Boolean
Attribute GetUserValidation.VB_Description = "memeriksa keabsahan pemakai user atau password"
    '## Your code goes here ...
    
    Dim oConn As New ADODB.Connection
    Dim oRsValidUser As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT PERSONNEL_ID, USER_NAME, PASSWORD FROM USR_PERSONNELS " & _
     "WHERE USER_NAME ='" & m_sUserName & "'" & _
     " AND PASSWORD ='" & m_sPassword & "'"
     
    'Set oRsValidUser = getReadOnlyRS(sSQL)
    oConn.open sConn
    oRsValidUser.open sSQL, oConn
    
    If Not (oRsValidUser.EOF) Then
       GetUserValidation = True
       m_iPersonnelID = oRsValidUser.Fields("PERSONNEL_ID")
       m_iGroupID = GetUserGroupsID(m_iPersonnelID)
       m_sGroupName = GetUserGroupsName(m_iGroupID)
       
     Else
       GetUserValidation = False
    End If
    
    Set oRsValidUser = Nothing
    Set oConn = Nothing
    
End Function
Public Function GantiPassword(ByVal m_sUserName As String, _
 ByVal sNewPwd As String) As Boolean
 
  Dim sSQL As String
  sSQL = "UPDATE USR_PERSONNELS " & _
     " SET PASSWORD ='" & sNewPwd & "'" & _
     " WHERE USER_NAME ='" & m_sUserName & "'"
  
  UpdateRS (sSQL)
  

End Function
'memeriksa hak user dengan masukan nama modul dan kembalian ke property bEdit,bAdd,bLaunch,bDelete
'##ModelId=431D00DB00DA
Public Sub GetPermission()
Attribute GetPermission.VB_Description = "memeriksa hak user dengan masukan nama modul dan kembalian ke property bEdit,bAdd,bLaunch,bDelete"
    '## Your code goes here ...
End Sub

'memeriksa groupID user, dari Group ini dapat ditemtukan hak-hak User hasilnya dikembalikan ke property GroupID
'##ModelId=431D01840232
Private Function GetUserGroupsID(m_iPersonnelID) As Integer
Attribute GetUserGroupsID.VB_Description = "memeriksa groupID user, dari Group ini dapat ditemtukan hak-hak User hasilnya dikembalikan ke property GroupID"
 '## Your code goes here ...
 Dim oRsGrp As New ADODB.Recordset
 Dim sSQL As String
    
 sSQL = "SELECT GROUP_ID FROM USR_GROUPS_PERSONNELS " & _
  "WHERE PERSONNEL_ID = " & m_iPersonnelID
  
 Set oRsGrp = getReadOnlyRS(sSQL)
 
 If Not (oRsGrp.EOF) Then
   GetUserGroupsID = oRsGrp.Fields("GROUP_ID")
 Else
   
 End If
    
 Set oRsGrp = Nothing
    
End Function
Private Function GetUserGroupsName(m_iGroupID) As String
 '## Your code goes here ...
 Dim oRsGrp As New ADODB.Recordset
 Dim sSQL As String
    
 sSQL = "SELECT GROUP_NAME FROM USR_GROUPS " & _
  "WHERE GROUP_ID = " & m_iGroupID
  
 Set oRsGrp = getReadOnlyRS(sSQL)
 
 If Not (oRsGrp.EOF) Then
   GetUserGroupsName = oRsGrp.Fields("GROUP_NAME")
 Else
   
 End If
    
 Set oRsGrp = Nothing
    
End Function

Public Property Get UserName() As String
  UserName = m_sUserName
End Property

Public Property Let UserName(ByVal vNewValue As String)
 m_sUserName = vNewValue
 
End Property

Public Property Get Password() As String
 Password = m_sPassword
End Property

Public Property Let Password(ByVal vNewValue As String)
 m_sPassword = vNewValue
End Property
Public Property Get PersonnelID() As Integer
 PersonnelID = m_iPersonnelID
End Property

Public Property Get GroupName() As String
 GroupName = m_sGroupName
End Property

