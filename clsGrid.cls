VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public oColNames  As New Collection
Public colCount As Integer

Private m_strColName As String
'

Sub AddCol()
  oColNames.Add colName
  colCount = oColNames.Count
End Sub
Public Property Get colName() As String
  colName = m_strColName
End Property

Public Property Let colName(ByVal sNewValue As String)
  m_strColName = sNewValue
End Property

Sub WriteCols()
  
End Sub
