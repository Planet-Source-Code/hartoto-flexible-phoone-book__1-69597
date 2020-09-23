Attribute VB_Name = "Global"
Option Explicit

Public sActiveUserGroup As String

Public Sub Main()
    
    'CekSerial

    'Open App.Path & "\conn.udl" For Input As #1
    'Do While Not EOF(1)
    ' Input #1, sConn
    'Loop
    'Close #1
    'MsgBox sConn
    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=super123"
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = sConn
    conn.open
    MainFrm.Show
    LoginForm2.Show vbModal
    
End Sub

Public Sub Fillcombo(ctl As ComboBox, ByVal ComboList As String)
    On Error Resume Next
    Dim strAll() As String, i&
    strAll = Split(ComboList, "|")
    For i = 0 To UBound(strAll)
        ctl.AddItem strAll(i)
    Next
End Sub

Public Function GetComboValue(ctl As ComboBox, ByVal ComboList As String)
    On Error Resume Next
    Dim strAll() As String, strRow() As String
    strAll = Split(ComboList, "|")
    strRow = Split(strAll(ctl.ListIndex), vbTab)
    GetComboValue = strRow(0)
End Function

Function ErrorMsg(errNum As Long, ErrDesc As String, _
    strFunction As String, strModule As String)
    On Error Resume Next
    Dim anErrorMessage As String
    anErrorMessage = "Error Number: " & errNum & "." & vbCrLf & _
        "Error Description: " & ErrDesc & vbCrLf & _
        "Module Name: " & strModule & vbCrLf & _
        "Sub/Function: " & strFunction & vbCrLf
    MsgBox anErrorMessage, vbCritical
End Function
Function FixDate(dDate As Date) As String
  
  '''''''''''''''''''''''''''''''''''''
  'pengubahan tgl utk sql server
  'membuat format mm/dd/yyy
  '''''''''''''''''''''''''''''''''
  Dim dd As String, mm As String, yyyy As String
  Dim sDate As String
  
  mm = CStr(Month(dDate))
  dd = CStr(Day(dDate))
  yyyy = CStr(Year(dDate))
  
  sDate = mm & "/" & dd & "/" & yyyy
  
  FixDate = sDate
  
  'MsgBox dDate & " " & sDate, , "FixDate"
    
End Function
Sub ShowDialog(sDialog As String)
  DialogFrm.lblKomentar.Caption = sDialog
  DialogFrm.Show vbModal
End Sub
