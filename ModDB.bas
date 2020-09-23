Attribute VB_Name = "ModDB"
Public sUserGroupName As String
Public sUserName As String

Public EditMode As Long

Public Enum lEditMode
  Add = 0
  Edit = 1
  Read = 2
End Enum

Public conn As ADODB.Connection
Public RS As ADODB.Recordset
Public ocmd As ADODB.Command

'Public Const sConn As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=D:\stockPabrik\stockProd.mdb"
'Public Const sConn As String = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=SA;Initial Catalog=StockProd;Data Source=XB"
'Public Const sConn As String = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=StockProd;Data Source=w2kserver"
'
Global sConn As String
Function AngkaToRomawi(ByVal iAngka As Integer) As String
    Dim angka As Variant
    Romawi = Array("-", "I", "II", "II", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII")
    AngkaToRomawi = Romawi(iAngka)
End Function
Function ToTerbilangX(bilangan) As String

Dim bilratusjutaan
Dim bilMiliaran
Dim belasribu
Dim uangsatuan
Dim uangpuluhan
Dim uangratusan
Dim RatusanJuta As Integer
Dim uangribuan
Dim uangpuluhribuan
Dim uangratusribuan
Dim uangjutaan
Dim uangpulujutaan
Dim UangRatusJutaan
Dim UangMiliaran
'Dim bilangan
Dim Uang
bilsatuan = ""
bilpuluhan = ""
bilratusan = ""
bilribuan = ""
bilpuluhribuan = ""
bilratusribuan = ""
biljutaan = ""
bilpuluhjutaan = ""
bilratusjutaan = ""
bilMiliaran = ""
Uang = bilangan

If bilangan >= 1000000000 Then
   UangMiliaran = Int(bilangan / 1000000000)
   bilMiliaran = Satuan(UangMiliaran) + " Miliar "
   bilangan = bilangan Mod 1000000000
End If

If bilangan >= 100000000 Then
    UangRatusJutaan = bilangan
    RatusanJuta = Int(UangRatusJutaan / 1000000)
    If RatusanJuta Mod 100 = 0 Then
     bilratusjutaan = Ratusan(RatusanJuta) + " Juta "
    Else
     bilratusjutaan = Ratusan(RatusanJuta)
    End If
    bilangan = bilangan - (Int(RatusanJuta / 100) * 100000000)
End If

If bilangan >= 10000000 Then
    uangpuluhjutaan = bilangan
    puluhanjuta = Int(uangpuluhjutaan / 1000000)
    If uangpuluhjutaan < 20000000 Then
     If uangpuluhjutaan >= 11000000 Then
         bilpuluhjutaan = belasan(puluhanjuta)
         bilangan = bilangan - (1000000 * bilpuluhjutaan1)
      End If
    Else
      bilpuluhjutaan = puluhan(puluhanjuta) + Satuan((puluhanjuta Mod 10)) + " Juta "
      bilangan = bilangan Mod 1000000
    End If
End If
If bilangan >= 1000000 Then
   uangjutaan = Int(bilangan / 1000000)
   biljutaan = Satuan(uangjutaan) + " Juta "
   bilangan = bilangan Mod 1000000
End If
If bilangan >= 100000 Then
    uangratusribuan = bilangan
    ratusanribu = Int(uangratusribuan / 1000)
    If ratusanribu Mod 100 = 0 Then
     bilratusribuan = Ratusan(ratusanribu) + " Ribu "
    Else
     bilratusribuan = Ratusan(ratusanribu)
    End If
    bilangan = bilangan - (Int(ratusanribu / 100) * 100000)
End If


If bilangan >= 10000 Then
    uangpuluhribuan = bilangan
    puluhanribu = Int(uangpuluhribuan / 1000)
    If uangpuluhribuan < 20000 And uangpuluhribuan >= 11000 Then
         bilpuluhribuan = belasan(puluhanribu)
    Else
     bilpuluhribuan = puluhan(puluhanribu) + Satuan((puluhanribu Mod 10)) + " Ribu "
      bilangan = bilangan Mod 1000
    End If
End If
If bilangan >= 1000 Then
   uangribuan = bilangan
   bilribuan = ribuan(uangribuan, Uang)
   bilangan = bilangan Mod 1000
End If
If bilangan >= 100 Then
  uangratusan = bilangan
  bilratusan = Ratusan(uangratusan)
  bilangan = bilangan Mod 100
End If
If bilangan > 10 Then
  If bilangan < 20 Then
    bilpuluhan = belasan(bilangan)
    bilangan = ""
  Else
    uangpuluhan = bilangan
    bilpuluhan = puluhan(uangpuluhan)
    bilangan = bilangan Mod 10
  End If
End If
If bilangan >= 0 Then
  bilsatuan = bilangan
End If
ToTerbilangX = bilMiliaran + bilratusjutaan + bilpuluhjutaan + biljutaan + bilratusribuan + bilpuluhribuan + bilribuan + bilratusan + bilpuluhan + Satuan(bilsatuan)
End Function
Function belasan(belas)
   belas = belas - 10
   If belas = 1 Then
     belasan = "Sebelas"
   Else
     belasan = Satuan(belas) + " Belas"
   End If
End Function
Function puluhan(puluh)
   If puluh = 10 Then
     puluhan = " Sepuluh"
     Else
     puluhan = Satuan(Int(puluh / 10)) + " Puluh "
   End If
End Function
Function Ratusan(ratus) As String
If (Int(ratus / 100)) = 1 Then
   Ratusan = " Seratus "
 Else
   Ratusan = Satuan(Int(ratus / 100)) + " Ratus "
 End If
End Function
Function ribuan(ribu, Uang)
 Dim bilribu
 bilribu = Right((Str(Uang)), 6)
 bilribu = Int(bilribu)
 
 
 If (Int(bilribu / 100000) > 0) Or (Uang < 19999 And Uang > 9999) Then
   ribuan = Satuan(Int(ribu / 1000)) + " Ribu "
 Else
 If (Int(ribu / 1000)) = 1 And (Uang < 9999 Or Uang > 999999) Then
   ribuan = " Seribu "
 Else
   ribuan = Satuan(Int(ribu / 1000)) + " Ribu "
 End If
 End If
 
End Function
Private Function Satuan(x) As String
    Select Case x
    Case 0:   Satuan = ""
    Case 1:   Satuan = "Satu"
    Case 2:   Satuan = "Dua"
    Case 3:   Satuan = "Tiga"
    Case 4:   Satuan = "Empat"
    Case 5:   Satuan = "Lima"
    Case 6:   Satuan = "Enam"
    Case 7:   Satuan = "Tujuh"
    Case 8:   Satuan = "Delapan"
    Case 9:   Satuan = "Sembilan"
    Case 10:  Satuan = " Sepuluh"
    End Select
End Function

Function getCounterRS(ByVal sTblCounter As String) As ADODB.Recordset
 
 'On Error GoTo localErr
 
 Dim sSQL As String
 Dim oConn As New ADODB.Connection
 oConn.open sConn
 
 Dim rsTemp As ADODB.Recordset
 Set rsTemp = New ADODB.Recordset
 
 sSQL = "SELECT NO_URUT FROM " & sTblCounter
 
 'MsgBox sSQL, , "ModDB.getCouterRS"
 
 rsTemp.open sSQL, oConn, adOpenStatic
 Set getCounterRS = rsTemp
 Set rsTemp = Nothing
 Set oConn = Nothing
 
 Exit Function
 
localErr:
 MsgBox Err.Number & " " & Err.Description
 writeError Err.Number & " " & Err.Description
 Resume Next
 
End Function
Function getCounterMDRS(ByVal sTblCounter As String, sNoVoucher As String) As ADODB.Recordset
 
 On Error GoTo localErr
 
 Dim sSQL As String
 Dim rsTemp As ADODB.Recordset
 Set rsTemp = New ADODB.Recordset
 
 sSQL = "SELECT NO_URUT FROM " & sTblCounter & _
  " WHERE KODE_VOUCHER = '" & sNoVoucher & "'"
 
 'MsgBox sSQL, , "ModDB.getCouterMDRS"
 
 rsTemp.open sSQL, conn, adOpenStatic
 Set getCounterMDRS = rsTemp
 Set rsTemp = Nothing
 
 Exit Function
 
localErr:
 MsgBox Err.Number & " " & Err.Description
 writeError Err.Number & " " & Err.Description
 Resume Next
 
End Function
Function FindRS(sSQL) As Boolean
 
 Dim oConn As New ADODB.Connection
 Dim rsTemp As ADODB.Recordset
 Set rsTemp = New ADODB.Recordset
 oConn.open sConn
 
 rsTemp.open sSQL, oConn, adOpenStatic
 If Not rsTemp.EOF Then
   FindRS = True
 Else
   FindRS = False
 End If
 
 Set rsTemp = Nothing
 Set oConn = Nothing
 
End Function

Function getReadOnlyRS(sSQL) As ADODB.Recordset
 
 Dim oConnTemp As New ADODB.Connection
 Dim rsTemp As ADODB.Recordset
 Set rsTemp = New ADODB.Recordset
 
 'MsgBox sSQL
 oConnTemp.open sConn
 rsTemp.open sSQL, oConnTemp, adOpenStatic
 If Not rsTemp.EOF Then
   Set getReadOnlyRS = rsTemp
 Else
   Set getReadOnlyRS = Nothing
 End If
 Set oConnTemp = Nothing
 
End Function

Function IsAdaRS(sSQL) As Boolean
 
 Dim oConnTemp As New ADODB.Connection
 Dim rsTemp As ADODB.Recordset
 Set rsTemp = New ADODB.Recordset
 
 'MsgBox sSQL
 oConnTemp.open sConn
 rsTemp.open sSQL, oConnTemp, adOpenStatic
 If Not rsTemp.EOF Then
    IsAdaRS = True
 Else
    IsAdaRS = False
 End If
 Set oConnTemp = Nothing
 
End Function

Function UpdateRS(ByVal sSQL As String) As Boolean
  Dim rsUpdate As New ADODB.Recordset
  rsUpdate.open sSQL, conn, adOpenDynamic, adLockPessimistic
  'MsgBox "rsUpdate.State = " & rsUpdate.State
  If rsUpdate.State = adRecModified Then
    UpdateRS = True
  End If
  Set rsUpdate = Nothing
End Function
Function getAutoNum(ByVal oRs As ADODB.Recordset) As Long
  On Error GoTo localErr
  
  Dim i As Long
  
  i = oRs.RecordCount + 1
  getAutoNum = i
  'MsgBox i, , "getAutoNum"
  
Exit Function

localErr:
 MsgBox Err.Number & " " & Err.Description, vbCritical, "ModDB.getAutoNum"
 writeError Err.Number & " " & Err.Description & " ModDB.getAutoNum"
 Resume Next
 
End Function
Function getAutoTranCode(sTrans As String, lCount As Long) As String
  
  Dim s As String
  
  s = sTrans & "." & CStr(Year(Date)) & "."
  If Len(CStr(Month(Date))) = 1 Then
     s = s & "0" & CStr(Month(Date))
   Else
    s = s & CStr(Month(Date))
  End If
  
  If Len(CStr(Day(Date))) = 1 Then
    s = s & "." & "0" & CStr(Day(Date))
  Else
    s = s & "." & CStr(Day(Date))
  End If
  
  If Len(CStr(lCount)) = 1 Then
   s = s & "." & "00000" & CStr(lCount)
  End If
  If Len(CStr(lCount)) = 2 Then
   s = s & "." & "0000" & CStr(lCount)
  End If
  If Len(CStr(lCount)) = 3 Then
   s = s & "." & "000" & CStr(lCount)
  End If
  If Len(CStr(lCount)) = 4 Then
   s = s & "." & "00" & CStr(lCount)
  End If
  If Len(CStr(lCount)) = 5 Then
   s = s & "." & "0" & CStr(lCount)
  End If
  If Len(CStr(lCount)) = 6 Then
   s = s & "." & CStr(lCount)
  End If
  
  getAutoTranCode = s
  
End Function
Function getAutoTranCodeShort(sTrans As String, lCount As Long) As String
  
  Dim s As String
  
  s = sTrans
  s = Trim(sTrans)
  
  If Len(CStr(lCount)) = 1 Then
   s = s & "00000" & CStr(lCount)
  End If
  If Len(CStr(lCount)) = 2 Then
   s = s & "." & "0000" & CStr(lCount)
  End If
  If Len(CStr(lCount)) = 3 Then
   s = s & "." & "000" & CStr(lCount)
  End If
  If Len(CStr(lCount)) = 4 Then
   s = s & "." & "00" & CStr(lCount)
  End If
  If Len(CStr(lCount)) = 5 Then
   s = s & "." & "0" & CStr(lCount)
  End If
  If Len(CStr(lCount)) = 6 Then
   s = s & "." & CStr(lCount)
  End If
  
  getAutoTranCodeShort = s
  
End Function

Sub writeError(sError As String)
 Dim sFile As String
 sFile = App.Path & "\error.log"
 'MsgBox sFile
 
 Open sFile For Append As #1
 Print #1, Date & " " & Time & " " & sError
 Close #1
End Sub
Sub ClearDBSemX()
 
 Dim cn As New ADODB.Connection
 Dim rsSem As New ADODB.Recordset
 Dim rsSemCount As New ADODB.Recordset
 
 cn.open sConn
 Dim sSQLCount As String
 
 rsSem.open "DELETE FROM PENJUALAN_D_SEM", cn, adOpenDynamic
 rsSem.open "DELETE FROM PENJUALAN_D_SEM_COUNTER", cn, adOpenDynamic
 rsSem.open "DELETE FROM PEMBELIAN_D_SEM", cn, adOpenDynamic
 rsSem.open "DELETE FROM PEMBELIAN_D_SEM_COUNTER", cn, adOpenDynamic
 
 Set rsSem = Nothing
 Set cn = Nothing
End Sub
Sub ShowReport(ByVal sSProc As String, ByVal sRpt As String)
     
Dim oConn As New ADODB.Connection
Dim ocmd As New ADODB.Command

oConn.CursorLocation = adUseServer
oConn.open sConn
ocmd.ActiveConnection = oConn
ocmd.CommandType = adCmdStoredProc
ocmd.CommandText = sSProc
Set frmRptViewer.RS = ocmd.Execute
frmRptViewer.rptName = sRpt
frmRptViewer.Show

Set ocmd = Nothing
Set oConn = Nothing

End Sub
Sub ShowReport1Par(ByVal sSProc As String, ByVal sRpt As String, _
  ByVal sParName As String, ByVal sParValue As String)
     
Dim oConn As New ADODB.Connection
Dim ocmd As New ADODB.Command
Dim oPar As New ADODB.Parameter

oConn.CursorLocation = adUseClient
oConn.open sConn
Set oPar = ocmd.CreateParameter(sParName, adVarChar, adParamInput, 20, sParValue)
With ocmd
  .ActiveConnection = oConn
  .CommandType = adCmdStoredProc
  .CommandText = sSProc
  .Parameters.Append oPar
  Set frmRptViewer.RS = .Execute
End With

frmRptViewer.rptName = sRpt
frmRptViewer.Show

Set oPar = Nothing
Set ocmd = Nothing
Set oConn = Nothing

End Sub

Sub ShowReportSPHistory(ByVal sSProc As String, _
  ByVal sRpt As String, _
  ByVal sPar1Name As String, ByVal sPar1Value As String, _
  ByVal sPar2Name As String, ByVal sPar2Value As Variant, _
  ByVal sPar3Name As String, ByVal sPar3Value As Variant)
     
Dim oConn As New ADODB.Connection
Dim ocmd As New ADODB.Command
Dim oPar1 As New ADODB.Parameter
Dim oPar2 As New ADODB.Parameter
Dim oPar3 As New ADODB.Parameter

oConn.CursorLocation = adUseServer
oConn.open sConn
With ocmd
    .ActiveConnection = oConn
    .CommandType = adCmdStoredProc
    .CommandText = sSProc
    Set oPar1 = .CreateParameter(sPar1Name, adVarChar, adParamInput, 10, sPar1Value)
    Set oPar2 = .CreateParameter(sPar2Name, adDBDate, adParamInput, 8, sPar2Value)
    Set oPar3 = .CreateParameter(sPar3Name, adDBDate, adParamInput, 8, sPar3Value)
    Set frmRptViewer.RS = .Execute
End With

frmRptViewer.rptName = sRpt
frmRptViewer.Show

Set oPar1 = Nothing
Set oPar2 = Nothing
Set oPar3 = Nothing
Set ocmd = Nothing
Set oConn = Nothing

End Sub

Sub ShowReportcmdText(ByVal sSProc As String, ByVal sRpt As String)
     
Dim oConn As New ADODB.Connection
Dim ocmd As New ADODB.Command

oConn.CursorLocation = adUseServer
oConn.open sConn
ocmd.ActiveConnection = oConn
'ocmd.CommandType = adCmdStoredProc
ocmd.CommandType = adCmdText
ocmd.CommandText = sSProc
Set frmRptViewer.RS = ocmd.Execute
frmRptViewer.rptName = sRpt
frmRptViewer.Show vbModal

Set ocmd = Nothing
Set oConn = Nothing

End Sub


Sub ShowReportKomisi(ByVal sSProc As String, _
   ByVal sRpt As String, ByVal sPar As String)
     
Dim oConn As New ADODB.Connection
Dim ocmd As New ADODB.Command
Dim oPar As New ADODB.Parameter

oConn.CursorLocation = adUseServer
oConn.open sConn
ocmd.ActiveConnection = oConn
ocmd.CommandType = adCmdStoredProc
ocmd.CommandText = sSProc
Set oPar = ocmd.CreateParameter("@KodeSales", adVarChar, adParamInput, 5, sPar)
ocmd.Parameters.Append oPar

Set frmRptViewer.RS = ocmd.Execute
frmRptViewer.rptName = sRpt
frmRptViewer.Show

Set ocmd = Nothing
Set oConn = Nothing

End Sub
Sub ShowReportKomisiGrid(ByVal sSProc As String, _
   ByVal sRpt As String, ByVal sPar As String)
     
Dim oConn As New ADODB.Connection
Dim ocmd As New ADODB.Command
Dim oPar As New ADODB.Parameter

oConn.CursorLocation = adUseServer
oConn.open sConn
ocmd.ActiveConnection = oConn
ocmd.CommandType = adCmdStoredProc
ocmd.CommandText = sSProc
Set oPar = ocmd.CreateParameter("@KodeSales", adVarChar, adParamInput, 5, sPar)
ocmd.Parameters.Append oPar

Set frmKomisi.RS = ocmd.Execute
frmKomisi.Show vbModal

Set ocmd = Nothing
Set oConn = Nothing

End Sub

Sub ShowReportRS(ByVal sSQL As String, ByVal sRptName As String)
 
 Dim oConn As New ADODB.Connection
 Dim oRs As New ADODB.Recordset
 
 oConn.open sConn
 'MsgBox sSQL
 
 oRs.open sSQL, oConn

 Set frmRptViewer.RS = oRs
 frmRptViewer.rptName = sRptName
 frmRptViewer.Show

 Set oRs = Nothing
 Set oConn = Nothing
 

End Sub

Sub ShowReportWithParam(ByVal sSProc As String, _
 ByVal sPar1 As String, ByVal sPar2 As String, _
 ByVal sRpt As String)
     
Dim oConn As New ADODB.Connection
Dim ocmd As New ADODB.Command
Dim oPar1 As New ADODB.Parameter
Dim oPar2 As New ADODB.Parameter

oConn.CursorLocation = adUseServer
oConn.open sConn
ocmd.ActiveConnection = oConn

oPar1 = ocmd.CreateParameter("@TglAwal", adVarChar, adParamInput, 20, sPar1)
oPar2 = ocmd.CreateParameter("@TglAkhir", adVarChar, adParamInput, 20, sPar2)
ocmd.Parameters.Append oPar1
ocmd.Parameters.Append oPar2

ocmd.CommandType = adCmdStoredProc
ocmd.CommandText = sSProc

Set frmRptViewer.RS = ocmd.Execute
frmRptViewer.rptName = sRpt
frmRptViewer.Show

Set ocmd = Nothing
Set oConn = Nothing

End Sub

Public Sub BukaExcel(sFile As String)
  
End Sub

Sub UpdateDB_XML(SPName As String, sXML As String)

 Dim oConn As New ADODB.Connection
 Dim ocmd As New ADODB.Command
 Dim oPar As New ADODB.Parameter
 
 Set oPar = ocmd.CreateParameter("@strXML", adVarChar, adParamInput, 10000, sXML)
 ocmd.Parameters.Append oPar
 
 oConn.open sConn
 With ocmd
    .ActiveConnection = conn
    .CommandType = adCmdStoredProc
    .CommandText = SPName
    .Execute
 End With
 
 Set ocmd = Nothing
 Set oPar = Nothing
 Set oConn = Nothing

'Debug.Print sXML
'Stop

End Sub
Public Sub execSP1Par(sSPName, sParName, sParValue)
 
 Dim oConnTmp As New ADODB.Connection
 Dim ocmd As New ADODB.Command
 
 Dim oPar1 As ADODB.Parameter
 
 oConnTmp.open sConn
 
 With ocmd
  .ActiveConnection = oConnTmp
  .CommandType = adCmdStoredProc
  .CommandText = sSPName
   Set oPar1 = .CreateParameter(sParName, adLongVarChar, adParamInput, 2000, sParValue)
  .Parameters.Append oPar1
  .Execute
 End With
 
 Set oPar1 = Nothing
 Set ocmd = Nothing
 Set oConnTmp = Nothing
 
End Sub
Public Sub execSP2Par(sSPName, sPar1Name, sPar1Value, sPar2Name, sPar2Value)
 
 Dim oConnTmp As New ADODB.Connection
 Dim ocmd As New ADODB.Command
 
 Dim oPar1 As ADODB.Parameter
 Dim oPar2 As ADODB.Parameter
 
 oConnTmp.open sConn
 
 With ocmd
  .ActiveConnection = oConnTmp
  .CommandType = adCmdStoredProc
  .CommandText = sSPName
   Set oPar1 = .CreateParameter(sPar1Name, adLongVarChar, adParamInput, 2000, sPar1Value)
   Set oPar2 = .CreateParameter(sPar2Name, adLongVarChar, adParamInput, 2000, sPar2Value)
  .Parameters.Append oPar1
  .Parameters.Append oPar2
  .Execute
 End With
 
 
 Set oPar1 = Nothing
 Set oPar2 = Nothing
 Set ocmd = Nothing
 Set oConnTmp = Nothing
 
End Sub

Public Sub execSP3Par(sSPName, sPar1Name, sPar1Value, _
  sPar2Name, sPar2Value, _
  sPar3Name, sPar3Value)
 
 Dim oConnTmp As New ADODB.Connection
 Dim ocmd As New ADODB.Command
 
 Dim oPar1 As ADODB.Parameter
 Dim oPar2 As ADODB.Parameter
 Dim oPar3 As ADODB.Parameter
 
 oConnTmp.open sConn
 
 With ocmd
  .ActiveConnection = oConnTmp
  .CommandType = adCmdStoredProc
  .CommandText = sSPName
   Set oPar1 = .CreateParameter(sPar1Name, adLongVarChar, adParamInput, 2000, sPar1Value)
   Set oPar2 = .CreateParameter(sPar2Name, adLongVarChar, adParamInput, 2000, sPar2Value)
   Set oPar3 = .CreateParameter(sPar3Name, adLongVarChar, adParamInput, 2000, sPar3Value)
  .Parameters.Append oPar1
  .Parameters.Append oPar2
  .Parameters.Append oPar3
  .Execute
 End With
 
 Set oPar1 = Nothing
 Set oPar2 = Nothing
 Set oPar3 = Nothing
 Set ocmd = Nothing
 Set oConnTmp = Nothing
 
End Sub

Public Function execSP1ParToRS(sSPName, sParName, sParValue) As ADODB.Recordset
 
 Dim oConnTmp As New ADODB.Connection
 Dim ocmd As New ADODB.Command
 Dim oPar1 As New ADODB.Parameter
 
 oConnTmp.open sConn
 oConnTmp.CursorLocation = adUseClient
 Set oPar1 = ocmd.CreateParameter(sParName, adVarChar, adParamInput, 1000, sParValue)
 
 With ocmd
  .ActiveConnection = oConnTmp
  .CommandType = adCmdStoredProc
  .CommandText = sSPName
  .Parameters.Append oPar1
  Set execSP1ParToRS = .Execute
 End With
 
 Set oPar1 = Nothing
 Set ocmd = Nothing
 Set oConnTmp = Nothing
 
End Function

Public Sub ShowReportSP1Par(sSPName, sParName, sParValue, sRptName)
 
 Dim oConnTmp As New ADODB.Connection
 Dim ocmd As New ADODB.Command
 Dim oPar1 As New ADODB.Parameter
 
 oConnTmp.CursorLocation = adUseClient
 oConnTmp.open sConn
 Set oPar1 = ocmd.CreateParameter(sParName, adVarChar, adParamInput, 1000, sParValue)
 
 With ocmd
  .ActiveConnection = oConnTmp
  .CommandType = adCmdStoredProc
  .CommandText = sSPName
  .Parameters.Append oPar1
  Set frmRptViewer.RS = .Execute
End With

frmRptViewer.rptName = sRptName
frmRptViewer.Show

 Set oPar1 = Nothing
 Set ocmd = Nothing
 Set oConnTmp = Nothing
 
End Sub

Public Sub UpdateStocks()

   Dim oConn As New ADODB.Connection
   oConn.CursorLocation = adUseServer
   oConn.open sConn
   
   Dim ocmd As New ADODB.Command
   With ocmd
     .ActiveConnection = oConn
     .CommandType = adCmdStoredProc
     .CommandText = "USP_STOCKS_HITUNG"
     .Execute
   End With
   
   Set ocmd = Nothing
   Set oConn = Nothing
   
End Sub
Public Sub UpdateStocks_B(sKdBrg As String, sKdLokasi As String)

   Dim oConn As New ADODB.Connection
   oConn.CursorLocation = adUseServer
   oConn.open sConn
   
   Dim oParKdBrg As New ADODB.Parameter
   Dim oParKdLokasi As New ADODB.Parameter
   
   Set oParKdBrg = ocmd.CreateParameter("@Kode_Barang", adVarChar, adParamInput, 20, sKdBrg)
   Set oParKdLokasi = ocmd.CreateParameter("@Kode_Lokasi", adVarChar, adParamInput, 10, sKdLokasi)

   'Dim ocmd As New ADODB.Command
   With ocmd
     .ActiveConnection = oConn
     .CommandType = adCmdStoredProc
     .CommandText = "USP_STOCKS_HITUNG"
     .Parameters.Append oParKdBrg
     .Parameters.Append oParKdLokasi
     .Execute
   End With
   
   Set oParKdBrg = Nothing
   Set oParKdLokasi = Nothing
   Set ocmd = Nothing
   Set oConn = Nothing
   
End Sub

Public Sub UpdateStocksDetails()
   
   Dim oConn As New ADODB.Connection
   oConn.CursorLocation = adUseServer
   oConn.open sConn
   
   Dim rsIn As New ADODB.Recordset
   
End Sub

Public Function ToTerbilang(ByVal x As Double) As String
  
  Dim abil As Variant
  abil = Array("", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas")
  
  If x < 12 Then
    ToTerbilang = " " & abil(x)
  ElseIf x < 20 Then
    ToTerbilang = ToTerbilang(x - 10) & " belas"
  ElseIf x < 100 Then
    ToTerbilang = ToTerbilang(x \ 10) & " puluh" & ToTerbilang(x Mod 10)
  ElseIf x < 200 Then
    ToTerbilang = " seratus" & ToTerbilang(x - 100)
  ElseIf x < 1000 Then
    ToTerbilang = ToTerbilang(x \ 100) & " ratus" & ToTerbilang(x Mod 100)
  ElseIf x < 2000 Then
    ToTerbilang = " seribu" & ToTerbilang(x - 1000)
  ElseIf x < 1000000 Then
    ToTerbilang = ToTerbilang(x \ 1000) & " ribu" & ToTerbilang(x Mod 1000)
  ElseIf x < 1000000000 Then
    ToTerbilang = ToTerbilang(x \ 1000000) & " juta" & ToTerbilang(x Mod 1000000)
  ElseIf x < 1000000000000# Then
    ToTerbilang = ToTerbilang(x \ 1000000000) & " miliar" & ToTerbilang(x Mod 1000000000)
  End If
End Function

Public Sub HPP_FIFO()
  
Dim oConn As New ADODB.Connection
Dim rsStok As New ADODB.Recordset
Dim rsBeli As New ADODB.Recordset
Dim rsJual As New ADODB.Recordset
Dim s As String
Dim QtyBeli As Double
Dim RpBeli As Double, RpJual As Double

oConn.open sConn
s = "SELECT KODE_BARANG,QTY_AKHIR,RP_AKHIR,RP_AWAL,HPP FROM STOCKS"
rsStok.open s, oConn, adOpenDynamic, adLockPessimistic

Do While Not rsStok.EOF
  '--- beli
  s = "select  kode_voucher, ISNULL(harga,0) as harga, ISNULL(qty,0) as qty "
  s = s & "from pembelian_d  where kode_barang = '" & rsStok.Fields("KODE_BARANG") & "'"
  s = s & "order by kode_voucher desc"
  rsBeli.open s, oConn, adOpenKeyset
  Do While Not rsBeli.EOF
    QtyBeli = QtyBeli + rsBeli.Fields("qty")
    RpBeli = RpBeli + (rsBeli.Fields("qty") * rsBeli.Fields("harga"))
    If QtyBeli >= rsBeli.Fields("qty") Then
      Exit Do
    End If
    rsBeli.MoveNext
  Loop
  rsBeli.Close
 
  '--- jual
  s = "select  kode_voucher, QTY, HARGA   from PENJUALAN_D "
  s = s & "where kode_barang = '" & rsStok.Fields("KODE_BARANG") & "'"
  rsJual.open s, oConn, adOpenKeyset
  Do While Not rsJual.EOF
    RpJual = RpJual + (rsJual.Fields("qty") * rsJual.Fields("harga"))
    rsJual.MoveNext
  Loop
  rsJual.Close

  rsStok.Fields("RP_AKHIR") = (rsStok.Fields("RP_AWAL") + RpBeli) - RpJual
  rsStok.Update
  
  'MsgBox rsStok.Fields("KODE_BARANG") & "  RpJual = " & RpJual
  If RpJual > 0 Then
    HPP = rsStok.Fields("RP_AWAL") + RpBeli - rsStok.Fields("RP_AKHIR")
    rsStok.Fields("HPP") = HPP
    rsStok.Update
  Else
    HPP = 0
    rsStok.Fields("HPP") = HPP
    rsStok.Update
  End If
  
  rsStok.MoveNext
Loop

  
  
End Sub
