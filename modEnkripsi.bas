Attribute VB_Name = "modEnkripsi"
'Encryption Module (Crypt.mod)
'
' When I wrote this encryption module I had a couple of goals:
'  1. Must be at least moderately tough to crack
'  2. All characters must be able to be displayed by Windows
'     (For sending via e-mail and using copy/paste to decrypt)
'
' Usage:
' Useing these functions is as easy as you'd assume it'd be,
' simply call the function as follows...
'   ReturnString = [De|En]cryptString("Text to be de/encrypted", "Crypt Key")
'
' This will encrypt any string that you can type on the
' (English) keyboard.  Extended ASCII is not supported in
' either string to be encrypted or the encryption key.
'
'get disk serial number
Public Const MAX_FILENAME_LEN = 256
Public Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)


'  Returns a volume's serial number
'
Public Function GetSerialNumber(sDrive As String) As Long
   Dim ser As Long
   Dim s As String * MAX_FILENAME_LEN
   Dim s2 As String * MAX_FILENAME_LEN
   Dim i As Long
   Dim j As Long
   
   Call GetVolumeInformation(sDrive + ":\" & Chr$(0), s, MAX_FILENAME_LEN, ser, i, j, s2, MAX_FILENAME_LEN)
   GetSerialNumber = ser
End Function
Public Function EncryptString(ByVal InString As String, ByVal EncryptKey As String) As String
 
 Dim TempKey, OutString As String
 Dim OldChar, NewChar, CryptChar As Long
 Dim i As Integer
 
 ' Initilize i and make sure the EncryptKey is long enough
 i = 0
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 
 ' Loop through the string to encrypt each character.
 Do
  i = i + 1
  OldChar = Asc(Mid(InString, i, 1))
  CryptChar = Asc(Mid(TempKey, i, 1))
  
  ' If it's an even character, add the ASCII value of the
  ' appropriate character in the Key, otherwise, subract it.
  ' Also, make sure the value is between 0 and 127.
  Select Case i Mod 2
   Case 0      'Even Character
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
   Case Else   'Odd Character
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
  End Select
  ' If the value is less than 35, add 40 to it (to make sure
  ' it's in the display range) and put it in an escape
  ' sequence (using ! [ASCII Value 33] as the escape char)
  If NewChar < 35 Then
   OutString = OutString + "!" + Chr(NewChar + 40)
  Else
   OutString = OutString + Chr(NewChar)
  End If
 Loop Until i = Len(InString)
 
 EncryptString = OutString

End Function


Public Function DecryptString(ByVal InString As String, ByVal EncryptKey As String) As String
 
 Dim TempKey, OutString As String
 Dim OldChar, NewChar, CryptChar As Long
 Dim i, c As Integer
 
 ' Initialize c and i (loop variables)
 c = 0       ' c is used for InString
 i = 0       ' i is used for EncryptKey
 ' Make sure the EncryptKey is long enough
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 
 Do
  ' In the decrypt function, two integers are need keeping
  ' track of location (becuase the escape sequence it two
  ' chars long, but only has one placeholder in the key)
  
  i = i + 1
  c = c + 1
  OldChar = Asc(Mid(InString, c, 1))
  ' If this is an escape sequence, get the next character and
  ' subtract 40 from it's value.
  If OldChar = 33 Then
   c = c + 1
   OldChar = Asc(Mid(InString, c, 1))
   OldChar = OldChar - 40
  End If
  CryptChar = Asc(Mid(TempKey, i, 1))
  
  ' If it's an even character, subract the appropriate key
  ' value... also, if it's out of range, bring it back in.
  Select Case i Mod 2
   Case 0      'Even Character
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
   Case Else   'Odd Character
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
  End Select
  OutString = OutString + Chr(NewChar)
 Loop Until c = Len(InString)
 
 DecryptString = OutString

End Function

Sub CheckExpired()

' On Error Resume Next

'string terenkripsi
Dim strEnkCurDateDay As String
Dim strEnkCurDateMonth As String
Dim strEnkInstalledDay As String
Dim strEnkInstalledMonth As String
Dim strEnkElapsedDay As String
Dim strEnkMaxUsedDay As String

'string asli
Dim strCurDateDay As String
Dim strCurDateMonth As String
Dim strInstalledDay As String
Dim strInstalledMonth As String
Dim strElapsedDay As String
Dim strMaxUsedDay As String

' CurDate = Current Date
Dim sCheck As String
Dim dCurDate As Date

Dim iCurDateDay As Integer
Dim iCurDateMonth As Integer
Dim iInstalledDay As Integer
Dim iInstalledMonth As Integer
Dim iElapsedDay As Integer
Dim iMaxUsedDay As Integer

'deklarasi enkrip nilai
Dim strEnkiCurDateDay As String
Dim strEnkiCurDateMonth As String
Dim strEnkiInstalledDay As String
Dim strEnkiInstalledMonth As String
Dim strEnkiElapsedDay As String
Dim strEnkiMaxUsedDay As String
Dim strEnkiElapsedMonth As String

Dim iElapsedMonth As Integer
' Dim iMonthElapsed As Integer

' enkripsi kunci setting registri Enk=enkripsi
strEnkCurDateDay = EncryptString("CurDateDay", "a")
strEnkCurDateMonth = EncryptString("CurDateMonth", "a")
strEnkElapsedDay = EncryptString("ElapsedDay", "a")
strEnkInstalledDay = EncryptString("InstalledDay", "a")
strEnkInstalledMonth = EncryptString("InstalledMonth", "a")
strEnkMaxUsedDay = EncryptString("MaxUsedDay", "a")
 
strCurDateDay = DecryptString(strEnkCurDateDay, "a")
strCurDateMonth = DecryptString(strEnkCurDateMonth, "a")
strElapsedDay = DecryptString(strEnkElapsedDay, "a")
strInstalledDay = DecryptString(strEnkInstalledDay, "a")
strInstalledMonth = DecryptString(strEnkInstalledMonth, "a")
strMaxUsedDay = DecryptString(strEnkMaxUsedDay, "a")
 
bFirstInstall = False

iMaxUsedDay = 15
dCurDate = Date
iCurDateDay = Day(dCurDate)
iCurDateMonth = Month(dCurDate)

'enkripsi nilai
strEnkiMaxUsedDay = EncryptString(CStr(iMaxUsedDay), "1")
strEnkiCurDateDay = EncryptString(CStr(iCurDateDay), "1")
strEnkiCurDateMonth = EncryptString(CStr(iCurDateMonth), "1")

SaveSetting App.Title, "Setting", strEnkCurDateDay, strEnkiCurDateDay
SaveSetting App.Title, "Setting", strEnkCurDateMonth, strEnkiCurDateMonth

' bFirstInstall = True
If bFirstInstall = True Then
  iInstalledDay = iCurDateDay
  iInstalledMonth = iCurDateMonth
  strEnkiInstalledDay = EncryptString(CStr(iInstalledDay), "1")
  strEnkiInstalledMonth = EncryptString(CStr(iInstalledMonth), "1")
  SaveSetting App.Title, "Setting", strEnkInstalledDay, strEnkiInstalledDay
  SaveSetting App.Title, "Setting", strEnkInstalledMonth, strEnkiInstalledMonth
Else
  strEnkiInstalledDay = GetSetting(App.Title, "Setting", strEnkInstalledDay, strEnkiInstalledDay)
  strEnkiInstalledMonth = GetSetting(App.Title, "Setting", strEnkInstalledMonth, strEnkiInstalledMonth)
  iInstalledDay = Val(DecryptString(strEnkiInstalledDay, "1"))
  iInstalledMonth = Val(DecryptString(strEnkiInstalledMonth, "1"))
End If

 sCheck = strInstalledDay & iInstalledDay & vbCr & _
 strInstalledMonth & iInstalledMonth & vbCr

 'MsgBox sCheck

'maximum day in use 30 day dgn enkripsi
SaveSetting App.Title, "Setting", strEnkMaxUsedDay, strEnkiMaxUsedDay

If iCurDateDay - iInstalledDay < 0 Then
  iCurDateDay = iCurDateDay + 30
  iCurDateMonth = iCurDateMonth - 1
End If

' sCheck = sCheck & strCurDateDay & iCurDateDay & vbCr & _
 strCurDateMonth & iCurDateMonth & vbCr

'MsgBox sCheck

  iElapsedMonth = iCurDateMonth - iInstalledMonth
  iElapsedDay = (iCurDateDay - iInstalledDay) + iElapsedMonth * 30

 sCheck = sCheck & strElapsedDay & iElapsedDay & vbCr & _
   strElapsedDay & iElapsedDay & vbCr & _
   strMaxUsedDay & iMaxUsedDay
   
'MsgBox sCheck

 SaveSetting App.Title, "Setting", strEnkElapsedDay, strEnkiElapsedDay

 ' MsgBox sCheck, vbInformation, "BasicCOM2002_SetExpired"

bExpired = False
If iElapsedDay >= iMaxUsedDay Or iElapsedDay < 0 Then
  bExpired = True
End If

If bExpired = True Then
  MsgBox "Sudah expired..."
  End
 Else
  MsgBox "Masih oke..." & iMaxUsedDay - iElapsedDay & " hari ", vbInformation
  
End If


End Sub

Public Sub CekSerial()

  On Error GoTo localErr
    
  Dim s As String
  Dim strTemp As String
  Dim strMsg
  Dim sHDSerialFile As String
    
  Dim strTeknisi As String
  Dim strPemilik As String
  Dim strHDSerNum As String
  Dim strSerNum As String
    
  Dim strEncTeknisi As String
  Dim strEncPemilik As String
  Dim strEncHDSerNum As String
  Dim strEncSerNum As String
  
  strEncTeknisi = GetSetting("XBWGen", "XBW", "Support", strTemp)
  strEncPemilik = GetSetting("XBWGen", "XBW", "Owner", strTemp)
  strEncHDSerNum = GetSetting("XBWGen", "XBW", "HD", strTemp)
  strEncSerNum = GetSetting("XBWGen", "XBW", "SN", strTemp)
  
  strTeknisi = DecryptString(strEncTeknisi, "x")
  strPemilik = DecryptString(strEncPemilik, "x")
  strHDSerNum = DecryptString(strEncHDSerNum, "x")
  strSerNum = DecryptString(strEncSerNum, "x")
  
  strMsg = "Srial No = " & strSerNum & vbCr & _
          "Teknisi = " & strTeknisi & vbCr & "Pemilik = " & strPemilik
  
  'strHDSerNum = "-796716010" And _

  'MsgBox strTeknisi & " " & strPemilik & " " & strSerNum
  
  Open App.Path & "\hds.xbw" For Input As #1
  Do While Not EOF(1)
   Input #1, sHDSerialFile
  Loop
  Close #1
  
  If strPemilik = "XBASICPRO" And _
     strHDSerNum = sHDSerialFile Then
         
     'strHDSerNum = "1278061001" Then
     'strSerNum = "X28B04P72" Then
       'MsgBox "Legal software..  Selamat", vbInformation
    Else
     MsgBox "Sorry  Illegal software!!!" & vbCr & vbCr & _
        "Hubungi XBasicPro (061)77374922 - 081376772820", vbCritical
     End
  End If
     
 Exit Sub
localErr:
  s = "File 'hds.xbw' tidak ada! buat dengan XBWKeyGen.exe"
  MsgBox s, vbCritical
  End
End Sub




