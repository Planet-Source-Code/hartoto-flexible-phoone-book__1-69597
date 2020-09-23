Attribute VB_Name = "modXML"
Public oDom As MSXML.DOMDocument
Public oNode As IXMLDOMNode
Public Const XML_LOAD_COMPLETE = 4

Dim oXML As New DOMDocument
Dim oNodeTes As IXMLDOMNode
Dim oNodeCount As IXMLDOMNode
Dim oEl As IXMLDOMElement
'Dim oNode As IXMLDOMNode

Public Function getXMLDB(sSPName As String, sParName As String, sParValue As String) As MSXML.DOMDocument

 Dim oXML As MSXML.DOMDocument
 Set oXML = New MSXML.DOMDocument
 
 Dim ocmd As New ADODB.Command
 Dim oPar1 As ADODB.Parameter
 
 Dim oParOut As ADODB.Parameter
 
 ocmd.ActiveConnection = conn
 'MsgBox oNode.xml
 ocmd.CommandText = sSPName
 'MsgBox oNode.Attributes(1).Text, , "EditFormDetails.getRsDetails"
 ocmd.CommandType = adCmdStoredProc
 Set oPar1 = ocmd.CreateParameter(sParName, adVarChar, adParamInput, 1000, sParValue)
 ocmd.Parameters.Append oPar1

 ocmd.Properties("Output stream") = oXML
 ocmd.Execute , , &H400  '&H400 = adExecuteStream
 
 Set getXMLDB = oXML
 
 'MsgBox "oXML.xml=" & oXML.xml, , "GetXMLDB"
  
 Set ocmd = Nothing
 Set oXML = Nothing

End Function

Public Sub AddNode(ByVal oXML As DOMDocument)
  
  Set oNode = oXML.childNodes(0)
  Dim oNodeCount As IXMLDOMNode
  Dim oEl As IXMLDOMElement
  Set oEl = oXML.createElement("D")
  Set oNodeCount = oXML.childNodes(0).childNodes(0)
 'MsgBox oNodeCount.Attributes.length
  AttrLength = oNodeCount.Attributes.length - 1
  'MsgBox oNode.childNodes(0).xml
  For i = 0 To AttrLength
    'MsgBox oNodeCount.Attributes(i).baseName
    oEl.setAttribute oNodeCount.Attributes(i).baseName, ""
  Next i
  oNode.appendChild oEl
  
  'MsgBox oXML.xml, , "AddNode"
  
End Sub
Public Sub DelNode(ByVal oXML As DOMDocument, ByVal iChild As Integer)

  Dim oNodeDel As IXMLDOMNode
  'MsgBox oNode.xml
  'MsgBox oXML.childNodes(0).childNodes.length
  Set oNode = oXML.childNodes(0)
  Set oNodeDel = oXML.childNodes(0).childNodes(iChild - 1)
  'MsgBox oNodeDel.xml
  oNode.removeChild oNodeDel
  
End Sub

Function XMLDataBlankGet(ByVal poXML As MSXML.DOMDocument) As String
    
   Dim kt2 As String, kt As String, sp As String
   kt = Chr(34)
   kt2 = Chr(34) & Chr(34)
   sp = Space(1)

   Dim oNd_M As IXMLDOMNode
   Dim oNd_D As IXMLDOMNode
   
   Set oNd_M = oDom.selectSingleNode("//editform[@name=""" & "myform" & """]")
   
   '--- template for blank XMLData
   s = "<M " & Space(1)
   For i = 0 To oNd_M.childNodes.length - 1
     If oNd_M.childNodes(i).Attributes(2).Text = "Tanggal" Then
       s = s & oNd_M.childNodes(i).Attributes(2).Text & "=" & kt & Date & kt & sp
     Else
       s = s & oNd_M.childNodes(i).Attributes(2).Text & "=" & kt2 & sp
     End If
   Next i
   s = s & ">"
  
'   Set oNd_D = oDom.selectSingleNode("//details/editForm[@name=""" & "myformdetails" & """]")
'   'For j = 1 To 2 'buat 2 brs detail
'    s = s & "<D" & sp
'    For i = 0 To oNd_D.childNodes.length - 1
'      If oNd_D.childNodes(i).Attributes(7).Text = "dtdate" Then
'        s = s & oNd_D.childNodes(i).Attributes(2).Text & "=" & kt & Date & kt & sp
'      ElseIf oNd_D.childNodes(i).Attributes(7).Text = "number" Then
'        s = s & oNd_D.childNodes(i).Attributes(2).Text & "=" & kt & "0" & kt & sp
'      ElseIf oNd_D.childNodes(i).Attributes(7).Text = "string" Then
'        s = s & oNd_D.childNodes(i).Attributes(2).Text & "=" & kt2 & sp
'      Else
'       ' s = s & oNd_D.childNodes(i).Attributes(2).Text & "=" & kt2 & sp
'       MsgBox "Type data '" & oNd_D.childNodes(i).Attributes(7).Text & "' belum didefenisikan"
'      End If
'    Next i
'    s = s & "/>"
'   'Next j
   s = s & "</M>"
   
   XMLDataBlankGet = s
   
   Set oNd_M = Nothing
   Set oNd_D = Nothing

End Function

Function oXMLDataBlankGetFor(ByVal sXMLFile As String) As MSXML.DOMDocument
   
   '-- create double entry blank data for jurnal
   Dim kt2 As String, kt As String, sp As String
   Dim s As String
   kt = Chr(34)
   kt2 = Chr(34) & Chr(34)
   sp = Space(1)
   
   Dim sFileXml As String
   Dim oXML As New MSXML.DOMDocument
   Dim oXMLTemp As New MSXML.DOMDocument
   
   sFileXml = App.Path & "\" & sXMLFile
   'MsgBox "sFileGLXml = " & sFileGLXml
   oXML.Load sXMLFile
   Do
    DoEvents
   Loop Until oXML.readyState = XML_LOAD_COMPLETE
   'MsgBox "oXMLGL.xml = " & oXMLGL.xml
   
   If Len(Trim(oXML.xml)) = "" Then
     GoSub ErrXml
   End If
   
   Dim oNd_M As IXMLDOMNode
   Dim oNd_D As IXMLDOMNode
   Set oNd_M = oXML.selectSingleNode("//editForm[@name=""" & "myform" & """]")
   
   '--- template for blank XMLData
   s = "<M " & Space(1)
   For i = 0 To oNd_M.childNodes.length - 1
     If oNd_M.childNodes(i).Attributes(7).Text = "dtdate" Then
       s = s & oNd_M.childNodes(i).Attributes(2).Text & "=" & kt & Date & kt & sp
     ElseIf oNd_M.childNodes(i).Attributes(7).Text = "number" Then
       s = s & oNd_M.childNodes(i).Attributes(2).Text & "=" & kt & "0" & kt & sp
     ElseIf oNd_M.childNodes(i).Attributes(7).Text = "string" Then
       s = s & oNd_M.childNodes(i).Attributes(2).Text & "=" & kt2 & sp
     End If
   Next i
   s = s & ">"
  
   Set oNd_D = oXML.selectSingleNode("//details/editForm[@name=""" & "myformdetails" & """]")
   'For j = 1 To 2 'buat 2 brs detail
    s = s & "<D" & sp
    For i = 0 To oNd_D.childNodes.length - 1
      If oNd_D.childNodes(i).Attributes(7).Text = "dtdate" Then
        s = s & oNd_D.childNodes(i).Attributes(2).Text & "=" & kt & Date & kt & sp
      ElseIf oNd_D.childNodes(i).Attributes(7).Text = "number" Then
        s = s & oNd_D.childNodes(i).Attributes(2).Text & "=" & kt & "0" & kt & sp
      ElseIf oNd_D.childNodes(i).Attributes(7).Text = "string" Then
        s = s & oNd_D.childNodes(i).Attributes(2).Text & "=" & kt2 & sp
      End If
    Next i
    s = s & "/>"
   'Next j
   s = s & "</M>"
   
   oXMLTemp.loadXML s
   Set oXMLDataBlankGetFor = oXMLTemp
   
   Set oNd_M = Nothing
   Set oNd_D = Nothing
   Set oXML = Nothing
   Set oXMLTemp = Nothing
   
   Exit Function
   
ErrXml:
  MsgBox "Error file " & sXMLFile & " tidak valid atau tdk ada!", , "Pembelian.XMLDataBlankGetForGL"
Return

End Function

Function XMLDataBlankGetMD(ByVal poXML As MSXML.DOMDocument) As String
        
   Dim kt2 As String, kt As String, sp As String
   kt = Chr(34)
   kt2 = Chr(34) & Chr(34)
   sp = Space(1)

   Dim oNd_M As IXMLDOMNode
   Dim oNd_D As IXMLDOMNode
   
   Set oNd_M = oDom.selectSingleNode("//editform[@name=""" & "myform" & """]")
   
   '--- template for blank XMLData
   s = "<M " & Space(1)
   For i = 0 To oNd_M.childNodes.length - 1
     'MsgBox oNd_M.childNodes(i).Attributes(7).Text
     If oNd_M.childNodes(i).Attributes(7).Text = "dt_date" Then
       s = s & oNd_M.childNodes(i).Attributes(2).Text & "=" & kt & Date & kt & sp
     ElseIf oNd_M.childNodes(i).Attributes(7).Text = "dt_number" Then
       s = s & oNd_M.childNodes(i).Attributes(2).Text & "=" & kt & 0 & kt & sp
     ElseIf oNd_M.childNodes(i).Attributes(7).Text = "dt_string" Then
       s = s & oNd_M.childNodes(i).Attributes(2).Text & "=" & kt2 & sp
     ElseIf oNd_M.childNodes(i).Attributes(7).Text = "dt_bit" Then
       s = s & oNd_M.childNodes(i).Attributes(2).Text & "=" & kt & 0 & kt & sp
     Else
       MsgBox "tidak ditemukan tipe data '" & _
          oNd_M.childNodes(i).Attributes(7).Text & "'", vbCritical, "modXML.XMLDataBankGetMD"
     End If
   Next i
   s = s & ">"
  
   Set oNd_D = oDom.selectSingleNode("//details/editform[@name=""" & "myformdetails" & """]")
   'For j = 1 To 2 'buat 2 brs detail
    s = s & "<D" & sp
    For i = 0 To oNd_D.childNodes.length - 1
      If oNd_D.childNodes(i).Attributes(7).Text = "dt_date" Then
        s = s & oNd_D.childNodes(i).Attributes(2).Text & "=" & kt & Date & kt & sp
      ElseIf oNd_D.childNodes(i).Attributes(7).Text = "dt_number" Then
        s = s & oNd_D.childNodes(i).Attributes(2).Text & "=" & kt & "0" & kt & sp
      ElseIf oNd_D.childNodes(i).Attributes(7).Text = "dt_string" Then
        s = s & oNd_D.childNodes(i).Attributes(2).Text & "=" & kt2 & sp
      Else
       ' s = s & oNd_D.childNodes(i).Attributes(2).Text & "=" & kt2 & sp
       MsgBox "Type data '" & oNd_D.childNodes(i).Attributes(7).Text & "' belum didefenisikan"
      End If
    Next i
    s = s & "/>"
   'Next j
   s = s & "</M>"
   
   XMLDataBlankGetMD = s
   
   Set oNd_M = Nothing
   Set oNd_D = Nothing

End Function

Public Function getXMLDB1Par(sSPName As String, _
  sParName As String, sParValue As String) As MSXML.DOMDocument

 Dim oXML As MSXML.DOMDocument
 Set oXML = New MSXML.DOMDocument
 
 Dim ocmd As New ADODB.Command
 Dim oPar1 As ADODB.Parameter
 
 Dim oParOut As ADODB.Parameter
 
 ocmd.ActiveConnection = conn
 'MsgBox oNode.xml
 ocmd.CommandText = sSPName
 'MsgBox oNode.Attributes(1).Text, , "EditFormDetails.getRsDetails"
 ocmd.CommandType = adCmdStoredProc
 Set oPar1 = ocmd.CreateParameter(sParName, adVarChar, adParamInput, 1000, sParValue)
 ocmd.Parameters.Append oPar1

 ocmd.Properties("Output stream") = oXML
 ocmd.Execute , , &H400  '&H400 = adExecuteStream
 
 Set getXMLDB1Par = oXML
 
 'MsgBox "oXML.xml=" & oXML.xml, , "GetXMLDB"
  
 Set ocmd = Nothing
 Set oXML = Nothing

End Function


