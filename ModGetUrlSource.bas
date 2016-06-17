<<<<<<< HEAD
Attribute VB_Name = "ModGetUrlSource"
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Const FLAG_ICC_FORCE_CONNECTION As Long = &H1&
Public GetConnectionErr As Boolean

Public Function CheckConnection() As Boolean
    If InternetCheckConnection(TestHost, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then

        'Print "网络未连接"
        CheckConnection = False
    Else

        'Print "网络已连接"
        CheckConnection = True
    End If

End Function


Public Function GetUrlResource(ByVal URL As String)
On Local Error GoTo Err
Dim xmlHTTP1 As Object
Set xmlHTTP1 = CreateObject("Microsoft.XMLHTTP")
xmlHTTP1.Open "POST", URL, True '这里可以用POST和GET，不过POST更好些，得到的网页始终最新
xmlHTTP1.send
While xmlHTTP1.ReadyState <> 4
DoEvents
Wend
GetUrlResource = BytesToBstr(xmlHTTP1.responseBody)
Set ObjXML = Nothing
GetConnectionErr = False
Exit Function

Err:
GetConnectionErr = True
Err.Clear
End Function


Private Function BytesToBstr(Bytes)
Dim Unicode As String

If IsUTF8(Bytes) Then
    Unicode = "UTF-8"
Else
    Unicode = "GB2312"
End If

Dim objstream As Object
Set objstream = CreateObject("ADODB.Stream")
With objstream
    .Type = 1
    .Mode = 3
    .Open
    .Write Bytes
    .Position = 0
    .Type = 2
    .Charset = Unicode
    BytesToBstr = .ReadText
    .Close
End With
End Function


Private Function IsUTF8(Bytes) As Boolean
Dim I As Long, AscN As Long, Length As Long
Length = UBound(Bytes) + 1

If Length < 3 Then
    IsUTF8 = False
    Exit Function
ElseIf Bytes(0) = &HEF And Bytes(1) = &HBB And Bytes(2) = &HBF Then
    IsUTF8 = True
    Exit Function
End If


Do While I <= Length - 1
    If Bytes(I) < 128 Then
        I = I + 1
        AscN = AscN + 1
    ElseIf (Bytes(I) And &HE0) = &HC0 And (Bytes(I + 1) And &HC0) = &H80 Then
        I = I + 2
    ElseIf I + 2 < Length Then
        If (Bytes(I) And &HF0) = &HE0 And (Bytes(I + 1) And &HC0) = &H80 And (Bytes(I + 2) And &HC0) = &H80 Then
            I = I + 3
        Else
            IsUTF8 = False
            Exit Function
        End If
    Else
        IsUTF8 = False
        Exit Function
    End If
Loop

If AscN = Length Then
    IsUTF8 = False
Else
    IsUTF8 = True
End If

End Function



Public Function DownloadFile(ByVal URL As String, ByVal SavePath As String)

    Dim xmlHTTP, Sobj
    Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")
    xmlHTTP.Open "POST", URL, False
    DoEvents
    xmlHTTP.send
    Set Sobj = CreateObject("ADODB.Stream")
    Sobj.Type = 1
    Sobj.Open
    Sobj.Write xmlHTTP.responseBody
    Sobj.SaveToFile SavePath, 2
    Sobj.Close

End Function


=======
Attribute VB_Name = "ModGetUrlSource"
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Const FLAG_ICC_FORCE_CONNECTION As Long = &H1&
Public GetConnectionErr As Boolean

Public Function CheckConnection() As Boolean
    If InternetCheckConnection(TestHost, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then

        'Print "网络未连接"
        CheckConnection = False
    Else

        'Print "网络已连接"
        CheckConnection = True
    End If

End Function


Public Function GetUrlResource(ByVal URL As String)
On Local Error GoTo Err
Dim xmlHTTP1 As Object
Set xmlHTTP1 = CreateObject("Microsoft.XMLHTTP")
xmlHTTP1.Open "POST", URL, True '这里可以用POST和GET，不过POST更好些，得到的网页始终最新
xmlHTTP1.send
While xmlHTTP1.ReadyState <> 4
DoEvents
Wend
GetUrlResource = BytesToBstr(xmlHTTP1.responseBody)
Set ObjXML = Nothing
GetConnectionErr = False
Exit Function

Err:
GetConnectionErr = True
Err.Clear
End Function


Private Function BytesToBstr(Bytes)
Dim Unicode As String

If IsUTF8(Bytes) Then
    Unicode = "UTF-8"
Else
    Unicode = "GB2312"
End If

Dim objstream As Object
Set objstream = CreateObject("ADODB.Stream")
With objstream
    .Type = 1
    .Mode = 3
    .Open
    .Write Bytes
    .Position = 0
    .Type = 2
    .Charset = Unicode
    BytesToBstr = .ReadText
    .Close
End With
End Function


Private Function IsUTF8(Bytes) As Boolean
Dim I As Long, AscN As Long, Length As Long
Length = UBound(Bytes) + 1

If Length < 3 Then
    IsUTF8 = False
    Exit Function
ElseIf Bytes(0) = &HEF And Bytes(1) = &HBB And Bytes(2) = &HBF Then
    IsUTF8 = True
    Exit Function
End If


Do While I <= Length - 1
    If Bytes(I) < 128 Then
        I = I + 1
        AscN = AscN + 1
    ElseIf (Bytes(I) And &HE0) = &HC0 And (Bytes(I + 1) And &HC0) = &H80 Then
        I = I + 2
    ElseIf I + 2 < Length Then
        If (Bytes(I) And &HF0) = &HE0 And (Bytes(I + 1) And &HC0) = &H80 And (Bytes(I + 2) And &HC0) = &H80 Then
            I = I + 3
        Else
            IsUTF8 = False
            Exit Function
        End If
    Else
        IsUTF8 = False
        Exit Function
    End If
Loop

If AscN = Length Then
    IsUTF8 = False
Else
    IsUTF8 = True
End If

End Function



Public Function DownloadFile(ByVal URL As String, ByVal SavePath As String)

    Dim xmlHTTP, Sobj
    Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")
    xmlHTTP.Open "POST", URL, False
    DoEvents
    xmlHTTP.send
    Set Sobj = CreateObject("ADODB.Stream")
    Sobj.Type = 1
    Sobj.Open
    Sobj.Write xmlHTTP.responseBody
    Sobj.SaveToFile SavePath, 2
    Sobj.Close

End Function


>>>>>>> 2abdb2e3a24a7144f6e9a89c8664e5d0469ec446
