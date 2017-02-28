Attribute VB_Name = "led"
Public Function led_coad(style As String, adress As String, preamble As String, id As String, postamble As String) As String
    Dim strSj As String, strHexSj As String, n As String, bytSj() As Byte
    strHexSj = "3A2A48463630"
    strHexSj = strHexSj & (30 + adress)  '地址

    Select Case style
        Case 0
            strHexSj = strHexSj & "3031373739312020"
        Case 1
            strHexSj = strHexSj & "3031533739312020"
    End Select
    
    bytSj = StrConv(preamble, vbFromUnicode)  '前文
    For i = 0 To UBound(bytSj)
        strHexSj = strHexSj & Right("0" & Hex(bytSj(i)), 2)
    Next
    
'    strHexSj = strHexSj & "20202020"
'
    If Len(id) = 3 Then id = Right("00" & id, 4)
'
   bytSj = StrConv(id, vbFromUnicode)
    For i = 0 To UBound(bytSj)
        strHexSj = strHexSj & Right("0" & Hex(bytSj(i)), 2)
    Next

    bytSj = StrConv(postamble, vbFromUnicode)     '后文
    For i = 0 To UBound(bytSj)
        strHexSj = strHexSj & Right("0" & Hex(bytSj(i)), 2)
    Next

    strHexSj = strHexSj & "0D0A"
    
'    For i = 1 To Len(strHexSj) Step 2
'            strSj = Mid(strHexSj, i, 2) & " "
'            n = n & strSj
'    Next i
   led_coad = strHexSj
 'led_coad = "3A 2A 48 46 36 30 31 30 31 39 36 30 31 B7 A2 CB CD B5 BD CA C7 B5 C4 B7 A2 CB CD B4 F3 B8 F6 B6 EE B0 AC C9 FD B4 F3 C9 B5 C9 B5 B5 C4 CA C7 B5 C4 B0 A2 CB B9 B5 D9 B7 D2 B7 A2 CB CD B5 BD B7 B6 CE C4 B7 BC B0 B2 CE BF B7 A8 B0 A2 C8 F8 B5 C2 B6 F7 B0 AE 0D 0A"
End Function

