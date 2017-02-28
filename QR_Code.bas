Attribute VB_Name = "QR_Code"
Public Function qrma(spjz As String, strs As String) As String
        Dim lsstrss As String, lsstrss1 As String, lsstrss2 As String, lsstrss3 As Long, mkls As String
        lsstrss = "1b 40"
        
        lsstrss3 = 3
        For i = 1 To Len(strs)
            If Asc(Mid(strs, i, 1)) > 0 Then  '为英文
                lsstrss3 = lsstrss3 + 1
            Else
                 lsstrss3 = lsstrss3 + 2
            End If
        Next
        Select Case lsstrss3  '根据字符长度选择打印大小
            Case 0 To 15
                mkls = 11
            Case 16 To 20
                mkls = 10
            Case 21 To 30
                mkls = 9
            Case 31 To 40
                mkls = 8
            Case 41 To 60
                mkls = 7
            Case 61 To 90
                mkls = 6
            Case 91 To 140
                mkls = 5
            Case Else
                mkls = 4
        End Select

        
        If Len(Hex(Val(mkls))) = 1 Then lsstrss = lsstrss & " 1d 28 6b 03 00 31 43 0" & Hex(Val(mkls))   'QR 码的模块类型
        If Len(Hex(Val(mkls))) = 2 Then lsstrss = lsstrss & " 1d 28 6b 03 00 31 43 " & Hex(Val(mkls))   'QR 码的模块类型
        
        If spjz = "7%" Then lsstrss2 = "30"
        If spjz = "15%" Then lsstrss2 = "31"
        If spjz = "25" Then lsstrss2 = "32"
        If spjz = "30%" Then lsstrss2 = "33"
        lsstrss = lsstrss & " 1d 28 6b 03 00 31 45 " & lsstrss2   '校正水平误差
        
        lsstrss3 = 3
        For i = 1 To Len(strs)
            If Asc(Mid(strs, i, 1)) > 0 Then  '为英文
            
                    If Len(Hex(Asc(Mid(strs, i, 1)))) = 1 Then lsstrss1 = lsstrss1 & " 0" & Hex(Asc(Mid(strs, i, 1)))
                    If Len(Hex(Asc(Mid(strs, i, 1)))) = 2 Then lsstrss1 = lsstrss1 & " " & Hex(Asc(Mid(strs, i, 1)))
                lsstrss3 = lsstrss3 + 1
            Else
                 lsstrss3 = lsstrss3 + 2
              lsstrss1 = lsstrss1 & " " & Mid(Hex(AscW(StrConv(Mid(strs, i, 1), vbFromUnicode))), 3, 2) & " " & Mid(Hex(AscW(StrConv(Mid(strs, i, 1), vbFromUnicode))), 1, 2)
            End If
           
        Next
        
        
        
        If Len(Hex(lsstrss3)) = 1 Then
           lsstrss = lsstrss & " 1d 28 6b 0" & Hex(lsstrss3) & " 00 31 50 30" & lsstrss1
        Else
           lsstrss = lsstrss & " 1d 28 6b " & Hex(lsstrss3) & " 00 31 50 30" & lsstrss1
        End If
          lsstrss = lsstrss & " 1b 61 01 1d 28 6b 03 00 31 52 30 1d 28 6b 03 00 31 51 30" '& " 1b 4a 15 1b 4a 15 1b 4a 15 1b 4a 15 1b 4a 15 1b 4a 15 1b 69"
        qrma = lsstrss
End Function

