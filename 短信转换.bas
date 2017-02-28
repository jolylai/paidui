Attribute VB_Name = "Module2"

Public Function chg(rmsg As String) As String
'中文转换为Unicode码
    Dim tep As String
    Dim temp As String
    Dim i As Integer
    Dim b As Integer
    tep = rmsg
    i = Len(tep)
    b = i / 4
    If i = b * 4 Then
      b = b - 1
       tep = Left(tep, b * 4)
    Else
       tep = Left(tep, b * 4)
    End If
       chg = ""
    For i = 1 To b
  
       temp = "&H" & Mid(tep, (i - 1) * 4 + 1, 4)
       chg = chg & ChrW(CInt(Val(temp)))
    Next i
  End Function

Public Function telc(num As String) As String
    Dim tl As Integer
    Dim ltem, rtem, ttem As String
    Dim ti As Integer
    ttem = ""
    tl = Len(num)
    If tl <> 11 And tl <> 13 Then
      MsgBox " 11"
      Exit Function
    End If
    If tl = 11 Then
         tl = tl + 2
       num = "86" & num
    End If
    num = num & "F"
    
    For ti = 1 To tl + 1 Step 2
      ltem = Mid(num, ti, 1)
      rtem = Mid(num, ti + 1, 1)
      ttem = Trim(ttem) & Trim(rtem) & Trim(ltem)
      
      
      
        
     ' If ti = tl Then rtem = "F"
      
    Next ti
   ' ttem = ttem & rtem
    telc = ttem
  End Function

Public Function ascg(smsg As String) As String

    Dim si, sb As Integer, cdu As String
    Dim stmp As Integer
    Dim stemp As String
    sb = Len(smsg)
    ascg = ""
    For si = 1 To sb
    stmp = AscW(Mid(smsg, si, 1))
    If Abs(stmp) < 127 Then
    stemp = "00" & Hex(stmp)
    Else
    stemp = Hex(stmp)
    End If
    ascg = ascg & stemp
    Next si
    
    
    
    ascg = Trim(ascg)
    If Len(Hex(Len(ascg) / 2)) = 1 Then
      cdu = "0"
    End If
      cdu = cdu & Hex(Len(ascg) / 2)
    
    ascg = cdu & Trim(ascg)
    End Function

