Attribute VB_Name = "jiajiemi"


Function UserCode(password As String) As String
'用户口令加密
    Dim il_bit, il_x, il_y, il_z, il_len, I As Long
    Dim is_out As String
     If Trim(password) = "" Then
       UserCode = ""
       Exit Function
    End If
    il_len = Len(password)
    il_x = 0
    il_y = 0
    is_out = ""
    For I = 1 To il_len
        il_bit = AscW(Mid(password, I, 1))    'W系列支持unicode
        
        il_y = (il_bit * 13 Mod 256) + il_x
        is_out = is_out & ChrW(Fix(il_y))  '取整 int和fix区别: fix修正负数
        il_x = il_bit * 13 / 256
    Next
    is_out = is_out & ChrW(Fix(il_x))
    
    password = is_out
    il_len = Len(password)
    il_x = 0
    il_y = 0
    is_out = ""
    For I = 1 To il_len
        il_bit = AscW(Mid(password, I, 1))
        '取前4位值
        il_y = il_bit / 16 + 64
        is_out = is_out & ChrW(Fix(il_y))
        '取后4位值
        il_y = (il_bit Mod 16) + 64
        is_out = is_out & ChrW(Fix(il_y))
    Next
    UserCode = is_out
End Function
Function UserDeCode(password As String) As String
'口令解密
    Dim is_out As String
    Dim il_x, il_y, il_len, I, il_bit As Long
    If Trim(password) = "" Then
       UserDeCode = ""
       Exit Function
    End If
    
    
    
    
    
    il_len = Len(password)
    il_x = 0
    il_y = 0
    is_out = ""
    For I = 1 To il_len Step 2
        il_bit = AscW(Mid(password, I, 1))
        '取前4位值
        il_y = (il_bit - 64) * 16
        '取后4位值
        'dd = AscW(Mid(password, i + 1, 1)) - 64
        il_y = il_y + AscW(Mid(password, I + 1, 1)) - 64
        is_out = is_out & ChrW(il_y)
    Next
    il_x = 0
    il_y = 0
    password = is_out
    is_out = ""
    il_len = Len(password)
    il_x = AscW(Mid(password, il_len, 1))
    For I = (il_len - 1) To 1 Step -1
        il_y = il_x * 256 + AscW(Mid(password, I, 1))
        il_x = il_y Mod 13
        is_out = ChrW(Fix(il_y / 13)) & is_out
    Next
    UserDeCode = is_out
End Function

