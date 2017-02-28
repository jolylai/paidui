VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form duanx 
   Caption         =   "发短信"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4545
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Text            =   "Text7"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   6480
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "启用发送"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Text            =   "8"
      Top             =   2400
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   4920
   End
   Begin VB.TextBox Text3 
      Height          =   1695
      Left            =   1800
      TabIndex        =   5
      Text            =   "你好ssss "
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "15392029842"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "13010380500"
      Top             =   240
      Width           =   2415
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3840
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "计秒："
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "端口号："
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "内容："
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "接收号码："
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "短信中心号码："
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "duanx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Buffer As String, sslkja

Private Sub Command1_Click()
Text6.text = "0891" & telc(Text1.text) & "11000D91" & telc(Text2.text) & "000800" & ascg(Text3.text)
Text7.text = (Len(Text6.text) - InStr(Text6.text, "11000D91") + 1) / 2

End Sub



Private Sub Form_Load()
Dim strcmd  As String
Dim STR, wjm1, wjm2
Buffer = "1"

strcmd = Trim(Command()) '这就是那个参数了
'strcmd = "35132132135"
'MsgBox strcmd
If InStr(strcmd, "VbCrVbLf") > 0 Then
  STR = Split(strcmd, "VbCrVbLf")
  Text1.text = Trim(STR(0))
  Text2.text = Trim(STR(1))
  Text3.text = Trim(STR(2))
  Text4.text = Trim(STR(3))
  If Val(Text1.text) = 0 Or Val(Text2.text) = 0 Or Val(Text4.text) = 0 Or Text3.text = "" Then
    End
  End If
  
Else
  End
End If

End Sub

Private Sub MSComm1_OnComm()
Dim strsss As String
  Select Case MSComm1.CommEvent
        Case 2
        strsss = MSComm1.Input
        Text5.text = strsss
        MSComm1.InBufferCount = 0   '清空缓冲区
     End Select

End Sub

Private Sub Timer1_Timer()
On Error GoTo CuoWu    '增加这行
Dim ml(0) As Byte
ml(0) = &H1A  '命令结尾
Command1_Click
Label3.Caption = Val(Label3.Caption) + 1
If Val(Label3.Caption) > 30 Then
  End
End If

            With MSComm1
                If .PortOpen = True Then
                .PortOpen = False
                End If
                .CommPort = Val(Text4.text)
                .settings = "9600,n,8,1"
                .InBufferSize = 1024
                .OutBufferSize = 1024
                
                .InputMode = comInputModeText    '设置接收数据模式为文本形式
                '-----------------------------------------------------------------------------------------------------
                .InputLen = 0                     '设置Input 一次从接收缓冲读取全部字节数
                .SThreshold = 0                   '设置发送完所有产生OnComm事件
                .InBufferCount = 0                '清除接收缓冲区
                .OutBufferCount = 0               '清除发送缓冲区
                .RThreshold = 1                   '设置接收一个字节产生OnComm事件     '
                .RTSEnable = True
                    If Not .PortOpen Then             '判断通信口是否打开
                    On Error Resume Next
                    .PortOpen = True                '打开通信口
                    End If
           End With
           
  If Val(Label3.Caption) = 20 And sslkja <> 888 Then
     Buffer = "1"
     Text5.text = ""
     sslkja = 888
     Exit Sub
  End If

  If Buffer = "1" Then
    Buffer = "10"
    Text5.text = ""
    MSComm1.Output = "AT" + Chr(13)
    Exit Sub
  End If
  If Buffer = "10" And InStr(Text5.text, "OK") > 0 Then '检测到 AT回指令
     Buffer = "2"
     Text5.text = ""
     MSComm1.Output = "AT+CMGF=0" + Chr(13)
     Label3.Caption = 0
     Exit Sub
  End If
  If Buffer = "2" And InStr(Text5.text, "OK") > 0 Then '检测到 AT+CMGF=0回指令
     Buffer = "3"
     Text5.text = ""
     MSComm1.Output = "AT+CMGS=" & Text7.text & Chr(13)
      Label3.Caption = 0
     Exit Sub
  End If
  If Buffer = "3" And InStr(Text5.text, ">") > 0 Then '检测到 AT+CMGS=回指令
      Buffer = "4"
      Text5.text = ""
      MSComm1.Output = Text6.text '0891683110300805F011000D91685193029248F2000800044F60597D
       Label3.Caption = 0
     Exit Sub
  End If
  If Buffer = "4" Then  '发送 0x1a
      Buffer = "5"
      Text5.text = ""
      MSComm1.Output = ml
       Label3.Caption = 0
     Exit Sub
  End If
  If Buffer = "5" And InStr(Text5.text, "CMGS:") > 0 Then  '成功指令
     End
     Exit Sub
  End If
  
  'CMGS:
 ' MSComm1.Output = ml
  
  
  Exit Sub
CuoWu:

 Buffer = "1"
End Sub
