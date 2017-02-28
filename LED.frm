VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LEDpz 
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   7635
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8280
      Top             =   7800
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除"
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "发送"
      Height          =   495
      Left            =   5760
      TabIndex        =   12
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   480
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   6960
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   5760
      Width           =   4935
   End
   Begin VB.ComboBox Combo5 
      Height          =   300
      Left            =   3120
      TabIndex        =   9
      Text            =   "Combo5"
      Top             =   4440
      Width           =   2055
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      Left            =   3120
      TabIndex        =   8
      Text            =   "Combo4"
      Top             =   3720
      Width           =   2055
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   3120
      TabIndex        =   7
      Text            =   "Combo3"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   3120
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   3120
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "删除节目"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "停留"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "速度"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "特技"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "节目"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "LEDpz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim cIndex As Integer
    Select Case Combo2.text
        Case Is = "翻页"
            cIndex = 0
        Case Is = "上覆盖"
            cIndex = 1
        Case Is = "下覆盖"
            cIndex = 2
        Case Is = "左覆盖"
            cIndex = 3
        Case Is = "右覆盖"
            cIndex = 4
        Case Is = "向上移"
            cIndex = 5
        Case Is = "向下移"
            cIndex = 6
        Case Is = "向左移"
            cIndex = 7
        Case Is = "向右移"
            cIndex = 8
        Case Is = "垂直百叶窗"
            cIndex = 9
        Case Is = "闪烁"
            cIndex = 10
        Case Is = "单页静止"
            cIndex = 11
    End Select
    Open App.Path & "\led.ini" For Output As #1
    Write #1, Trim(Combo1.text - 1) & "VbCrVbLf" & Trim(cIndex) & "VbCrVbLf" & Trim(Combo3.text - 1) & "VbCrVbLf" & Trim(Combo4.text - 1) & "VbCrVbLf" & Trim(Combo5.text - 1)
    Close #1
    jzzh
    Text1.text = Val(Combo2.ListIndex)
End Sub

Private Sub Command2_Click()
    scjm
End Sub

Private Sub Form_Load()
'Adodc2.ConnectionString = sqlcnn
Adodc1.ConnectionString = sqlcnn
  Dim i As Integer, s As String, led() As String
    For i = 1 To 9                           '初始化combo1-5
        Combo1.AddItem i
        Combo4.AddItem i
        Combo5.AddItem i
    Next i
        Combo2.AddItem "翻页"
        Combo2.AddItem "上覆盖"
        Combo2.AddItem "下覆盖"
        Combo2.AddItem "左覆盖"
        Combo2.AddItem "右覆盖"
        Combo2.AddItem "向上移"
        Combo2.AddItem "向下移"
        Combo2.AddItem "向左移"
        Combo2.AddItem "向右移"
        Combo2.AddItem "垂直百叶窗"
        Combo2.AddItem "闪烁"
        Combo2.AddItem "单页静止"
    For i = 1 To 7
        Combo3.AddItem i
    Next i
    
    
    If Dir(App.Path & "\led.ini") <> "" Then

    Else
        Open App.Path & "\led.ini" For Output As #1
        Write #1, "0VbCrVbLf0VbCrVbLf0VbCrVbLf0VbCrVbLf0"
        Close #1
    End If
        Open App.Path & "\led.ini" For Input As #1
        Input #1, s
        Close #1
        led = Split(s, "VbCrVbLf")
        Combo1.text = Combo1.List(led(0))
        Combo2.text = Combo2.List(led(1))
        Combo3.text = Combo3.List(led(2))
        Combo4.text = Combo4.List(led(3))
        Combo5.text = Combo4.List(led(4))
        
    jzzh
End Sub
Private Sub scjm()   '删除节目
Dim s As String, x As String
    x = "3A2A48463630" & "" & "30305130" & Trim(Combo5.text + 30) & "30" & Trim(Combo5.text + 30) & "ODOA"
    Text2.text = ""
    For i = 1 To Len(x) Step 2
        s = Mid(x, i, 2) & " "
        Text2.text = Text2.text & s
    Next i
    Text2.text = Trim(Text2.text)
End Sub
Private Sub jzzh()    '转换十六进制
    Dim s As String, i As Integer, x As String, pinghao As Integer
    
    
        x = "3A2A484636"
'        pinghao = Adodc2.Recordset.Fields("屏号") '修改 屏幕地址
'        pinghao = pinghao + 3030
'        x = x & "trim(pinghao）"

        x = x & Trim(Combo1.text + 3030)  '节目号
        Select Case Combo2.text
            Case Is = "翻页"
                x = x & "30"
            Case Is = "上覆盖"
                x = x & "31"
            Case Is = "下覆盖"
                x = x & "32"
            Case Is = "左覆盖"
                x = x & "33"
            Case Is = "右覆盖"
                x = x & "34"
            Case Is = "向上移"
                x = x & "35"
            Case Is = "向下移"
                x = x & "36"
            Case Is = "向左移"
                x = x & "37"
            Case Is = "向右移"
                x = x & "38"
            Case Is = "垂直百叶窗"
                x = x & "39"
            Case Is = "闪烁"
                x = x & "41"
            Case Is = "单页静止"
                x = x & "53"
            End Select
        x = x & Trim(Combo3.text + 3030)  '速度
        x = x & Trim(Combo4.text + 3030)   '停留
        For i = 1 To Len(Text1.text)
            s = Hex(Asc(Mid(Text1.text, i, 1)))  '文字
            x = x & s
        Next i
        x = x & "310D0A"
        Text2.text = ""
        For i = 1 To Len(x) Step 2
            s = Mid(x, i, 2) & " "
            Text2.text = Text2.text & s
        Next i
        Text2.text = Trim(Text2.text)
End Sub

