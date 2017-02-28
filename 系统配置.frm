VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form xtpeiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "系统配置"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   9600
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   8160
      TabIndex        =   71
      Text            =   "Text25"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox Text24 
      Height          =   270
      Left            =   1320
      TabIndex        =   69
      Text            =   "Text24"
      Top             =   6600
      Width           =   6015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出系统"
      Height          =   975
      Left            =   8400
      TabIndex        =   66
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "保存"
      Height          =   975
      Left            =   8400
      TabIndex        =   65
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "屏幕键盘"
      Height          =   615
      Left            =   7200
      TabIndex        =   64
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "短信"
      Height          =   1215
      Left            =   3600
      TabIndex        =   57
      Top             =   2520
      Width           =   5895
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   60
         Top             =   240
         Width           =   5175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "测试短信"
         Height          =   375
         Left            =   4200
         TabIndex        =   59
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   2040
         TabIndex        =   58
         Text            =   "13800100500"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "内容"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "短信中心号码"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   61
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "页面布置"
      Height          =   1215
      Left            =   120
      TabIndex        =   34
      Top             =   5280
      Width           =   9375
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   1200
         TabIndex        =   63
         Text            =   "Combo4"
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox combo1 
         Height          =   300
         Left            =   5760
         TabIndex        =   46
         Text            =   "Combo1"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text14 
         Height          =   270
         Left            =   3240
         TabIndex        =   45
         Text            =   "4000"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text15 
         Height          =   270
         Left            =   4680
         TabIndex        =   44
         Text            =   "4000"
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox combo2 
         Height          =   300
         Left            =   7680
         TabIndex        =   43
         Text            =   "Combo2"
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "有广告"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text16 
         Height          =   270
         Left            =   1200
         TabIndex        =   41
         Text            =   "2000"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text17 
         Height          =   270
         Left            =   10080
         TabIndex        =   40
         Text            =   "122"
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text18 
         Height          =   270
         Left            =   3240
         TabIndex        =   39
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text19 
         Height          =   270
         Left            =   4680
         TabIndex        =   38
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text20 
         Height          =   270
         Left            =   5760
         TabIndex        =   37
         Text            =   "1000"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text21 
         Height          =   270
         Left            =   6840
         TabIndex        =   36
         Text            =   "1000"
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   7680
         TabIndex        =   35
         Text            =   "Combo3"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "行"
         Height          =   255
         Left            =   6960
         TabIndex        =   56
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "列"
         Height          =   255
         Left            =   8640
         TabIndex        =   55
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "上边距"
         Height          =   255
         Left            =   2520
         TabIndex        =   54
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "左边距"
         Height          =   255
         Left            =   3960
         TabIndex        =   53
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "图片间距"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Left            =   9360
         TabIndex        =   51
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "上边距"
         Height          =   255
         Left            =   2520
         TabIndex        =   50
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "左边距"
         Height          =   255
         Left            =   3960
         TabIndex        =   49
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "宽"
         Height          =   255
         Left            =   5400
         TabIndex        =   48
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "高"
         Height          =   255
         Left            =   6480
         TabIndex        =   47
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "端口设置"
      Height          =   1335
      Left            =   3600
      TabIndex        =   25
      Top             =   3840
      Width           =   4215
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   3240
         TabIndex        =   32
         Text            =   "Text11"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   1320
         TabIndex        =   31
         Text            =   "2"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text22 
         Height          =   270
         Left            =   1320
         TabIndex        =   27
         Text            =   "4"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text23 
         Height          =   270
         Left            =   3240
         TabIndex        =   26
         Text            =   "Text23"
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "短信端口号"
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   33
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "打印端口"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "led/叫号端口"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "叫号端口"
         Height          =   375
         Left            =   2280
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "清空未叫"
      Height          =   495
      Left            =   7200
      TabIndex        =   24
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "清除数据"
      Height          =   615
      Left            =   7200
      TabIndex        =   23
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "led效果"
      Height          =   375
      Left            =   7920
      TabIndex        =   22
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "打印机数据编辑"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   6960
      Width           =   9375
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   5280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Text            =   "系统配置.frx":0000
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   5280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "系统配置.frx":000B
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Text            =   "系统配置.frx":001C
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "系统配置.frx":002A
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "标尾"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "说明："
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "电话："
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "标题"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   11640
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   11880
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据连接设置"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "测试连接"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "服务器："
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "数据库："
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "用户名："
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "密  码："
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   8520
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "会员信息"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "系统配置.frx":003B
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   22
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "座位类别"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11280
      Top             =   11040
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin VB.Label Label17 
      Caption         =   "自动确定时间"
      Height          =   255
      Left            =   8040
      TabIndex        =   70
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label16 
      Caption         =   "二维码网址"
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "a123456789123456"
      Height          =   255
      Left            =   7320
      TabIndex        =   67
      Top             =   2160
      Width           =   2175
   End
End
Attribute VB_Name = "xtpeiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conn As New ADODB.Connection



Private Sub Command1_Click()
'Text10.text = ascg("工作愉快！")
'Text10.text = "0891" & telc(Text13.text) & "11000D91" & telc("13902433649") & "000800" & ascg("工作愉快！")
'MsgBox Len(Text10.text) - InStr(Text10.text, "11000D91") + 1
'13010380500
Dim STR
STR = Text13.text & "VbCrVbLf" & "15392029842" & "VbCrVbLf" & Text10.text & "VbCrVbLf" & Text11.text
Shell (App.Path & "\短信.exe " & STR)


End Sub

Private Sub Command2_Click()
'On Error GoTo CuoWu    '增加这行
Text5.text = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Text3.text & ";pwd=" & Text4.text & ";Data Source=" & Text1.text & ";database=" & Text2.text
'Exit Sub
 Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Text3.text & ";pwd=" & Text4.text & ";Data Source=" & Text1.text & ";database=" & Text2.text
 Adodc1.RecordSource = "select * from 桌子配置 order by id"
 Adodc1.Refresh
 MsgBox "连接成功__并未保存数据", , App.Title
 
 Exit Sub
CuoWu:                               '增加这行
  MsgBox "连接错误，请重新设置连接内容", , App.Title
End Sub

Private Sub Command3_Click()
    Shell "cmd.exe /c taskkill /im  jiaohao.exe", vbMinimizedNoFocus
    End
End Sub

Private Sub Command4_Click()
Dim cIndex As Integer
     Open App.Path & "\duanx.ini" For Output As #1
     Write #1, (Trim(Text11.text)) & "VbCrVbLf" & (Trim(Text10.text)) & "VbCrVbLf" & (Trim(Text13.text))
     Close #1
     
     Open App.Path & "\dykz.ini" For Output As #1
     Write #1, (Trim(Text6.text)) & "VbCrVbLf" & (Trim(Text7.text)) & "VbCrVbLf" & (Trim(Text8.text)) & "VbCrVbLf" & (Trim(Text9.text)) & "VbCrVbLf" & (Trim(Text12.text))
     Close #1
     
    Open App.Path & "\my.ini" For Output As #1
    Print #1, UserCode(Text1.text) & "VbCrVbLf" & UserCode(Text2.text) & "VbCrVbLf" & UserCode(Text3.text) & "VbCrVbLf" & UserCode(Text4.text)
    Close #1
    
    Open App.Path & "\port.ini" For Output As #1
     Write #1, (Trim(Text22.text)) & "VbCrVbLf" & (Trim(Text23.text))
     Close #1
     
     Open App.Path & "\QR_Code.ini" For Output As #1
     Write #1, (Trim(Text24.text)) & "VbCrVbLf" & (Trim(Label15.Caption))
     Close #1
Select Case Combo4.text
    Case "左移"
        cIndex = 0
    Case "静止"
        cIndex = 1
End Select
 
    Open App.Path & "\yemianbuzhi.ini" For Output As #1
    Write #1, (Trim(Text14.text)) & "VbCrVbLf" & (Trim(Text15)) & "VbCrVbLf" & Trim(Combo1.text) & "VbCrVbLf" & Trim(Combo2.text) & "VbCrVbLf" & Check1.Value & "VbCrVbLf" & (Trim(Text16)) & "VbCrVbLf" & (Trim(Text17)) & "VbCrVbLf" & (Trim(Text18)) & "VbCrVbLf" & (Trim(Text19)) & "VbCrVbLf" & (Trim(Text20)) & "VbCrVbLf" & (Trim(Text21)) & "VbCrVbLf" & Val(Combo3.ListIndex) & "VbCrVbLf" & Trim(cIndex)
    Close #1
    
    Open App.Path & "\自动确定.ini" For Output As #1
     Write #1, Text25.text
     Close #1
MsgBox "成功保存数据", , App.Title
'Unload Me
Shell "cmd.exe /c taskkill /im  jiaohao.exe", vbMinimizedNoFocus
End
End Sub

Private Sub Command5_Click()
Shell "osk.exe", 1
End Sub

Private Sub Command6_Click()
    Unload Me
    Load LEDpz
    LEDpz.Show 1
    
End Sub

Private Sub Command7_Click()
    y = MsgBox("是否清除数据", vbYesNo + vbQuestion)
    If y = vbYes Then
        Conn.Open sqlcnn
        Conn.Execute "TRUNCATE TABLE led显示"
        Conn.Execute "TRUNCATE TABLE 语音叫号"
        Conn.Close
    End If
End Sub

Private Sub Command8_Click()
y = MsgBox("是否清空未叫", vbYesNo + vbQuestion)
    If y = vbYes Then
        Conn.Open sqlcnn
        Conn.Execute "update 语音叫号 set 状态='2' where (状态='0')"
        Conn.Execute "update 排队列表 set 状态='3' where (状态='0')"
        Conn.Execute "update 语音叫号 set 状态='2' where (状态='0')"
        Conn.Execute "update led显示  set 状态='3' where (状态='2')"
        Conn.Close
    End If
End Sub

Private Sub DataGrid1_DblClick()
If Adodc1.Recordset.RecordCount > 0 Then
  jicxxi.Label1(12).Caption = Adodc1.Recordset.Fields("id")
Else
   jicxxi.Label1(12).Caption = 0
End If
  Load jicxxi
  jicxxi.Show 1


End Sub


Private Sub Form_Activate()
On Error Resume Next
 Adodc1.RecordSource = "select * from 桌子配置 order by id"
 Adodc1.Refresh
End Sub

Private Sub Form_Load()
    Combo1.AddItem "1" '
    Combo1.AddItem "2"
    Combo1.AddItem "3"
    Combo1.AddItem "4"
    Combo1.AddItem "5"
    Combo1.text = Combo1.List(0)
    Combo2.AddItem "1"
    Combo2.AddItem "2"
    Combo2.AddItem "3"
    Combo2.text = Combo2.List(0)
    Combo3.AddItem "手动0"
    For i = 1 To 20
        If Dir(App.Path & "\muban" & Trim(i) & ".ini") <> "" Then
            Combo3.AddItem "模板" & Trim(i)
        End If
    Next i
    Combo4.AddItem "左移"
    Combo4.AddItem "静止"
'On Error GoTo CuoWu    '增加这行

Dim s As String, text() As String, row1 As Integer, column1 As Integer
'Close
    If Dir(App.Path & "\my.ini") <> "" Then
  
    Else
        Open App.Path & "\my.ini" For Output As #1
        Write #1, UserCode(Text1.text) & "VbCrVbLf" & UserCode(Text2.text) & "VbCrVbLf" & UserCode(Text3.text) & "VbCrVbLf" & UserCode(Text4.text)
         Close #1
     End If
Open App.Path & "\my.ini" For Input As #1
Input #1, s

    text = Split(s, "VbCrVbLf")
    Text1.text = UserDeCode(text(0))
    Text2.text = UserDeCode(text(1))
    Text3.text = UserDeCode(text(2))
    Text4.text = UserDeCode(text(3))
    Close #1
sqlcnn = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Text3.text & ";pwd=" & Text4.text & ";Data Source=" & Text1.text & ";database=" & Text2.text


Adodc1.ConnectionString = sqlcnn
 
 If Dir(App.Path & "\media_volum.ini") <> "" Then
 Else
    Open App.Path & "\media_volum.ini" For Output As #1
    Write #1, "0" & "VbCrVbLf" & "0"
    Close #1
End If
    
 If Dir(App.Path & "\duanx.ini") <> "" Then
  
 Else
     Open App.Path & "\duanx.ini" For Output As #1
     Write #1, ("8") & "VbCrVbLf" & ("请稍等片刻马上就能到您入座啦") & "VbCrVbLf" & ("13800100500") & "VbCrVbLf" & "4"
     Close #1
 End If
 
Open App.Path & "\duanx.ini" For Input As #1
Input #1, s

text = Split(s, "VbCrVbLf")

Text11.text = (text(0))
Text10.text = (text(1))
Text13.text = (text(2))
Close #1

If Dir(App.Path & "\port.ini") <> "" Then
  
 Else
     Open App.Path & "\port.ini" For Output As #1
     Write #1, ("1") & "VbCrVbLf" & ("2")
     Close #1
 End If
 
Open App.Path & "\port.ini" For Input As #1
Input #1, s

text = Split(s, "VbCrVbLf")

Text22.text = (text(0))
Text23.text = (text(1))

Close #1


If Dir(App.Path & "\rc_code.ini") <> "" Then
 Else
     Open App.Path & "\rc_code.ini" For Output As #1
     Write #1, ("http://www.baidu.com") & "VbCrVbLf" & ("a123456789123456")
     Close #1
 End If
 
Open App.Path & "\rc_code.ini" For Input As #1
Input #1, s

text = Split(s, "VbCrVbLf")

Text24.text = (text(0))
Label15.Caption = (text(1))
Close #1



 If Dir(App.Path & "\dykz.ini") <> "" Then
  
 Else
     Open App.Path & "\dykz.ini" For Output As #1
     Write #1, (Trim(Text6.text)) & "VbCrVbLf" & (Trim(Text7.text)) & "VbCrVbLf" & (Trim(Text8.text)) & "VbCrVbLf" & (Trim(Text9.text)) & "VbCrVbLf" & (Trim(Text12.text))
     Close #1
 End If
Open App.Path & "\dykz.ini" For Input As #1
Input #1, s

text = Split(s, "VbCrVbLf")

Text6.text = (text(0))
Text7.text = (text(1))
Text8.text = (text(2))
Text9.text = (text(3))
Text12.text = (text(4))

Close #1



If Dir(App.Path & "\yemianbuzhi.ini") <> "" Then
Else
    Open App.Path & "\yemianbuzhi.ini" For Output As #1
    Write #1, "4000" & "VbCrVbLf" & "2000" & "VbCrVbLf" & "4" & "VbCrVbLf" & "1" & "VbCrVbLf" & "0" & "VbCrVbLf" & "2000" & "VbCrVbLf" & "122" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "0" & "VbCrVbLf" & "4" & "VbCrVbLf" & "0"
    Close #1
End If
Open App.Path & "\yemianbuzhi.ini" For Input As #1
Input #1, s
text = Split(s, "VbCrVbLf")
Close #1


If text(11) = 1 Then
    Open App.Path & "\muban1.ini" For Input As #1
    Input #1, s
    text = Split(s, "VbCrVbLf")
    Close #1
End If
If text(11) = 2 Then
        Open App.Path & "\muban2.ini" For Input As 1#
        Input #1, s
            text = Split(s, "VbCrVbLf")
        Close #1
End If
If text(11) = 3 Then
        Open App.Path & "\muban3.ini" For Input As 1#
        Input #1, s
            text = Split(s, "VbCrVbLf")
        Close #1
End If
If text(11) = 4 Then
        Open App.Path & "\muban4.ini" For Input As 1#
        Input #1, s
            text = Split(s, "VbCrVbLf")
        Close #1
End If
If text(11) = 5 Then
        Open App.Path & "\muban5.ini" For Input As 1#
        Input #1, s
            text = Split(s, "VbCrVbLf")
        Close #1
End If
        Text14.text = text(0)
        Text15.text = text(1)
        Combo1.text = Combo1.List(text(2) - 1)
        Combo2.text = Combo1.List(text(3) - 1)
        Check1.Value = text(4)
        Text16.text = text(5)
        Text17.text = text(6)
        Text18.text = text(7)
        Text19.text = text(8)
        Text20.text = text(9)
        Text21.text = text(10)
        Combo3.text = Combo3.List(text(11))
        Combo4.text = Combo4.List(text(12))


     If Dir(App.Path & "\自动确定.ini") <> "" Then
    Else
        Open App.Path & "\自动确定.ini" For Output As #1
            Write #1, "1"
        Close #1
    End If
 
    Open App.Path & "\自动确定.ini" For Input As #1
        Input #1, s
        Text25.text = s
    Close #1
 

Exit Sub

CuoWu:                               '增加这行
Close #1
    Text1.text = ""
    Text2.text = ""
    Text3.text = ""
    Text4.text = ""










 
 
 
 
 
 

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label15_DblClick()
Me.Height = 8910
Me.Top = 100
End Sub

Private Sub Label2_DblClick(Index As Integer)
If Index = 0 Then
   msg = MsgBox("保存模板", vbQuestion + vbYesNo, App.Title)
    If msg = vbYes Then
        Select Case Combo4.text
        Case "左移"
            cIndex = 0
        Case "静止"
            cIndex = 1
        End Select
        For i = 1 To 6
            If Dir(App.Path & "\muban" & Trim(i) & ".ini") = "" Then
                Open App.Path & "\muban" & Trim(i) & ".ini" For Output As #1
                Write #1, (Trim(Text14.text)) & "VbCrVbLf" & (Trim(Text15)) & "VbCrVbLf" & Trim(Combo1.text) & "VbCrVbLf" & Trim(Combo2.text) & "VbCrVbLf" & Check1.Value & "VbCrVbLf" & (Trim(Text16)) & "VbCrVbLf" & (Trim(Text17)) & "VbCrVbLf" & (Trim(Text18)) & "VbCrVbLf" & (Trim(Text19)) & "VbCrVbLf" & (Trim(Text20)) & "VbCrVbLf" & (Trim(Text21)) & "VbCrVbLf" & Val(Combo3.ListIndex) & "VbCrVbLf" & Trim(cIndex)
                Close #1
                Exit For
            End If
        Next i
    End If
  End If
End Sub

