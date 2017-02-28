VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form qued 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "艾升科技"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   18
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   7185
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6720
      Top             =   10320
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   4920
      Top             =   11400
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
      Caption         =   "Adodc5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3360
      Top             =   11760
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
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   1560
      Top             =   11640
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
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   5520
      Top             =   10320
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   855
      Left            =   11160
      TabIndex        =   22
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   170000387
      UpDown          =   -1  'True
      CurrentDate     =   42592
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   855
      Left            =   9600
      TabIndex        =   21
      Top             =   6240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1508
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   0
      Format          =   169934851
      UpDown          =   -1  'True
      CurrentDate     =   42592
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   9600
      TabIndex        =   8
      Top             =   1680
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   1320
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   18
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   2280
         TabIndex        =   17
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   3240
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   4200
         TabIndex        =   15
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   5160
         TabIndex        =   14
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   1320
         TabIndex        =   13
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   2280
         TabIndex        =   12
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   3240
         TabIndex        =   11
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   4200
         TabIndex        =   10
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   5160
         TabIndex        =   9
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "输入手机号"
         Height          =   615
         Left            =   2160
         TabIndex        =   20
         Top             =   120
         Width           =   3495
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4800
      Top             =   10200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3360
      Top             =   10440
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1800
      Top             =   10440
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
   Begin VB.CommandButton Command4 
      Caption         =   "回退首页......."
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      Top             =   8160
      Width           =   4695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "重置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   1
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "确认"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Image Image5 
      Height          =   975
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   972
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   732
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   720
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   720
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   720
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3840
      TabIndex        =   28
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1560
      TabIndex        =   27
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "预约时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   26
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "类别"
      Height          =   495
      Left            =   1680
      TabIndex        =   25
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3360
      TabIndex        =   24
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   23
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   10080
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "人等候"
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "当前"
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1215
      Index           =   0
      Left            =   840
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "qued"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim Conn As New ADODB.Connection
  Dim textcc() As String
            Dim Rs As New ADODB.Recordset
    Dim downcount As Integer
    Dim auto As Integer

Private Sub Command1_Click(Index As Integer)
Text1.text = Trim(Text1.text) & Trim(Index)
Text1.SelStart = Len(Trim(Text1.text))
Text1.SetFocus
End Sub

Private Sub Command2_Click()

Dim i, a(2) As Byte, j(0) As Byte, b(2) As Byte, d(1) As Byte, e(30) As Byte, biaos, yysj As String, shijian As String, rc_code As String, s As String, strHexSj As String, company As String, bytSj() As Byte
'On Error Resume Next
Dim bianhao As String
Dim QR_Code_strin As String, QR_Code_hex() As Byte, STR() As String
'    yysj = Label12.Caption & Label13.Caption
''    shijian = Split(Label11.Caption, ":")
''    For i = 0 To 1
''        If Len(shijian(i)) = 1 Then shijian(i) = Right("00" & shijian(i), 2)
''        yysj = yysj & shijian(i)
''    Next i
'    If yysj - Format(Time, "hhmm") > 0 Then
'        bianhao = yysj
'    Else
'        If yysj = Format(Time, "hhmm") Then
'            bianhao = Format(Time, "hhmm")
'        Else
'            MsgBox "预约时间必须在当前时间之后，请重新预约"
'            Unload Me
'        End If
'    End If
'    MsgBox yysj
'bianhao = Format(Time, "hhmm")
bianhao = Label12.Caption & Label13.Caption
 Adodc1.RecordSource = "SELECT 编号,日期  FROM 排队列表 where (日期 = CONVERT(DATETIME, '" & Date & " 00:00:00', 102)) and (编号 >= " & Val(bianhao) & ")"
 Adodc1.Refresh
 If Adodc1.Recordset.RecordCount > 0 Then
  bianhao = Val(bianhao) + 1
   For i = 0 To 50
          Adodc1.Recordset.MoveFirst  '第一行
          For n = 0 To Adodc1.Recordset.RecordCount - 1
             If Val(Adodc1.Recordset.Fields("编号")) = Val(bianhao) Then
                 biaos = "有存在"
                 bianhao = Val(bianhao) + 1
                 n = Adodc1.Recordset.RecordCount + 1
               
             Else
             
                 Adodc1.Recordset.MoveNext
        
             End If
          Next n
          If biaos = "有存在" Then
          
          Else
            i = 51
          End If
   
  Next i
  

 Else

 End If
 
' bianhao = bianhao + 1

' Loop While Adodc1.Recordset.RecordCount > 0





'Adodc1.RecordSource = "SELECT MAX(编号) AS 最大编号 FROM 排队列表 GROUP BY 日期 HAVING (日期 = CONVERT(DATETIME, '" & Date & " 00:00:00', 102))"
'Adodc1.Refresh

'If Not Adodc1.Recordset.EOF And Not Adodc1.Recordset.EOF Then
'   bianhao = Val(Adodc1.Recordset.Fields(0)) + 1
'Else
'  bianhao = 1
'  sql = "UPDATE 桌子配置 SET 当前已用 = 0"
'             Conn.Open sqlcnn
'             Conn.Execute sql
'             Conn.Close
'End If

Adodc1.RecordSource = "SELECT id, 名称, 数量, 当前已用, 绑定无线, 前叫号文件, 后叫号文件, 图片, 屏号, 前文,后文, 状态 FROM 桌子配置 WHERE (id = " & Val(Label5.Caption) & ")"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
   Adodc2.RecordSource = "SELECT id, 座位id, 电话号码, 日期,时间, 状态, 绑定无线, 前叫号文件, 后叫号文件, 屏号,前文, 后文,编号 FROM 排队列表"
   Adodc2.Refresh
   Adodc2.Recordset.AddNew
   Adodc2.Recordset.Fields("座位id") = Adodc1.Recordset.Fields("id")
   Adodc2.Recordset.Fields("电话号码") = Trim(Text1.text)
   Adodc2.Recordset.Fields("日期") = Date
   Adodc2.Recordset.Fields("时间") = Time
   
   Adodc2.Recordset.Fields("状态") = "0"
   Adodc2.Recordset.Fields("绑定无线") = Trim(Adodc1.Recordset.Fields("绑定无线"))
   Adodc2.Recordset.Fields("前叫号文件") = Adodc1.Recordset.Fields("前叫号文件")
   Adodc2.Recordset.Fields("后叫号文件") = Adodc1.Recordset.Fields("后叫号文件")
   Adodc2.Recordset.Fields("屏号") = Adodc1.Recordset.Fields("屏号")
   Adodc2.Recordset.Fields("前文") = Adodc1.Recordset.Fields("前文")
   Adodc2.Recordset.Fields("后文") = Adodc1.Recordset.Fields("后文")
   
   If Len(bianhao) = 3 Then bianhao = Right("00" & Trim(bianhao), 4)
   Adodc2.Recordset.Fields("编号") = bianhao
   Adodc2.Recordset.Update
End If

'Adodc3.RecordSource = "select * from 遥控配置 WHERE (名称 = " & Val(Label5.Caption) & ")"
'Adodc3.Refresh
'If Adodc3.Recordset.RecordCount > 0 Then
'    Adodc4.RecordSource = "select * from led显示"
'    Adodc4.Refresh
'    Adodc4.Recordset.AddNew
'    Adodc4.Recordset.Fields("名称") = Adodc3.Recordset.Fields("id")
'    Adodc4.Recordset.Fields("编号") = bianhao
'    Adodc4.Recordset.Fields("状态") = "0"
'    Adodc4.Recordset.Fields("日期") = Date
'    Adodc4.Recordset.Fields("前文") = Adodc3.Recordset.Fields("前文")
'    Adodc4.Recordset.Fields("后文") = Adodc3.Recordset.Fields("后文")
'    Adodc4.Recordset.Fields("屏号") = Adodc3.Recordset.Fields("屏号")
'    Adodc4.Recordset.Fields("设备编码") = Adodc3.Recordset.Fields("设备编码")
'    Adodc4.Recordset.Update
'
'    Adodc5.RecordSource = "select * from 语音叫号"
'    Adodc5.Refresh
'    Adodc5.Recordset.AddNew
'    Adodc5.Recordset.Fields("名称") = Adodc3.Recordset.Fields("id")
'    Adodc5.Recordset.Fields("编号") = bianhao
'    Adodc5.Recordset.Fields("状态") = "0"
'    Adodc5.Recordset.Fields("日期") = Date
'    Adodc5.Recordset.Fields("前叫号") = Adodc3.Recordset.Fields("前叫号")
'    Adodc5.Recordset.Fields("后叫号") = Adodc3.Recordset.Fields("后叫号")
'    Adodc5.Recordset.Fields("设备编码") = Adodc3.Recordset.Fields("设备编码")
'    Adodc5.Recordset.Update
'End If


'Text6.text = (textcc(0))  标题
'Text7.text = (textcc(1))  电话
'Text8.text = (textcc(2))  说明
'Text9.text = (textcc(3))  标未


'二维码

     



'Unload Me
'Exit Sub

'打印logo
Dim logo() As String
Dim logoToHex() As Byte
If Dir(App.Path & "\logo.ini") <> "" Then
    Open App.Path & "\logo.ini" For Input As #1
    Input #1, s
    Close #1
    logo = Split(s, " ")
    ReDim logoToHex(UBound(logo))
    For i = 0 To UBound(logo) - 1
        logoToHex(i) = "&H" & logo(i)
    Next
    MSComm1.Output = logoToHex
End If
If Dir(App.Path & "\logo1.ini") <> "" Then
    Open App.Path & "\logo1.ini" For Input As #1
    Input #1, s
    Close #1
    logo = Split(s, " ")
    ReDim logoToHex(UBound(logo))
    For i = 0 To UBound(logo) - 1
        logoToHex(i) = "&H" & logo(i)
    Next
    MSComm1.Output = logoToHex
End If
If Dir(App.Path & "\logo2.ini") <> "" Then
    Open App.Path & "\logo2.ini" For Input As #1
    Input #1, s
    Close #1
    logo = Split(s, " ")
    ReDim logoToHex(UBound(logo))
    For i = 0 To UBound(logo) - 1
        logoToHex(i) = "&H" & logo(i)
    Next
    MSComm1.Output = logoToHex
End If
If Dir(App.Path & "\logo3.ini") <> "" Then
    Open App.Path & "\logo3.ini" For Input As #1
    Input #1, s
    Close #1
    logo = Split(s, " ")
    ReDim logoToHex(UBound(logo))
    For i = 0 To UBound(logo) - 1
        logoToHex(i) = "&H" & logo(i)
    Next
    MSComm1.Output = logoToHex
End If

If Dir(App.Path & "\logo4.ini") <> "" Then
    Open App.Path & "\logo4.ini" For Input As #1
    Input #1, s
    Close #1
    logo = Split(s, " ")
    ReDim logoToHex(UBound(logo))
    For i = 0 To UBound(logo) - 1
        logoToHex(i) = "&H" & logo(i)
    Next
    MSComm1.Output = logoToHex
End If

d(0) = &H1B
'd(1) = &H6D '半切
d(1) = &H69 '全切

b(0) = &H1B
b(1) = &H4A
b(2) = &H15   ' 走纸

j(0) = &HA   '打印并换行

a(0) = &H1B
a(1) = &H40  '初始
MSComm1.Output = a

a(0) = &H1B
a(1) = &H61
a(2) = &H1
MSComm1.Output = a
 
If Trim(textcc(0)) <> "" Then  '打印标题

a(0) = &H1D
a(1) = &H21
a(2) = &H11 '100000
MSComm1.Output = a
MSComm1.Output = Trim(textcc(0))
MSComm1.Output = j
MSComm1.Output = b

End If

a(0) = &H1D
a(1) = &H21
a(2) = &H11 '100000
MSComm1.Output = a
MSComm1.Output = "NO." & bianhao
MSComm1.Output = j
MSComm1.Output = b

MSComm1.Output = Trim(Label6.Caption)
MSComm1.Output = j
MSComm1.Output = b


a(0) = &H1D
a(1) = &H21
a(2) = &H10 '100000
MSComm1.Output = a
MSComm1.Output = "您前面还有：" & Label3.Caption & " 位在等候，请注意语言提示！"
MSComm1.Output = j




Dim k(3) As Byte  '居左对齐
k(0) = &H1B
k(1) = &H61
k(2) = &H0

If Trim(textcc(1)) <> "" Then  '打印标题

a(0) = &H1D
a(1) = &H21
a(2) = &H1 '100000
MSComm1.Output = k
MSComm1.Output = a
MSComm1.Output = Trim(textcc(1))
'MSComm1.Output = j


End If

If Trim(textcc(2)) <> "" Then  '打印标题
MSComm1.Output = k
a(0) = &H1D
a(1) = &H21
a(2) = &H1 '100000
MSComm1.Output = a
MSComm1.Output = Trim(textcc(2))
'MSComm1.Output = j


End If
If Trim(textcc(3)) <> "" Then  '打印标题
MSComm1.Output = k
a(0) = &H1D
a(1) = &H21
a(2) = &H1 '100000
MSComm1.Output = a
MSComm1.Output = Trim(textcc(3))
'MSComm1.Output = j


End If




Open App.Path & "\dykz.ini" For Input As #1
     Input #1, s
     text = Split(s, "VbCrVbLf")
     Close #1
company = text(0)
Open App.Path & "\QR_Code.ini" For Input As #1 '二维码打印
     Input #1, s
     If Len(Trim(s)) > 0 Then  '配置文件不为空
        text = Split(s, "VbCrVbLf")
        Adodc1.RecordSource = "select 名称 from 桌子配置 WHERE (id = " & Val(Label5.Caption) & ")"
        Adodc1.Refresh
        
        rc_code = Trim(text(0)) & "?company=" & Trim(company) & "&name=" & Adodc1.Recordset.Fields("名称") & "&machine_no=" & Trim(text(1))
        rc_code = Replace(rc_code, " ", "")
      
'        bytSj = StrConv(rc_code, vbFromUnicode)  '转换十六进制
'        For i = 0 To UBound(bytSj)
'            strHexSj = strHexSj & Right("0" & Hex(bytSj(i)), 2)
'        Next
        QR_Code_string = qrma(25, rc_code)
'        MsgBox QR_Code_string
        STR = Split(QR_Code_string, " ")
        ReDim QR_Code_hex(UBound(STR))
'        MsgBox STR(17)
        For i = O To UBound(STR)
           QR_Code_hex(i) = CLng("&H" & STR(i))

        Next
        MSComm1.Output = QR_Code_hex

    End If
    Close #1
    
a(0) = &H1D
a(1) = &H21
a(2) = &H1 '100000
MSComm1.Output = a
MSComm1.Output = Format(Now, "yyyy-mm-dd hh:mm:ss")
MSComm1.Output = j


a(0) = &H1D
a(1) = &H21
a(2) = &H0  '100000
MSComm1.Output = a
MSComm1.Output = "厦门市艾升科技技术支持"
MSComm1.Output = j
MSComm1.Output = "维护电话：0592-6029842"
MSComm1.Output = j
MSComm1.Output = b
    
    
MSComm1.Output = b
MSComm1.Output = b
MSComm1.Output = b
MSComm1.Output = b
MSComm1.Output = b

MSComm1.Output = d  '全切
MSComm1.Output = j




Unload Me
End Sub

Private Sub Command3_Click()
    Label12.Caption = Format(Time, "HH")
    Label13.Caption = Right(Format(Time, "hhmm"), 2)
'Text1.text = ""
'Text1.SelStart = Len(Trim(Text1.text))
'Text1.SetFocus
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()

 If Dir(App.Path & "\dykz.ini") <> "" Then
  
 Else
     Open App.Path & "\dykz.ini" For Output As #1
     Write #1, (Trim(Text6.text)) & "VbCrVbLf" & (Trim(Text7.text)) & "VbCrVbLf" & (Trim(Text8.text)) & "VbCrVbLf" & (Trim(Text9.text)) & "VbCrVbLf" & (Trim(Text12.text))
     Close #1
 End If
Open App.Path & "\dykz.ini" For Input As #1
Input #1, s

textcc = Split(s, "VbCrVbLf")
 
'Text6.text = (textcc(0))
'Text7.text = (textcc(1))
'Text8.text = (textcc(2))
'Text9.text = (textcc(3))
'Text12.text = (textcc(4))

Close #1







  With MSComm1   '打印
                If .PortOpen = True Then
                .PortOpen = False
                End If
                .CommPort = Val(textcc(4))
                .settings = "9600,n,8,1"
                .InBufferSize = 1024
                .OutBufferSize = 1024
                
                .InputMode = comInputModeBinary    '设置接收数据模式为文本形式
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

End Sub

Private Sub Form_Load()

    downcount = 30

    If Dir(App.Path & "\自动确定.ini") <> "" Then
    Else
        Open App.Path & "\自动确定.ini" For Output As #1
            Write #1, "1"
        Close #1
    End If
 
    Open App.Path & "\自动确定.ini" For Input As #1
        Input #1, s
        auto = Val(s)
    Close #1
               ' MsgBox (auto)
    Image2.Picture = LoadPicture(App.Path & "\pic\加号.jpg")
    Image3.Picture = LoadPicture(App.Path & "\pic\减号.jpg")
    Image4.Picture = LoadPicture(App.Path & "\pic\加号.jpg")
    Image5.Picture = LoadPicture(App.Path & "\pic\减号.jpg")
    Label12.Caption = Format(Time, "HH")
    Label13.Caption = Right(Format(Time, "hhmm"), 2)
     Me.BackColor = RGB(27, 146, 108) '&1b926c &
'    Command2.BackColor = RGB(23, 126, 93)
'    Me.BackColor = RGB(23, 126, 93)       '#177e5d;
'    Me.Picture = LoadPicture(App.Path & "\pic\配置背景.jpg")

    Adodc1.ConnectionString = sqlcnn
    Adodc2.ConnectionString = sqlcnn
    
     Adodc3.ConnectionString = sqlcnn
      Adodc4.ConnectionString = sqlcnn
     Adodc5.ConnectionString = sqlcnn
    
    
'    DTPicker1.CustomFormat = "HH"
'    DTPicker1.Value = Time
'    DTPicker2.CustomFormat = "mm"
'    DTPicker2.Value = Time
    Timer1.Interval = 500
    Label8.ForeColor = &H8000000F
'    Label11.Caption = Me.DTPicker1.Hour & ":" & Me.DTPicker2.Minute
'
End Sub

Private Sub Image2_Click()
    Label12.Caption = Label12.Caption + 1
    If Val(Label12.Caption) > 23 Then Label12.Caption = Format(Time, "HH")
    If Len(Label12.Caption) = 1 Then Label12.Caption = Right("00" & Label12.Caption, 2)
End Sub

Private Sub Image3_Click()
    Label12.Caption = Label12.Caption - 1
    If Val(Label12.Caption) < Format(Time, "HH") Then Label12.Caption = "23"
    If Len(Label12.Caption) = 1 Then Label12.Caption = Right("00" & Label12.Caption, 2)
End Sub

Private Sub Image4_Click()
    Label13.Caption = Label13.Caption + 1
    If Val(Label12.Caption) > Format(Time, "HH") Then
        If Val(Label13.Caption) > 59 Then Label13.Caption = "00"
    Else
        If Val(Label13.Caption) > 59 Then Label13.Caption = Right(Format(Time, "hhmm"), 2)
    End If
    If Len(Label13.Caption) = 1 Then Label13.Caption = Right("00" & Label13.Caption, 2)
End Sub

Private Sub Image5_Click()
    Label13.Caption = Label13.Caption - 1
    If Val(Label12.Caption) > Format(Time, "HH") Then
        If Val(Label13.Caption) < 0 Then Label13.Caption = 59
    Else
        If Val(Label13.Caption) < Right(Format(Time, "hhmm"), 2) Then Label13.Caption = 59
    End If
    If Len(Label13.Caption) = 1 Then Label13.Caption = Right("00" & Label13.Caption, 2)
End Sub


Private Sub Timer1_Timer()
'    Label11.Caption = Me.DTPicker1.Hour & ":" & Me.DTPicker2.Minute
    If Label8.ForeColor = &H8000000F Then
        Label8.ForeColor = vbBlack
    Else
        Label8.ForeColor = &H8000000F
    End If
End Sub

Private Sub Timer2_Timer()
    If downcount = 0 Then
        Unload qued
    Else
        downcount = downcount - 1
        auto = auto - 1
'MsgBox (auto)
        If auto = 0 Then
                
            Call Command2_Click
            Unload qued
        End If
    End If
End Sub
