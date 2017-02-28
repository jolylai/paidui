VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form ykpz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "遥控配置"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   4440
      TabIndex        =   31
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2160
      Top             =   9360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   3000
      TabIndex        =   29
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   1560
      TabIndex        =   28
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   120
      TabIndex        =   27
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   5880
      TabIndex        =   26
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   4440
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "保存"
      Height          =   1215
      Left            =   5880
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除"
      Height          =   855
      Left            =   5880
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新增"
      Height          =   975
      Left            =   5880
      TabIndex        =   9
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3120
      Top             =   9480
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "遥控配置.frx":0000
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
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
   Begin VB.Label Label14 
      Caption         =   "id"
      Height          =   495
      Left            =   4440
      TabIndex        =   30
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "五星评价"
      Height          =   495
      Left            =   3000
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "四星评价"
      Height          =   495
      Left            =   1560
      TabIndex        =   23
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "三星评价"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "二星评价"
      Height          =   495
      Left            =   5880
      TabIndex        =   21
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "一星评价"
      Height          =   495
      Left            =   4440
      TabIndex        =   20
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "后叫号"
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "前叫号"
      Height          =   495
      Left            =   1560
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "后文"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "前文"
      Height          =   495
      Left            =   5880
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "屏号"
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "下一位"
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "重叫"
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "设备编码"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "ykpz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COM_STRING As String, control As Integer
Private Sub Command1_Click()
    Me.Height = 7020

    Text1.SetFocus
End Sub

Private Sub Command2_Click()
 msg = MsgBox("确定删除", vbQuestion + vbYesNo, App.Title)
     If msg = vbYes Then
        If Not Adodc1.Recordset.BOF And Not Adodc1.Recordset.EOF Then
            Adodc1.Recordset.Delete adAffectCurrent
            Adodc1.Recordset.Update
        End If
    End If
End Sub

Private Sub Command3_Click()
    Me.Height = 3750
    If Not Adodc1.Recordset.BOF And Not Adodc1.Recordset.EOF Then  '修改
        Adodc1.RecordSource = "select * from 遥控配置"
        Adodc1.Refresh
        Adodc1.Recordset.Fields("名称") = Trim(Text14.text)
        Adodc1.Recordset.Fields("设备编码") = Trim(Text1.text)
        Adodc1.Recordset.Fields("重叫") = Trim(Text2.text)
        Adodc1.Recordset.Fields("下一位") = Trim(Text3.text)
        Adodc1.Recordset.Fields("屏号") = Trim(Text4.text)
        Adodc1.Recordset.Fields("前文") = Trim(Text5.text)
        Adodc1.Recordset.Fields("后文") = Trim(Text6.text)
        Adodc1.Recordset.Fields("前叫号") = Trim(Text7.text)
        Adodc1.Recordset.Fields("后叫号") = Trim(Text8.text)
        Adodc1.Recordset.Fields("一星评价") = Trim(Text9.text)
        Adodc1.Recordset.Fields("二星评价") = Trim(Text10.text)
        Adodc1.Recordset.Fields("三星评价") = Trim(Text11.text)
        Adodc1.Recordset.Fields("四星评价") = Trim(Text12.text)
        Adodc1.Recordset.Fields("五星评价") = Trim(Text13.text)
        Adodc1.Recordset.Update
        MsgBox "保存成功", vbOKOnly, App.Title
    Else
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields("名称") = Trim(Text14.text)
        Adodc1.Recordset.Fields("设备编码") = Trim(Text1.text)
        Adodc1.Recordset.Fields("重叫") = Trim(Text2.text)
        Adodc1.Recordset.Fields("下一位") = Trim(Text3.text)
        Adodc1.Recordset.Fields("屏号") = Trim(Text4.text)
        Adodc1.Recordset.Fields("前文") = Trim(Text5.text)
        Adodc1.Recordset.Fields("后文") = Trim(Text6.text)
        Adodc1.Recordset.Fields("前叫号") = Trim(Text7.text)
        Adodc1.Recordset.Fields("后叫号") = Trim(Text8.text)
        Adodc1.Recordset.Fields("一星评价") = Trim(Text9.text)
        Adodc1.Recordset.Fields("二星评价") = Trim(Text10.text)
        Adodc1.Recordset.Fields("三星评价") = Trim(Text11.text)
        Adodc1.Recordset.Fields("四星评价") = Trim(Text12.text)
        Adodc1.Recordset.Fields("五星评价") = Trim(Text13.text)
        Adodc1.Recordset.Update
        MsgBox "保存成功", vbOKOnly, App.Title
    End If
    
    Unload Me
    Load jicxxi
    jicxxi.Show 1
    
    
End Sub

Private Sub DataGrid1_Click()
'Me.Height = 7020
'    If Not Adodc1.Recordset.BOF And Not Adodc1.Recordset.EOF Then  '修改
'        Adodc1.RecordSource = "select * from 遥控配置"
'        Adodc1.Refresh
'        Text1.text = Adodc1.Recordset.Fields("设备编码")
'        Text2.text = Adodc1.Recordset.Fields("重叫")
'        Text3.text = Adodc1.Recordset.Fields("下一位")
'        Text4.text = Adodc1.Recordset.Fields("屏号")
'        Text5.text = Adodc1.Recordset.Fields("前文")
'        Text6.text = Adodc1.Recordset.Fields("后文")
'        Text7.text = Adodc1.Recordset.Fields("前叫号")
'        Text8.text = Adodc1.Recordset.Fields("后叫号")
'        Text9.text = Adodc1.Recordset.Fields("一星评价")
'        Text10.text = Adodc1.Recordset.Fields("二星评价")
'        Text11.text = Adodc1.Recordset.Fields("三星评价")
'        Text12.text = Adodc1.Recordset.Fields("四星评价")
'        Text13.text = Adodc1.Recordset.Fields("五星评价")
'        Text14.text = Adodc1.Recordset.Fields("名称")
'    End If
End Sub

Private Sub Form_Load()
Dim s As String, text() As String
    Text1.Visible = True
    Text2.Visible = True
    Text3.Visible = True
    Text4.Visible = True
    Text5.Visible = True
    Text6.Visible = True
    Text7.Visible = True
    Text8.Visible = True
    Text9.Visible = True
    Text10.Visible = True
    Text11.Visible = True
    Text12.Visible = True
    Text13.Visible = True
    Text14.Visible = True
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    Label9.Visible = True
    Label10.Visible = True
    Label11.Visible = True
    Label12.Visible = True
    Label13.Visible = True
    Label14.Visible = True
control = 0
Adodc1.ConnectionString = sqlcnn
    Adodc1.RecordSource = "select * from 遥控配置"
    Adodc1.Refresh
    
    If zye.MSComm2.PortOpen = True Then
        zye.MSComm2.PortOpen = False
    End If
    
    Open App.Path & "\port.ini" For Input As #1
    Input #1, s
    text = Split(s, "VbCrVbLf")
    Close #1

 With MSComm1  'led端口
                If .PortOpen = True Then
                .PortOpen = False
                End If
                .CommPort = text(0)
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
    
    Open App.Path & "\telecontrol.ini" For Input As #1
    Input #1, s
    Text14.text = s
    Close #1
   
End Sub

Private Sub MSComm1_OnComm()
Dim Buffer() As Byte, fbuffer() As String, mstring, remote_control_vale As String, remote_control_code As String
If Text1.Visible = True Then
    Select Case MSComm1.CommEvent
        Case 2
            Buffer = MSComm1.Input
            MSComm1.InBufferCount = 0

    For i = 0 To UBound(Buffer)
        If Len(Hex(Buffer(i))) = 1 Then
            COM_STRING = COM_STRING & "0" & Hex(Buffer(i)) & " "
        Else
            COM_STRING = COM_STRING & Hex(Buffer(i)) & " "
        End If
    
    
        If Hex(Buffer(i)) = "A" Then
            MSComm1.PortOpen = False
            If InStr(COM_STRING, "3A 2A 48 46 3D") > 0 And InStr(COM_STRING, "0A") > 0 And InStr(COM_STRING, "0A") - InStr(COM_STRING, "3A 2A 48 46 3D") = 42 Then  '接收遥控器处理
        
                COM_STRING = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D"), InStr(COM_STRING, "0A") - InStr(COM_STRING, "3A 2A 48 46 3D") + 2)  '提取完整字符串
            End If
                          
            Select Case control
                Case 0
                    remote_control_code = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D") + 15, 17) '遥控编码
                    remote_control_code = Replace(remote_control_code, " ", "")
                    Text1.text = Trim(remote_control_code)
                    Text2.SetFocus
                    control = control + 1
                Case 1
                    remote_control_vale = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D") + 33, 5)  '遥控键码
                    remote_control_vale = Replace(remote_control_vale, " ", "")
                    Text2.text = remote_control_vale
                    Text3.SetFocus
                    control = control + 1
                Case 2
                    remote_control_vale = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D") + 33, 5)  '遥控键码
                    remote_control_vale = Replace(remote_control_vale, " ", "")
                    Text3.text = remote_control_vale
                    Text9.SetFocus
                    control = control + 1
                Case 3
                    remote_control_vale = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D") + 33, 5)  '遥控键码
                    remote_control_vale = Replace(remote_control_vale, " ", "")
                    Text9.text = remote_control_vale
                    Text10.SetFocus
                    control = control + 1
                Case 4
                    remote_control_vale = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D") + 33, 5)  '遥控键码
                    remote_control_vale = Replace(remote_control_vale, " ", "")
                    Text10.text = remote_control_vale
                    Text11.SetFocus
                    control = control + 1
                Case 5
                    remote_control_vale = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D") + 33, 5)  '遥控键码
                    remote_control_vale = Replace(remote_control_vale, " ", "")
                    Text11.text = remote_control_vale
                    Text12.SetFocus
                    control = control + 1
                Case 6
                    remote_control_vale = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D") + 33, 5)  '遥控键码
                    remote_control_vale = Replace(remote_control_vale, " ", "")
                    Text12.text = remote_control_vale
                    Text13.SetFocus
                    control = control + 1
                Case 7
                    remote_control_vale = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D") + 33, 5)  '遥控键码
                    remote_control_vale = Replace(remote_control_vale, " ", "")
                    Text13.text = remote_control_vale
                    Text4.SetFocus
'
            End Select
'            If Text1.text <> "" Then
'                If Text2.text <> "" Then
'                    remote_control_vale = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 FF") + 33, 5)  '遥控键码
'                    remote_control_vale = Replace(remote_control_vale, " ", "")
'                    Text3.text = remote_control_vale
'                    If Text3.text <> "" Then
'                        Text4.SetFocus
'                    End If
'                Else
'                    remote_control_vale = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 FF") + 33, 5)  '遥控键码\
'                    remote_control_vale = Replace(remote_control_vale, " ", "")
'                    Text2.text = remote_control_vale
'                    If Text2.text <> "" Then
'                        Text3.SetFocus
'                    End If
'                End If
'
'            Else
'
'                remote_control_code = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 FF") + 15, 17) '遥控编码
'                remote_control_code = Replace(remote_control_code, " ", "")
'                Text1.text = Trim(remote_control_code)
'                If Text1.text <> "" Then
'                    Text2.SetFocus
'                End If
'            End If
            COM_STRING = ""
            MSComm1.PortOpen = True
            Exit For
        End If
    Next i

    End Select
End If
End Sub
