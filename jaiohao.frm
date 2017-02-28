VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form jaiohao 
   Caption         =   "½ÐºÅ"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   4410
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "±£´æ"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6120
      Top             =   5520
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
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
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
      Bindings        =   "jaiohao.frx":0000
      Height          =   3615
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6376
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
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
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6000
      Top             =   4320
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6120
      Top             =   5040
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
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "½ÐºÅ´ÎÊý"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "jaiohao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlcnn As String
Dim Conn As New ADODB.Connection
            Dim Rs As New ADODB.Recordset
Dim flag As Integer

Private Sub Command1_Click()
    If Dir(App.Path & "\callCount.ini") <> "" Then
        Open App.Path & "\callCount.ini" For Output As #1
            Write #1, Text1.text
        Close #1
        End
    End If
End Sub

Private Sub Form_Load()
flag = 0
Dim text() As String
    If Dir(App.Path & "\my.ini") <> "" Then
        Open App.Path & "\my.ini" For Input As #1
        Input #1, s
        text = Split(s, "VbCrVbLf")
        sqlcnn = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UserDeCode(text(2)) & ";pwd=" & UserDeCode(text(3)) & ";Data Source=" & UserDeCode(text(0)) & ";database=" & UserDeCode(text(1))
        Close #1
        Adodc1.ConnectionString = sqlcnn
        Adodc2.ConnectionString = sqlcnn
    Else
        
        
        
    End If
    
    


    If Dir(App.Path & "\callCount.ini") <> "" Then
        Open App.Path & "\callCount.ini" For Input As #1
        Input #1, s
        Text1.text = s
        Close #1
    End If
    Me.BackColor = RGB(27, 146, 108)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
'Dim xingx(12) As String, yyjh(7) As String
Dim bhao As String
Adodc1.RecordSource = "select top 1 id,±àºÅ,Ç°½ÐºÅ,ºó½ÐºÅ from ÓïÒô½ÐºÅ where ×´Ì¬='2'order by id"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then

    bhao = Adodc1.Recordset.Fields("±àºÅ")
    bhao = Right("000" & Trim(bhao), 4)
    PlayWavFile App.Path & "\ÉùÒô\" & Trim(Adodc1.Recordset.Fields("Ç°½ÐºÅ")), 1, 0
    For ccc = 1 To Len(bhao)
        PlayWavFile App.Path & "\ÉùÒô\" & Mid(bhao, ccc, 1) & ".wav", 1, 0
    Next ccc
    PlayWavFile App.Path & "\ÉùÒô\" & Trim(Adodc1.Recordset.Fields("ºó½ÐºÅ")), 1, 0
  flag = flag + 1
    If flag >= Text1.text Then
        SQL = "update ÓïÒô½ÐºÅ set ×´Ì¬='3' where (id=" & Adodc1.Recordset.Fields("id") & ")"  '½«×´Ì¬ÉèÖÃÎª3
        Conn.Open sqlcnn
        Conn.Execute SQL
        Conn.Close
       ' MsgBox SQL
        flag = 0
        Adodc1.Refresh
    End If
Else
        '²¥·Å¹ã¸æ
        Open App.Path & "\media_volum.ini" For Output As #1
        Write #1, "0" & "VbCrVbLf" & "60"
        Close #1
    
End If
'    Adodc1.RecordSource = "SELECT *  FROM ÅÅ¶ÓÁÐ±í where (ÈÕÆÚ = CONVERT(DATETIME, '" & Date & " 00:00:00', 102)) and (×´Ì¬='2') order by ±àºÅ"
'    Adodc1.Refresh
'    If Adodc1.Recordset.RecordCount > 0 Then
'        Adodc1.Recordset.MoveFirst
'        For j = 0 To 12
'            xingx(j) = Adodc1.Recordset.Fields(j)
'        Next j
'        SQL = "update ÅÅ¶ÓÁÐ±í set ×´Ì¬='3' where (±àºÅ=" & xingx(12) & ")"  '½«×´Ì¬ÉèÖÃÎª3
'        Conn.Open sqlcnn
'        Conn.Execute SQL
'        Conn.Close
'        If Len(xingx(12)) = 3 Then xingx(12) = Right("00" & xingx(12), 4)  '½«3Î»±àºÅ¸Ä³É4Î»±àºÅ
'        PlayWavFile App.Path & "\ÉùÒô\" & Trim(xingx(7)), 1, 0     'Æô¶¯½ÐºÅ
'        For ccc = 1 To Len(xingx(12))
'           PlayWavFile App.Path & "\ÉùÒô\" & Mid(Trim(xingx(12)), ccc, 1) & ".wav", 1, 0
'        Next ccc
'        PlayWavFile App.Path & "\ÉùÒô\" & Trim(xingx(8)), 1, 0
'    Else
'        '²¥·Å¹ã¸æ
'        Open App.Path & "\media_volum.ini" For Output As #1
'        Write #1, "0" & "VbCrVbLf" & "60"
'        Close #1
'    End If
    
    
'    Adodc2.RecordSource = "SELECT *  FROM ÓïÒô½ÐºÅ where (×´Ì¬='2') order by ±àºÅ"
'    Adodc2.Refresh
'    If Adodc2.Recordset.RecordCount > 0 Then
'        Adodc2.Recordset.MoveFirst
'        For k = 0 To 7
'            yyjh(k) = Adodc2.Recordset.Fields(k)
'        Next k
'        SQL = "UPDATE ÓïÒô½ÐºÅ SET ×´Ì¬ = '3' WHERE (±àºÅ= " & Val(yyjh(2)) & "))"  '½«×´Ì¬ÉèÎª3  (ÈÕÆÚ = CONVERT(DATETIME, '" & Date & " ') and
'        Conn.Open sqlcnn
'        Conn.Execute SQL
'        Conn.Close
'        If Len(yyjh(2)) = 3 Then yyjh(2) = Right("00" & yyjh(2), 4)  '½«3Î»±àºÅ¸Ä³É4Î»±àºÅ
'        PlayWavFile App.Path & "\ÉùÒô\" & Trim(yyjh(4)), 1, 0     'Æô¶¯½ÐºÅ
'        For l = 1 To Len(yyjh(2))
'           PlayWavFile App.Path & "\ÉùÒô\" & Mid(Trim(yyjh(2)), l, 1) & ".wav", 1, 0
'        Next l
'        PlayWavFile App.Path & "\ÉùÒô\" & Trim(yyjh(5)), 1, 0
'    Else
'        '²¥·Å¹ã¸æ
'        Open App.Path & "\media_volum.ini" For Output As #1
'        Write #1, "0" & "VbCrVbLf" & "60"
'        Close #1
'    End If
    
End Sub
