VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form zye 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   13605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   24405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13605
   ScaleWidth      =   24405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   18240
      Top             =   9840
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   17400
      Top             =   9840
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   19080
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
      Caption         =   "Adodc8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   17400
      Top             =   10560
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   495
      Left            =   19080
      Top             =   10200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
      Caption         =   "Adodc7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   495
      Left            =   22680
      Top             =   9360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
      Caption         =   "Adodc6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "��ҳ.frx":0000
      Height          =   6735
      Left            =   17280
      TabIndex        =   22
      Top             =   2280
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11880
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin MSCommLib.MSComm MSComm2 
      Left            =   16800
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   17520
      Top             =   11280
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   14760
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   975
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   1720
         _cy             =   1508
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9855
      Left            =   12960
      TabIndex        =   12
      Top             =   1920
      Width           =   3735
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   720
         Top             =   3120
      End
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   375
         Left            =   840
         Top             =   6360
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "Adodc5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   600
         Top             =   8880
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   495
         Left            =   720
         Top             =   7440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "��ҳ.frx":0015
         Height          =   3015
         Left            =   600
         TabIndex        =   13
         Top             =   4440
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "��ҳ.frx":002A
         Height          =   2895
         Left            =   600
         TabIndex        =   14
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   840
         Top             =   6840
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
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   600
         Top             =   5160
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
            Name            =   "����"
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
         Left            =   1440
         Top             =   480
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
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Caption         =   "����7"
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   9000
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�˵Ⱥ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   8880
      TabIndex        =   19
      Top             =   12480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�˵Ⱥ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   8880
      TabIndex        =   18
      Top             =   11040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   8400
      TabIndex        =   17
      Top             =   12480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   8400
      TabIndex        =   16
      Top             =   11040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1215
      Index           =   7
      Left            =   2760
      Top             =   12120
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   1215
      Index           =   6
      Left            =   2760
      Top             =   10800
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�˵Ⱥ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   8880
      TabIndex        =   11
      Top             =   9960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   10
      Top             =   9960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�˵Ⱥ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   8880
      TabIndex        =   9
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   8400
      TabIndex        =   8
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�˵Ⱥ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   7
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8400
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�˵Ⱥ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8400
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�˵Ⱥ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8880
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1215
      Index           =   5
      Left            =   2760
      Top             =   9240
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   1215
      Index           =   4
      Left            =   2760
      Top             =   7560
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   1215
      Index           =   3
      Left            =   2760
      Top             =   5760
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   1215
      Index           =   2
      Left            =   2760
      Top             =   4080
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   1215
      Index           =   1
      Left            =   2760
      Top             =   2400
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�˵Ⱥ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8880
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1215
      Index           =   0
      Left            =   2760
      Top             =   720
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "zye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Dim Conn As New ADODB.Connection
        Dim Rs As New ADODB.Recordset
        Dim mURL As Integer, send_led As Integer
Dim xmid(20), xmMC(20), datesss, diycyun, dianjics, dusnax() As String, led_out() As String, strstring, led_no As String
Dim COM_STRING As String '�����������ַ��洢��
Dim timer_var As Integer
Dim led_out_hex(31) As Byte 'led�����ַ���
Dim led_flag As Integer   'led�����ַ�������
Dim led_string1 As String '��ȡ��led�ַ���




Private Sub Form_Activate()
On Error Resume Next
If MSComm2.PortOpen = False Then
    MSComm2.PortOpen = True
End If
    diycyun = 1
End Sub

Private Sub Form_DblClick()
    dianjics = dianjics + 1
End Sub

Private Sub Form_Load()
On Error GoTo CuoWu    '��������
'On Error Resume Next





diycyun = 1
Dim s As String, text() As String
led_flag = 0 '��ʼ��ȫ�ֱ���
    dianjics = 0
    Me.Top = 0
    Me.Left = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    
    If Dir(App.Path & "\duanx.ini") <> "" Then
    Else
        Open App.Path & "\duanx.ini" For Output As #1
            Write #1, ("8") & "VbCrVbLf" & ("���Ե�Ƭ�����Ͼ��ܵ���������") & "VbCrVbLf" & ("13800100500")
        Close #1
    End If
 
    Open App.Path & "\duanx.ini" For Input As #1
        Input #1, s
        dusnax = Split(s, "VbCrVbLf")
        'Text11.text = (dusnax(0)) '�˿ں�
        'Text10.text = (dusnax(1)) '����
        'Text13.text = (dusnax(2)) '��������
    Close #1

    
    Open App.Path & "\my.ini" For Input As #1
        Input #1, s
        text = Split(s, "VbCrVbLf")
    Close #1
sqlcnn = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UserDeCode(text(2)) & ";pwd=" & UserDeCode(text(3)) & ";Data Source=" & UserDeCode(text(0)) & ";database=" & UserDeCode(text(1))
' sqlcnn = "Provider=SQLOLEDB;Password=ais123;Persist Security Info=False;User ID=sa;Initial Catalog=paidui;Data Source=127.0.0.1,1433"
Adodc1.ConnectionString = sqlcnn
Adodc2.ConnectionString = sqlcnn
Adodc3.ConnectionString = sqlcnn
Adodc4.ConnectionString = sqlcnn
Adodc5.ConnectionString = sqlcnn
Adodc7.ConnectionString = sqlcnn
Adodc8.ConnectionString = sqlcnn

'Adodc1.RecordSource = "SELECT ״̬,���,����  FROM �Ŷ��б� where (���� = CONVERT(DATETIME, '" & Date & " 00:00:00', 102)) and (״̬<'2') order by ��� "
'Adodc1.Refresh
'    If Adodc1.Recordset.RecordCount > 0 Then
'    Else
'        SQL = "UPDATE �������� SET ��ǰ���� = 0"
'        Conn.Open sqlcnn
'        Conn.Execute SQL
'        Conn.Close
'    End If


    Me.Picture = LoadPicture(App.Path & "\pic\��ҳ.jpg")
    yemtp

Dim lngWindow As Long  '���нк�exe�ļ�
     lngWindow = FindWindow(vbNullString, "jioah")
     If lngWindow <> 0 Then
     Else
        q = Shell(App.Path & "\jiaohao.exe", vbMinimizedNoFocus)
     End If

 If Dir(App.Path & "\duanx.ini") <> "" Then
 Else
     Open App.Path & "\duanx.ini" For Output As #1
     Write #1, ("8") & "VbCrVbLf" & ("���Ե�Ƭ�����Ͼ��ܵ���������") & "VbCrVbLf" & ("13800100500") & "VbCrVbLf" & Text22.text
     Close #1
 End If

Open App.Path & "\port.ini" For Input As #1
    Input #1, s
    text = Split(s, "VbCrVbLf")
Close #1

 With MSComm1 '�кŶ˿�
        If .PortOpen = True Then
            .PortOpen = False
        End If
        .CommPort = text(1)
        .settings = "9600,n,8,1"
        .InBufferSize = 1024
        .OutBufferSize = 1024
        
        .InputMode = comInputModeBinary    '���ý�������ģʽΪ�ı���ʽ
        '-----------------------------------------------------------------------------------------------------
        .InputLen = 0                     '����Input һ�δӽ��ջ����ȡȫ���ֽ���
        .SThreshold = 0                   '���÷��������в���OnComm�¼�
        .InBufferCount = 0                '������ջ�����
        .OutBufferCount = 0               '������ͻ�����
        .RThreshold = 1                   '���ý���һ���ֽڲ���OnComm�¼�     '
        .RTSEnable = True
        If Not .PortOpen Then             '�ж�ͨ�ſ��Ƿ��
        On Error Resume Next
            .PortOpen = True                '��ͨ�ſ�
        End If
  End With


 With MSComm2  'led�˿�
                If .PortOpen = True Then
                .PortOpen = False
                End If
                .CommPort = text(0)
                .settings = "9600,n,8,1"
                .InBufferSize = 1024
                .OutBufferSize = 1024

                .InputMode = comInputModeBinary    '���ý�������ģʽΪ�ı���ʽ
                '-----------------------------------------------------------------------------------------------------
                .InputLen = 0                     '����Input һ�δӽ��ջ����ȡȫ���ֽ���
                .SThreshold = 0                   '���÷��������в���OnComm�¼�
                .InBufferCount = 0                '������ջ�����
                .OutBufferCount = 0               '������ͻ�����
                .RThreshold = 1                   '���ý���һ���ֽڲ���OnComm�¼�     '
                .RTSEnable = True
                    If Not .PortOpen Then             '�ж�ͨ�ſ��Ƿ��
                  On Error Resume Next
                    .PortOpen = True                '��ͨ�ſ�
                    End If
  End With
  Timer2.Enabled = True
  Timer3.Enabled = True
  
  
Exit Sub

CuoWu:                               '��������
    Close #1
    Load xtpeiz
    xtpeiz.Show 1

End Sub

Private Sub yemtp()
    For i = 0 To 20
        xmid(i) = ""
    Next i

Dim imagetop, labeltop, s As String, Text1() As String

  imagetop = 8000
  labeltop = 5200
      
      '���綩Ʊ  ״̬Ϊ4
Adodc7.RecordSource = "select * from �Ŷ��б� where (���� = CONVERT(DATETIME, '" & Date & " 00:00:00', 102)) and (״̬='4') and (���<=" & Val(Format(Time, "hhmm")) & ")"
Adodc7.Refresh
If Adodc7.Recordset.RecordCount > 0 Then
     SQL = "UPDATE �Ŷ��б� SET ״̬='0' where (���� = CONVERT(DATETIME, '" & Date & " 00:00:00', 102)) and (״̬='4') and (���<=" & Val(Format(Time, "hhmm")) & ") "
    Conn.Open sqlcnn
    Conn.Execute SQL
    Conn.Close
End If
    

Adodc1.RecordSource = "SELECT id, ����, ����, ��ǰ����, ������, ǰ�к��ļ�, ��к��ļ�,ͼƬ, ����, ǰ��, ����,״̬ FROM �������� WHERE (״̬ = '1') ORDER BY id"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then

   For i = 0 To Adodc1.Recordset.RecordCount - 1

         Image1(i).Picture = LoadPicture(App.Path & "\pic\" & Trim(Adodc1.Recordset.Fields("ͼƬ")))
      
      '���� ��ѯ�Ѿ�ȡ�ŵ�����
      
      Adodc4.RecordSource = "SELECT * FROM �Ŷ��б� WHERE (���� = CONVERT(DATETIME, '" & Date & " 00:00:00', 102)) AND (״̬ < '2') and (��λid = " & Adodc1.Recordset.Fields("id") & ") "
      Adodc4.Refresh
      
      Label1(i).Caption = Adodc4.Recordset.RecordCount '- (Val(Adodc1.Recordset.Fields("����")) + Val(Adodc1.Recordset.Fields("��ǰ����")))
      If Label1(i).Caption < 0 Then
         Label1(i).Caption = 0
      End If
            
      xmMC(i) = Trim(Adodc1.Recordset.Fields("����"))
      xmid(i) = Adodc1.Recordset.Fields("id")
      Adodc1.Recordset.MoveNext
   Next
   



    If diycyun = 1 Then            'ҳ�沼��
        If Dir(App.Path & "\yemianbuzhi.ini") <> "" Then
        Else
            Open App.Path & "\yemianbuzhi.ini" For Output As #1
            Write #1, "4000" & "VbCrVbLf" & "2000" & "VbCrVbLf" & "4" & "VbCrVbLf" & "1" & "VbCrVbLf" & "0" & "VbCrVbLf" & "2000" & "VbCrVbLf" & "122" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "0" & "VbCrVbLf" & "0" & "VbCrVbLf" & "4"
            Close #1
        End If
        Open App.Path & "\yemianbuzhi.ini" For Input As #1
        Input #1, s
        Text1 = Split(s, "VbCrVbLf")
        Close #1
        
        
        If Text1(11) = 1 Then
            Open App.Path & "\muban1.ini" For Input As #1
            Input #1, s
            Text1 = Split(s, "VbCrVbLf")
            Close #1
        End If
        If Text1(11) = 2 Then
                Open App.Path & "\muban2.ini" For Input As 1#
                Input #1, s
                    Text1 = Split(s, "VbCrVbLf")
                Close #1
        End If
        If Text1(11) = 3 Then
        Label14.Caption = 6
                Open App.Path & "\muban3.ini" For Input As 1#
                Input #1, s
                    Text1 = Split(s, "VbCrVbLf")
                Close #1
        End If
        If Text1(11) = 4 Then
                Open App.Path & "\muban4.ini" For Input As 1#
                Input #1, s
                    Text1 = Split(s, "VbCrVbLf")
                Close #1
        End If
        If Text1(11) = 5 Then
        Open App.Path & "\muban5.ini" For Input As 1#
        Input #1, s
            Text1 = Split(s, "VbCrVbLf")
        Close #1
        End If
        If Text1(4) = 0 Then
            WindowsMediaPlayer1.Controls.Stop
            Picture1.Visible = False
            WindowsMediaPlayer1.Visible = False
            For k = 0 To (Text1(2) - 1)
                     imagetop = (Screen.Height - 5000) / Text1(2) * k + Text1(0)
                For j = 0 To (Text1(3) - 1)
                    Image1(j + k * Text1(3)).Visible = True
                    Label1(j + k * Text1(3)).Visible = True
                    Label2(j + k * Text1(3)).Visible = True
                    Image1(j + k * Text1(3)).Top = imagetop
                    Image1(j + k * Text1(3)).Left = Screen.Width / Text1(3) * j + Text1(1)
                    Label1(j + k * Text1(3)).Top = Image1(j + k * Text1(3)).Top + Image1(j + k * Text1(3)).Height - 320
                    Label1(j + k * Text1(3)).Left = Image1(j + k * Text1(3)).Left + Image1(j + k * Text1(3)).Width - 1500
                    Label2(j + k * Text1(3)).Top = Image1(j + k * Text1(3)).Top + Image1(j + k * Text1(3)).Height - 360
                    Label2(j + k * Text1(3)).Left = Label1(j + k * Text1(3)).Left + 300
                    If j + k * Text1(3) > Adodc1.Recordset.RecordCount - 1 Then
                        Image1(j + k * Text1(3)).Visible = False
                        Label1(j + k * Text1(3)).Visible = False
                        Label2(j + k * Text1(3)).Visible = False
                    End If
                Next j
            Next k
        End If
        If Text1(4) = 1 Then
            For k = 0 To (Text1(2) - 1)
                     imagetop = Text1(5) * k + Text1(0)
                For j = 0 To (Text1(3) - 1)
                    Image1(j + k * Text1(3)).Visible = True
                    Label1(j + k * Text1(3)).Visible = True
                    Label2(j + k * Text1(3)).Visible = True
                    Image1(j + k * Text1(3)).Top = imagetop
                    Image1(j + k * Text1(3)).Left = Screen.Width / Text1(3) * j + Text1(1)
                    Label1(j + k * Text1(3)).Top = Image1(j + k * Text1(3)).Top + Image1(j + k * Text1(3)).Height - 320
                    Label1(j + k * Text1(3)).Left = Image1(j + k * Text1(3)).Left + Image1(j + k * Text1(3)).Width - 1500
                    Label2(j + k * Text1(3)).Top = Image1(j + k * Text1(3)).Top + Image1(j + k * Text1(3)).Height - 360
                    Label2(j + k * Text1(3)).Left = Label1(j + k * Text1(3)).Left + 300
                    If j + k * Text1(3) > Adodc1.Recordset.RecordCount - 1 Then
                        Image1(j + k * Text1(3)).Visible = False
                        Label1(j + k * Text1(3)).Visible = False
                        Label2(j + k * Text1(3)).Visible = False
                    End If
                Next j
            Next k
                    If Picture1.Visible = False And WindowsMediaPlayer1.Visible = False Then
                        If Dir(App.Path & "\avi\1.avi") <> "" Then   '��洰��
                            Picture1.Visible = True
                            Picture1.Top = Text1(7)
                            Picture1.Left = Text1(8)
                            Picture1.Width = Text1(9)
                            Picture1.Height = Text1(10)
                            WindowsMediaPlayer1.Top = 0
                            WindowsMediaPlayer1.Left = 0
                            WindowsMediaPlayer1.Height = Picture1.Height
                            WindowsMediaPlayer1.Width = Picture1.Width
                            WindowsMediaPlayer1.uiMode = "none"
                            WindowsMediaPlayer1.URL = App.Path & "\avi\1.avi"
                        End If
                    End If
          End If
    End If
End If
                                                                                                                                                                                      
 diycyun = 2

End Sub

Private Sub Image1_Click(Index As Integer)
    Adodc7.RecordSource = "select id,Ӫҵ���� from �������� where (���� ='" & xmMC(Index) & "')"
    Adodc7.Refresh
    Adodc8.RecordSource = "select * from �Ŷ��б� where (��λid ='" & Adodc7.Recordset.Fields("id") & "') and (���� = CONVERT(DATETIME, '" & Date & "'))"
    Adodc8.Refresh
    Adodc4.RecordSource = "select Ӫҵʱ��,Ъҵʱ�� from �������� where (���� ='" & xmMC(Index) & "')"
    Adodc4.Refresh
    If Adodc8.Recordset.RecordCount < Adodc7.Recordset.Fields("Ӫҵ����") And Time > Adodc4.Recordset.Fields("Ӫҵʱ��") And Time < Adodc4.Recordset.Fields("Ъҵʱ��") Then  'Ӫҵ���� Ӫҵʱ��
        qued.Image1(0).Picture = Image1(Index).Picture
        qued.Label3.Caption = Val(Label1(Index).Caption) '+ 1
        qued.Label5.Caption = xmid(Index)
        qued.Label6.Caption = xmMC(Index)
        Load qued
        qued.Show 1
    Else
        Load yyrenxian
        yyrenxian.Show
    End If
End Sub

Private Sub MSComm2_OnComm()

Dim Buffer() As Byte, fbuffer() As String, mstring, remote_control_vale As String, remote_control_code As String, fsong As String, item() As Byte, current_no As String, rest As String, return_hex As String
    Select Case MSComm2.CommEvent
        Case 2
            Buffer = MSComm2.Input '�����ַ���
            MSComm2.InBufferCount = 0    '��ջ���
            
            For i = 0 To UBound(Buffer)   '�����յ��ַ�ת��Ϊ��λ��ʮ��������
                If Len(Hex(Buffer(i))) = 1 Then
                  COM_STRING = COM_STRING & "0" & Hex(Buffer(i)) & " "
                Else
                  COM_STRING = COM_STRING & Hex(Buffer(i)) & " "
                End If
              
                If Hex(Buffer(i)) = "A" Then    '���յ��ַ�������A
                    MSComm2.PortOpen = False
'                    MsgBox "�رմ���"
                    '�б���ң��
                    If InStr(COM_STRING, "3A 2A 48 46 3D") > 0 And InStr(COM_STRING, "0A") > 0 And InStr(COM_STRING, "0A") - InStr(COM_STRING, "3A 2A 48 46 3D") = 42 Then  '����ң��������
'                        MsgBox "��ң��"
                        COM_STRING = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D"), InStr(COM_STRING, "0A") - InStr(COM_STRING, "3A 2A 48 46 3D") + 2)  '��ȡ�����ַ���
                         
                        remote_control_code = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D") + 15, 17) 'ң�ر���
                         
                        remote_control_code = Replace(remote_control_code, " ", "") 'ȥ���ַ����Ŀո�
                
                        remote_control_vale = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 3D") + 33, 5)  'ң�ؼ���
                        
                        remote_control_vale = Replace(remote_control_vale, " ", "")  'ȥ���ַ����Ŀո�
                        
                        
                        Adodc4.RecordSource = "SELECT * FROM ң������ WHERE  (�豸����=" & Val(remote_control_code) & ") and (�ؽ�=" & Val(remote_control_vale) & ") "  '�ؽ�
                        Adodc4.Refresh
                        If Adodc4.Recordset.RecordCount > 0 Then
                            Adodc7.RecordSource = "select top 1 ��� from �����к� where (�豸����=" & Val(remote_control_code) & ") and (״̬='3') order by id DESC" '�޸�״̬"
                            Adodc7.Refresh
                            If Adodc7.Recordset.RecordCount > 0 Then
                                Conn.Open sqlcnn
                                Conn.Execute "UPDATE �����к� SET ״̬ = '2' WHERE (�豸����=" & Val(remote_control_code) & ") and (���= " & Adodc7.Recordset.Fields("���") & ")"
                                Conn.Close
                            End If
                        End If
                        
                        
                        Adodc7.RecordSource = "SELECT * FROM ң������ WHERE  (�豸����=" & Val(remote_control_code) & ") and (��һλ=" & Val(remote_control_vale) & ") "  '��һλ
                        Adodc7.Refresh
                                If Adodc7.Recordset.RecordCount > 0 Then
                                    Adodc5.RecordSource = "select top 1 * from �Ŷ��б� where (״̬='0') and (��λid=" & Adodc7.Recordset.Fields("����") & ") order by id"
                                    Adodc5.Refresh
                                    Adodc1.RecordSource = "select * from ң������ where (�豸����=" & Val(remote_control_code) & ")"
                                    Adodc1.Refresh
                                    If Adodc5.Recordset.RecordCount > 0 Then
                                        Conn.Open sqlcnn
                                        Conn.Execute "update �Ŷ��б� set ״̬='2' where (id=" & Adodc5.Recordset.Fields("id") & ") and (״̬='0')"
                                        Conn.Close
                                        Adodc4.RecordSource = "select * from �����к�"
                                        Adodc4.Refresh
                                        Adodc4.Recordset.AddNew
                                        Adodc4.Recordset.Fields("����") = Adodc1.Recordset.Fields("id")
                                        Adodc4.Recordset.Fields("���") = Adodc5.Recordset.Fields("���")
                                        Adodc4.Recordset.Fields("״̬") = "2"
                                        Adodc4.Recordset.Fields("ǰ�к�") = Adodc1.Recordset.Fields("ǰ�к�")
                                        Adodc4.Recordset.Fields("��к�") = Adodc1.Recordset.Fields("��к�")
                                        Adodc4.Recordset.Fields("�豸����") = Val(remote_control_code)
                                        Adodc4.Recordset.Fields("����") = Date
                                        Adodc4.Recordset.Update
                                        
                                        Dim text() As String
                                        Open App.Path & "\yemianbuzhi.ini" For Input As #1
                                        Input #1, s
                                            text = Split(s, "VbCrVbLf")
                                        Close #1
                                        fsong = led_coad(text(12), Trim(Adodc5.Recordset.Fields("����")), Trim(Adodc5.Recordset.Fields("ǰ��")), Trim(Adodc5.Recordset.Fields("���")), Trim(Adodc5.Recordset.Fields("����")))
        
                                        Adodc4.RecordSource = "select * from led��ʾ"
                                        Adodc4.Refresh
                                        Adodc4.Recordset.AddNew
                                        Adodc4.Recordset.Fields("����") = Adodc1.Recordset.Fields("����")
                                        Adodc4.Recordset.Fields("���") = Adodc5.Recordset.Fields("���")
                                        Adodc4.Recordset.Fields("״̬") = "2"
                                        Adodc4.Recordset.Fields("ǰ��") = Adodc1.Recordset.Fields("ǰ��")
                                        Adodc4.Recordset.Fields("����") = Adodc1.Recordset.Fields("����")
                                        Adodc4.Recordset.Fields("����") = Adodc1.Recordset.Fields("����")
                                        Adodc4.Recordset.Fields("�豸����") = Val(remote_control_code)
                                        Adodc4.Recordset.Fields("����") = Date
                                        Adodc4.Recordset.Fields("����") = fsong
                                        Adodc4.Recordset.Update
                                        
'                                        led_no = Adodc5.Recordset.Fields("���")
'                                        Timer4.Enabled = True
'                                        timer_var = 0
                                        Dim return_string As String, strHexSj As String
                                        Adodc8.RecordSource = "select * from �Ŷ��б� where ״̬='0' and (��λid=" & Adodc7.Recordset.Fields("����") & ")"
                                        Adodc8.Refresh
'                                        MsgBox Right("00" & Adodc5.Recordset.Fields("���"), 4)
                                        
                                        item = StrConv(Trim(Adodc7.Recordset.Fields("����")), vbFromUnicode) '��Ŀ����תΪʮ��������
                                        For k = 0 To UBound(item)
                                            return_hex = return_hex & Right("0" & Hex(item(i)), 2)
                                        Next
                                        
                                        return_hex = Left(return_hex & "2020202020202020", 16)
                                        For k = 1 To 4 Step 2
                                            current_no = current_no & Right("00" & Hex(Mid(Right("00" & Adodc5.Recordset.Fields("���"), 4), k, 2)), 2) '��ǰ����
                                        Next
'                                        MsgBox current_no
'                                        MsgBox Adodc7.Recordset.Fields("id")
'                                        MsgBox Adodc8.Recordset.RecordCount
                                        rest = Right("00" & Hex(Adodc8.Recordset.RecordCount), 2) 'ʣ������
                                        
                                        
'                                       MsgBox rest
                                      
                                       
                                        
                                        retutn_string = "3A2A484640" & remote_control_code & return_hex & current_no & rest & "0D0A"
'                                        MsgBox retutn_string
                                    End If
                                End If
                            
                    End If
'                Else
                    
                    'Сң��
                If InStr(COM_STRING, "3A 2A 48 46 23") > 0 And InStr(COM_STRING, "0A") > 0 And InStr(COM_STRING, "0A") - InStr(COM_STRING, "3A 2A 48 46 23") = 42 Then  '����ң��������
'                    MsgBox "Сң��"
                    COM_STRING = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 23"), InStr(COM_STRING, "0A") - InStr(COM_STRING, "3A 2A 48 46 23") + 2)  '��ȡ�����ַ���
                     
                    remote_control_code = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 23") + 15, 17) 'ң�ر���
                     
                    remote_control_code = Replace(remote_control_code, " ", "") 'ȥ���ַ����Ŀո�
            
                    remote_control_vale = Mid(COM_STRING, InStr(COM_STRING, "3A 2A 48 46 23") + 33, 5)  'ң�ؼ���
                    
                    remote_control_vale = Replace(remote_control_vale, " ", "")  'ȥ���ַ����Ŀո�
                    
'                    MsgBox remote_control_vale
                    
                    If remote_control_vale = "3031" Then 'Or remote_control_vale = "3033" Or remote_control_vale = "3035" Or remote_control_vale = "3037" Or remote_control_vale = "3039" Or remote_control_vale = "3131"
'                        MsgBox "��һλ"
                        Adodc4.RecordSource = "select * from �������� where (������='" & remote_control_vale & "') "
                        Adodc4.Refresh
                        If Adodc4.Recordset.RecordCount > 0 Then
                            Adodc1.RecordSource = "select TOP 1 * from �Ŷ��б� where (��λid= '" & Adodc4.Recordset.Fields("id") & "') and ״̬='0' order by id"
                            Adodc1.Refresh
                        End If
'                        MsgBox Adodc1.Recordset.Fields("���")
                        '
'                        Adodc1.RecordSource = "select * from �Ŷ��б� where (״̬='0') and (������= '" & strsss & "') order by ���"
'                        Adodc1.Refresh
                        If Adodc1.Recordset.RecordCount > 0 Then
'                            Adodc1.Recordset.MoveFirst
                            Conn.Open sqlcnn
                            Conn.Execute "update �Ŷ��б� set ״̬='3' where (id='" & Adodc1.Recordset.Fields("id") & "') and (״̬='0')"
                            Conn.Close
                            Adodc4.RecordSource = "select * from �����к�"
                            Adodc4.Refresh
                            Adodc4.Recordset.AddNew
                            Adodc4.Recordset.Fields("����") = Adodc1.Recordset.Fields("��λid")
                            Adodc4.Recordset.Fields("���") = Adodc1.Recordset.Fields("���")
                            Adodc4.Recordset.Fields("״̬") = "2"
                            Adodc4.Recordset.Fields("ǰ�к�") = Adodc1.Recordset.Fields("ǰ�к��ļ�")
                            Adodc4.Recordset.Fields("��к�") = Adodc1.Recordset.Fields("��к��ļ�")
                            Adodc4.Recordset.Fields("����") = Date
                            Adodc4.Recordset.Fields("�豸����") = remote_control_vale + 1
                            Adodc4.Recordset.Update

                            Open App.Path & "\yemianbuzhi.ini" For Input As #1
                            Input #1, s
                            text = Split(s, "VbCrVbLf")
                            Close #1
                            fsong = led_coad(text(12), Trim(Adodc1.Recordset.Fields("����")), Trim(Adodc1.Recordset.Fields("ǰ��")), Trim(Adodc1.Recordset.Fields("���")), Trim(Adodc1.Recordset.Fields("����")))

                            Adodc5.RecordSource = "select * from led��ʾ"
                            Adodc5.Refresh
                            Adodc5.Recordset.AddNew
                            Adodc5.Recordset.Fields("����") = Adodc1.Recordset.Fields("��λid")
                            Adodc5.Recordset.Fields("���") = Adodc1.Recordset.Fields("���")
                            Adodc5.Recordset.Fields("״̬") = "2"
                            Adodc5.Recordset.Fields("ǰ��") = Adodc1.Recordset.Fields("ǰ��")
                            Adodc5.Recordset.Fields("����") = Adodc1.Recordset.Fields("����")
                            Adodc5.Recordset.Fields("����") = Adodc1.Recordset.Fields("����")
                            Adodc5.Recordset.Fields("����") = fsong
                            Adodc5.Recordset.Fields("�豸����") = remote_control_vale + 1
                            Adodc5.Recordset.Fields("����") = Date
                            Adodc5.Recordset.Update


                        End If
                        
                    End If
                    If remote_control_vale = "3032" Then 'Or remote_control_vale = "3034" Or remote_control_vale = "3036" Or remote_control_vale = "3038" Or remote_control_vale = "3130" Or remote_control_vale = "3132" Then
'                        MsgBox "�ؽ�"
                        Adodc7.RecordSource = "select top 1 * from �����к� where (�豸����='" & Val(remote_control_vale) & "') and (״̬='3') order by id desc"
                        Adodc7.Refresh
'                        MsgBox Adodc7.Recordset.RecordCount
'                        Adodc1.Recordset.MoveLast
                        If Adodc1.Recordset.RecordCount > 0 Then
'                        MsgBox Adodc1.Recordset.Fields("id")
                            Conn.Open sqlcnn
                            Conn.Execute "update �����к� set ״̬='2' where (id='" & Adodc7.Recordset.Fields("id") & "')"
                            Conn.Close
                        End If
                    End If

                End If
                    
                    MSComm2.PortOpen = True
'            MsgBox COM_STRING
            COM_STRING = ""
                            End If
                          
                           Exit For
                         
                
            Next
            
    End Select
End Sub
Private Sub MSComm1_OnComm()
'�к�
'On Error Resume Next
Dim Buffer() As Byte, strsss, text() As String, fsong As String
    Select Case MSComm1.CommEvent
        Case 2
            Buffer = MSComm1.Input
            MSComm1.InBufferCount = 0   '��ջ�����
        For i = 0 To UBound(Buffer)
            If Len(Hex(Buffer(i))) = 1 Then
                strsss = "0" & Hex(Buffer(i))
            Else
                strsss = Hex(Buffer(i))
            End If
            
            If strsss = "08" Or strsss = "04" Or strsss = "0C" Or strsss = "02" Or strsss = "0A" Or strsss = "06" Or strsss = "0E" Or strsss = "01" Or strsss = "09" Or strsss = "05" Or strsss = "0D" Or strsss = "03" Then  '20 04 0C 02 0A 06 0E 01 09 05 0D 03
            SQL = "UPDATE �������� SET ��ǰ���� = ��ǰ���� - 1 WHERE (������ = '" & strsss & "') AND (��ǰ���� > 0)"
            Conn.Open sqlcnn
            Conn.Execute SQL
            Conn.Close
            End If

            Adodc1.RecordSource = "select * from �Ŷ��б� where (״̬='0') and (������= '" & strsss & "') order by ���"
            Adodc1.Refresh
            If Adodc1.Recordset.RecordCount > 0 Then
                Adodc1.Recordset.MoveFirst
                Conn.Open sqlcnn
                Conn.Execute "update �Ŷ��б� set ״̬='2' where (���=" & Adodc1.Recordset.Fields("���") & ") and (������= '" & strsss & "') and (״̬='0')"
                Conn.Close
                Adodc4.RecordSource = "select * from �����к�"
                Adodc4.Refresh
                Adodc4.Recordset.AddNew
                Adodc4.Recordset.Fields("����") = Adodc1.Recordset.Fields("��λid")
                Adodc4.Recordset.Fields("���") = Adodc1.Recordset.Fields("���")
                Adodc4.Recordset.Fields("״̬") = "2"
                Adodc4.Recordset.Fields("ǰ�к�") = Adodc1.Recordset.Fields("ǰ�к��ļ�")
                Adodc4.Recordset.Fields("��к�") = Adodc1.Recordset.Fields("��к��ļ�")
                Adodc4.Recordset.Fields("����") = Date
                Adodc4.Recordset.Update

                Open App.Path & "\yemianbuzhi.ini" For Input As #1
                Input #1, s
                text = Split(s, "VbCrVbLf")
                Close #1
                fsong = led_coad(text(12), Trim(Adodc1.Recordset.Fields("����")), Trim(Adodc1.Recordset.Fields("ǰ��")), Trim(Adodc1.Recordset.Fields("���")), Trim(Adodc1.Recordset.Fields("����")))

                Adodc5.RecordSource = "select * from led��ʾ"
                Adodc5.Refresh
                Adodc5.Recordset.AddNew
                Adodc5.Recordset.Fields("����") = Adodc1.Recordset.Fields("��λid")
                Adodc5.Recordset.Fields("���") = Adodc1.Recordset.Fields("���")
                Adodc5.Recordset.Fields("״̬") = "2"
                Adodc5.Recordset.Fields("ǰ��") = Adodc1.Recordset.Fields("ǰ��")
                Adodc5.Recordset.Fields("����") = Adodc1.Recordset.Fields("����")
                Adodc5.Recordset.Fields("����") = Adodc1.Recordset.Fields("����")
                Adodc5.Recordset.Fields("����") = fsong
                Adodc5.Recordset.Fields("����") = Date
                Adodc5.Recordset.Update


            End If
        Next
    End Select
End Sub
Private Sub ggck()   '��洰��
       Select Case mURL
        Case 4
                If Dir(App.Path & "\avi\5.avi") <> "" Then
                  mURL = 0
                  WindowsMediaPlayer1.URL = App.Path & "\avi\5.avi"
                  WindowsMediaPlayer1.Controls.Play
                Else
                  mURL = 0
                  WindowsMediaPlayer1.URL = App.Path & "\avi\1.avi"
                  WindowsMediaPlayer1.Controls.Play
                End If
       Case 3
                If Dir(App.Path & "\avi\5.avi") <> "" Then
                  mURL = 4
                  WindowsMediaPlayer1.URL = App.Path & "\avi\5.avi"
                  WindowsMediaPlayer1.Controls.Play
                Else
                  mURL = 0
                  WindowsMediaPlayer1.URL = App.Path & "\avi\1.avi"
                  WindowsMediaPlayer1.Controls.Play
                End If
        Case 2
               If Dir(App.Path & "\avi\4.avi") <> "" Then
                   mURL = 3
                  WindowsMediaPlayer1.URL = App.Path & "\avi\4.avi"
                  WindowsMediaPlayer1.Controls.Play
                Else
                   mURL = 0
                  WindowsMediaPlayer1.URL = App.Path & "\avi\1.avi"
                  WindowsMediaPlayer1.Controls.Play
                End If
         Case 1
                If Dir(App.Path & "\avi\3.avi") <> "" Then
                   mURL = 2
                  WindowsMediaPlayer1.URL = App.Path & "\avi\3.avi"
                  WindowsMediaPlayer1.Controls.Play
                Else
                   mURL = 0
                  WindowsMediaPlayer1.URL = App.Path & "\avi\1.avi"
                  WindowsMediaPlayer1.Controls.Play
                End If
          Case 0
                If Dir(App.Path & "\avi\2.avi") <> "" Then
                   mURL = 1
                  WindowsMediaPlayer1.URL = App.Path & "\avi\2.avi"
                  WindowsMediaPlayer1.Controls.Play
                Else
                   mURL = 0
                  WindowsMediaPlayer1.URL = App.Path & "\avi\1.avi"
                  WindowsMediaPlayer1.Controls.Play
                End If
        End Select
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Timer1.Enabled = False
Dim pduibid As Integer, q As Long
Dim xingx(12) As String
Dim text() As String
If dianjics > 2 Then
    Load xtpeiz
    xtpeiz.Show 1
End If

dianjics = 0
    For i = 0 To 20
        If Val(xmid(i)) > 0 Then
               Adodc2.RecordSource = "SELECT TOP 1 * FROM �Ŷ��б� WHERE (��λid = " & xmid(i) & ") AND (���� = CONVERT(DATETIME, '" & Date & " 00:00:00', 102)) AND (״̬ < '2')  ORDER BY ���"  'and (���<=" & Val(Format(Time, "hhmm")) + 8 & ")
               Adodc2.Refresh
              If Adodc2.Recordset.RecordCount > 0 Then
              
                  Adodc3.RecordSource = "SELECT * FROM �������� WHERE (���� - ��ǰ���� > 0) and (id = " & Val(xmid(i)) & ")"
                  Adodc3.Refresh
                  If Adodc3.Recordset.RecordCount > 0 Then  '�п���
                     '�޸� ��λ��������
                     Adodc3.Recordset.Fields("��ǰ����") = Val(Adodc3.Recordset.Fields("��ǰ����")) + 1
                     Adodc3.Recordset.Update
                     
                     '�޸��Ŷӱ� �к�״̬
                        pduibid = Adodc2.Recordset.Fields("id")
                        For k = 0 To 12
                              xingx(k) = Adodc2.Recordset.Fields(k)
                        Next k
                           
                             SQL = "UPDATE �Ŷ��б� SET ״̬ = '2' WHERE (id = " & xingx(0) & ")"  'and (���<=" & Val(Format(Time, "hhmm")) & ")
                             Conn.Open sqlcnn
                             Conn.Execute SQL
                             Conn.Close
                             
                             '�������
                             WindowsMediaPlayer1.settings.Volume = 0
                             Open App.Path & "\media_volum.ini" For Output As #1
                             Write #1, "1" & "VbCrVbLf" & "60"
                             Close #1

  
'                           If Len(xingx(12)) = 3 Then xingx(12) = Right("0000" & xingx(12), 4)
'
                           '���� �����к�
                          
'                           PlayWavFile App.Path & "\����\" & Trim(xingx(7)), 1, 0
'
'                              For ccc = 1 To Len(xingx(12))
'
'                                 PlayWavFile App.Path & "\����\" & Mid(Trim(xingx(12)), ccc, 1) & ".wav", 1, 0
'
'                              Next ccc
'                              PlayWavFile App.Path & "\����\" & Trim(xingx(8)), 1, 0
'
                           
                              
                                '����led��ʾ��(С��Ļ)
        If Dir(App.Path & "\yemianbuzhi.ini") <> "" Then
        Else
            Open App.Path & "\yemianbuzhi.ini" For Output As #1
            Write #1, "4000" & "VbCrVbLf" & "2000" & "VbCrVbLf" & "4" & "VbCrVbLf" & "1" & "VbCrVbLf" & "0" & "VbCrVbLf" & "2000" & "VbCrVbLf" & "122" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "5000" & "VbCrVbLf" & "0" & "VbCrVbLf" & "0" & "VbCrVbLf" & "4"
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
                            '��led������Ϣ
                            
'MsgBox Trim(xingx(10))
'MsgBox Trim(xingx(12))
'MsgBox Trim(xingx(11))

                                led_out = Split(led_coad(Trim(text(12)), Trim(xingx(9)), Trim(xingx(10)), Trim(xingx(12)), Trim(xingx(11))), " ")
'                                Text1.text = led_coad(Trim(text(12)), Trim(xingx(9)), Trim(xingx(10)), Trim(xingx(12)), Trim(xingx(11)))
'                                Label4.Caption = Len(Text1.text)
'                                For t = 0 To UBound(led_out)
'                                Label4.Caption = Label4.Caption & led_out(t) & " "
'                                Next
   If UBound(led_out) < 33 Then



                              Dim led_out_hex() As Byte
                              ReDim led_out_hex(UBound(led_out) - 1)



                             For l = 0 To UBound(led_out) - 1


                               led_out_hex(l) = CLng("&H" & led_out(l))

                              Next


                             For l = 0 To UBound(led_out_hex) - 1


'                              Text1.text = Text1.text & " " & Hex(led_out_hex(l))

                              Next
                              MSComm2.Output = led_out_hex
                              Else
                              Timer3.Enabled = True
                              End If
                              
                              
                              
                              
                                



    

                            
                              
                           'SELECT TOP 5 * FROM �Ŷ��б� WHERE (״̬ = '0') AND (��λid = " & xmid(i) & ") ORDER BY id DESC          ���ŷ��ͷ���
                           
                          Adodc3.RecordSource = "SELECT TOP 5 * FROM �Ŷ��б� WHERE (״̬ < '2') AND (��λid = " & xmid(i) & ") AND (���� = CONVERT(DATETIME, '" & Date & " 00:00:00', 102)) ORDER BY id"
                          Adodc3.Refresh
                          If Adodc3.Recordset.RecordCount = 5 Then
                             If Len((Trim(Adodc3.Recordset.Fields("�绰����")))) = 11 And Trim(Adodc3.Recordset.Fields("״̬")) = "0" Then
                                 
                                     Adodc3.Recordset.MoveLast
                                
                                  
                                    Shell (App.Path & "\����.exe " & dusnax(2) & "VbCrVbLf" & Trim(Adodc3.Recordset.Fields("�绰����")) & "VbCrVbLf" & dusnax(1) & " �𾴵�" & Trim(Adodc3.Recordset.Fields("���")) & "�ſͻ�" & "VbCrVbLf" & dusnax(0))
                                    SQL = "UPDATE �Ŷ��б� SET ״̬ = '1' WHERE (id = " & Adodc3.Recordset.Fields("id") & ")"
                                    Conn.Open sqlcnn
                                    Conn.Execute SQL
                                    Conn.Close
                             
                             End If
                          End If
                     '
              
                      Exit For
                  End If
              
              
              
              
              End If
  End If
Next i
yemtp


  
Timer1.Enabled = True
End Sub
Private Sub Timer2_Timer()
On Error Resume Next
yemtp
If dianjics > 2 Then
    Load xtpeiz
    xtpeiz.Show 1
End If
dianjics = 0
If WindowsMediaPlayer1.PlayState = 1 Or WindowsMediaPlayer1.PlayState = 2 Then
    ggck
End If

Adodc6.ConnectionString = sqlcnn
Adodc6.RecordSource = "SELECT ״̬,��λid,���,ǰ��,����  FROM �Ŷ��б� where (���� = CONVERT(DATETIME, '" & Date & " 00:00:00', 102)) and (״̬<'2') order by ��� "
Adodc6.Refresh

If Dir(App.Path & "\media_volum.ini") <> "" Then    '�����ʼ�кŵ��͹������
    Open App.Path & "\media_volum.ini" For Input As #1
    Input #1, s
    media_volum1 = Split(s, "VbCrVbLf")
    Close #1
End If
If Val(media_volum1(0)) = 0 Then WindowsMediaPlayer1.settings.Volume = media_volum1(1)

End Sub

Private Sub Timer3_Timer()   '��led�������ַ���
If Len(led_string1) > 30 Then
    Timer5.Enabled = True
Else
    Timer5.Enabled = False
End If
'Dim led_out_hex(31) As Byte, led_out_string(31) As String, led_out_hex_short() As Byte, textstring As String, l As Integer
'    If send_led + 31 > UBound(led_out) Then
'        For k = send_led To UBound(led_out)
'            led_out_string(k - send_led) = led_out(k)
'        Next k
'
'
'
'
'        ReDim led_out_hex_short(UBound(led_out) - send_led - 1) As Byte
'        For l = 0 To UBound(led_out) - send_led - 1
''            textstring = textstring & CLng("&H" & led_out_string(l)) & " "
'            led_out_hex_short(l) = CLng("&H" & led_out_string(l))
'        Next l
'
''        MsgBox led_out_hex_short(UBound(led_out) - send_led - 1)
'
'        MSComm2.Output = led_out_hex_short
'        send_led = 0
'        Timer3.Enabled = False
'    Else
'        For k = 0 To 31
'            led_out_string(k) = led_out(k + send_led)
'        Next k
'        For l = 0 To 31
'            led_out_hex(l) = CLng("&H" & led_out_string(l))
'        Next l
'        MSComm2.Output = led_out_hex
'        send_led = send_led + 32
'
''        For i = 0 To 31
''            led_text = led_text & led_out_string(i) & " "
''        Next
''        MsgBox led_text & " " & UBound(led_out_string)
'        End If
End Sub


Private Sub Timer4_Timer()   'led���ַ�
    Adodc1.RecordSource = "select * from led��ʾ where (״̬='2') order by id"
    Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        led_flag = led_flag + 1
        led_string1 = Trim(Adodc1.Recordset.Fields("����"))
        
        If Len(led_string1) Mod 64 > 0 Then
            For n = Len(led_string1) + 1 To (Len(led_string1) \ 64) * 64 + 64
                led_string1 = led_string1 & "0"
            Next
        Else
            For n = Len(led_string1) + 1 To Len(led_string1) \ 64 * 64
                led_string1 = led_string1 & "0"
            Next
        End If
        If led_flag > 0 Then
            Conn.Open sqlcnn
            Conn.Execute "update led��ʾ set ״̬='3' where (id = " & Adodc1.Recordset.Fields("id") & ")"
            Conn.Close
            led_falg = 0
        End If
        Timer4 = False
        Timer5 = True
        
    End If

End Sub

Private Sub Timer5_Timer()
Dim led_string() As String, led_fsong As String, str3 As String, led_string_short() As String
 
'        MsgBox led_string1
'        MsgBox Len(led_string1)
' MsgBox led_string1
' MsgBox Len(led_string1)
    If Len(led_string1) > 64 Then
        str3 = Mid(led_string1, 1, 64)
        led_string1 = Mid(led_string1, 65)
        Timer3.Enabled = True
        Timer5.Enabled = False
    Else
        str3 = led_string1
        led_string1 = ""
        Timer5.Enabled = False
        Timer4.Enabled = True
    End If
'     MsgBox str3
' MsgBox Len(str3)
 
    For i = 1 To Len(Trim(str3)) Step 2
        led_fsong = led_fsong & Mid(Trim(str3), i, 2) & " "
    Next i
    led_string_short = Split(led_fsong, " ")
    
    For i = 0 To 31
        led_out_hex(i) = CLng("&H" & led_string_short(i))
    Next
    MSComm2.Output = led_out_hex
'Dim textstring As String
'   For k = 0 To UBound(led_out_hex) - 1
'       textstring = textstring & " " & Hex(led_out_hex(k))
'   Next
' MsgBox textstring

' MsgBox Len(led_string1)
''        MsgBox UBound(led_string_short)
'        Dim counter As Integer
'        counter = 0
'        Do While counter < UBound(led_string_short)
'            For l = 0 To 31
'                led_out_hex(l) = CLng("&H" & led_string_short(l + counter))
'            Next
'            MSComm2.Output = led_out_hex
'            counter = counter + 32
           
'        Loop

'    If timer_var = 2 Then
'        Timer5.Enabled = False
'        timer_var = 0
'    Else
'        MSComm2.Output = led_out_hex
'        timer_var = timer_var + 1
'    End If
End Sub
