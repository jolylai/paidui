VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form jicxxi 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ϣ�༭"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6210
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   4560
      TabIndex        =   34
      Text            =   "Text14"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   4560
      MaxLength       =   8
      TabIndex        =   32
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   30
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ң������"
      Height          =   615
      Left            =   5160
      TabIndex        =   28
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   4560
      TabIndex        =   27
      Text            =   "Text11"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   2640
      TabIndex        =   24
      Text            =   "Text6"
      Top             =   3960
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7920
      Top             =   4680
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
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   615
      Left            =   4080
      TabIndex        =   23
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ɾ��"
      Height          =   495
      Left            =   3960
      TabIndex        =   20
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   1080
      TabIndex        =   19
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�Ƿ�����"
      Height          =   495
      Left            =   2400
      TabIndex        =   18
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Text            =   "Text9"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Text            =   "Text8"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Text            =   "Text7"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "�ؽ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   33
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ъҵʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   31
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ӫҵʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   29
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ӫҵ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   26
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ͼƬ�ļ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1080
      TabIndex        =   25
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   1800
      TabIndex        =   22
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "id��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   1200
      TabIndex        =   21
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   3720
      TabIndex        =   16
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ǰ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   840
      TabIndex        =   14
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   840
      TabIndex        =   13
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ���ã�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1200
      TabIndex        =   10
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�󱨺��ļ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   840
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ǰ�к��ļ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��һλ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "���ƣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "jicxxi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim text() As String
'text = Split(Text12.text, ":")
'If 0 <= text(0) < 24 And 0 <= text(1) < 60 And 0 <= text(2) < 60 Then
'Else
'    te

Adodc1.RecordSource = "SELECT id,Ӫҵʱ��,Ъҵʱ��, ����,Ӫҵ����, ����, ��ǰ����, ������, ǰ�к��ļ�, ��к��ļ�,ͼƬ, ����, ǰ��, ����,״̬ ,�ؽ� FROM �������� where id = " & Val(Label1(12).Caption)
Adodc1.Refresh
If Not Adodc1.Recordset.BOF And Not Adodc1.Recordset.EOF Then  '�޸�
   Adodc1.Recordset.Fields("����") = Trim(Text1.text)
   Adodc1.Recordset.Fields("����") = Val(Text2.text)
   Adodc1.Recordset.Fields("��ǰ����") = Val(Text3.text)
   Adodc1.Recordset.Fields("������") = Trim(Text4.text)
   Adodc1.Recordset.Fields("ǰ�к��ļ�") = Trim(Text5.text)
   Adodc1.Recordset.Fields("��к��ļ�") = Trim(Text6.text)
   Adodc1.Recordset.Fields("ͼƬ") = Trim(Text10.text)
   Adodc1.Recordset.Fields("����") = Trim(Text7.text)
   Adodc1.Recordset.Fields("ǰ��") = Trim(Text8.text)
   Adodc1.Recordset.Fields("����") = Trim(Text9.text)
   Adodc1.Recordset.Fields("״̬") = Trim(Check1.Value)
   Adodc1.Recordset.Fields("Ӫҵ����") = Trim(Text11.text)
   Adodc1.Recordset.Fields("Ӫҵʱ��") = Text12.text
   Adodc1.Recordset.Fields("Ъҵʱ��") = Text13.text
   Adodc1.Recordset.Fields("�ؽ�") = Text14.text
   Adodc1.Recordset.Update

   MsgBox "�������ݳɹ�", , App.Title
Else
   Adodc1.Recordset.AddNew                                '�½�
   Adodc1.Recordset.Fields("����") = Trim(Text1.text)
   Adodc1.Recordset.Fields("����") = Val(Text2.text)
   Adodc1.Recordset.Fields("��ǰ����") = Val(Text3.text)
   Adodc1.Recordset.Fields("������") = Trim(Text4.text)
   Adodc1.Recordset.Fields("ǰ�к��ļ�") = Trim(Text5.text)
   Adodc1.Recordset.Fields("��к��ļ�") = Trim(Text6.text)
   Adodc1.Recordset.Fields("ͼƬ") = Trim(Text10.text)
   Adodc1.Recordset.Fields("����") = Trim(Text7.text)
   Adodc1.Recordset.Fields("ǰ��") = Trim(Text8.text)
   Adodc1.Recordset.Fields("����") = Trim(Text9.text)
   Adodc1.Recordset.Fields("״̬") = Trim(Check1.Value)
   Adodc1.Recordset.Fields("Ӫҵ����") = Trim(Text11.text)
   Adodc1.Recordset.Fields("Ӫҵʱ��") = Text12.text
   Adodc1.Recordset.Fields("Ъҵʱ��") = Text13.text
   Adodc1.Recordset.Fields("�ؽ�") = Text14.text
   Adodc1.Recordset.Update

    MsgBox "�������ݳɹ�", , App.Title

End If
   
Unload Me


End Sub

Private Sub Command2_Click()
     msg = MsgBox("ȷ��ɾ��id=" & Val(Label1(12).Caption) & "����λ�����", vbQuestion + vbYesNo, App.Title)
     If msg = vbYes Then
        If Not Adodc1.Recordset.BOF And Not Adodc1.Recordset.EOF Then
            Adodc1.Recordset.Delete adAffectCurrent
            Adodc1.Recordset.Update
        End If
    End If

    Unload Me
  
End Sub

Private Sub Command3_Click()
    Adodc1.RecordSource = "select MAX(id) as maxid from ��������"
    Adodc1.Refresh
    Label1(12).Caption = Adodc1.Recordset.Fields("maxid").Value + 1

End Sub

Private Sub Command4_Click()
    Open App.Path & "\telecontrol.ini" For Output As #1
    Write #1, Label1(12).Caption
    Close #1
    
    Unload Me
    
    Load ykpz
    ykpz.Show 1
End Sub

Private Sub Form_Activate()
On Error Resume Next
Adodc1.RecordSource = "SELECT id,Ӫҵʱ��,Ъҵʱ��, ����, Ӫҵ����,����, ��ǰ����, ������, ǰ�к��ļ�, ��к��ļ�,ͼƬ, ����, ǰ��, ����,״̬,�ؽ� FROM �������� where id = " & Val(Label1(12).Caption)
Adodc1.Refresh
If Not Adodc1.Recordset.BOF And Not Adodc1.Recordset.EOF Then

  Text1.text = Trim(Adodc1.Recordset.Fields("����"))
  Text2.text = Val(Adodc1.Recordset.Fields("����"))
  Text3.text = Val(Adodc1.Recordset.Fields("��ǰ����"))
  Text4.text = Trim(Adodc1.Recordset.Fields("������"))
  Text5.text = Trim(Adodc1.Recordset.Fields("ǰ�к��ļ�"))
  Text6.text = Trim(Adodc1.Recordset.Fields("��к��ļ�"))
  Text10.text = Trim(Adodc1.Recordset.Fields("ͼƬ"))
  
  Text7.text = Trim(Adodc1.Recordset.Fields("����"))
  Text8.text = Trim(Adodc1.Recordset.Fields("ǰ��"))
  Text9.text = Trim(Adodc1.Recordset.Fields("����"))
  Check1.Value = Val(Adodc1.Recordset.Fields("״̬"))
  Text11.text = Trim(Adodc1.Recordset.Fields("Ӫҵ����"))
  Text12.text = Trim(Adodc1.Recordset.Fields("Ӫҵʱ��"))
  Text13.text = Trim(Adodc1.Recordset.Fields("Ъҵʱ��"))
  Text14.text = Trim(Adodc1.Recordset.Fields("�ؽ�"))
End If
Text1.MaxLength = 8
End Sub

Private Sub Form_Load()
   Adodc1.ConnectionString = sqlcnn
'   Me.Picture = LoadPicture(App.Path & "\pic\���ñ���.jpg")

End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    Text12.text = ""
End If
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    Text13.text = ""
End If
End Sub

Private Sub Text12_Change()
Dim text() As String
    If Len(Trim(Text12.text)) = 2 Then
        If 0 <= Val(Trim(Text12.text)) And Val(Trim(Text12.text)) < 24 Then
            Text12.text = Trim(Text12.text) & ":"
            Text12.SelStart = Len(Text12.text)
        Else
            Text12.text = ""
        End If
    End If
    If Len(Text12.text) = 5 Then
        text = Split(Trim(Text12.text), ":")
        If 0 <= Val(text(1)) And Val(text(1)) < 60 Then
            Text12.text = Trim(Text12.text) & ":00"
            Text12.SelStart = Len(Text12.text)
        Else
            Text12.text = text(0) & ":"
            Text12.SelStart = Len(Text12.text)
        End If
    End If
End Sub

Private Sub Text13_Change()
Dim text() As String
    If Len(Trim(Text13.text)) = 2 Then
        If 0 <= Val(Trim(Text13.text)) And Val(Trim(Text13.text)) < 24 Then
            Text13.text = Trim(Text13.text) & ":"
            Text13.SelStart = Len(Text13.text)
        Else
            If Trim(Text13.text) = 24 Then
                Text13.text = "23:59:59"
            Else
                Text13.text = ""
            End If
        End If
    End If
    If Len(Text13.text) = 5 Then
        text = Split(Trim(Text13.text), ":")
        If 0 <= Val(text(1)) And Val(text(1)) < 60 Then
            Text13.text = Trim(Text13.text) & ":00"
            Text13.SelStart = Len(Text13.text)
        Else
            Text13.text = text(0) & ":"
            Text13.SelStart = Len(Text13.text)
        End If
    End If
End Sub
