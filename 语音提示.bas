Attribute VB_Name = "yuyj"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Sub PlayWavFile(strFileName As String, PlayCount As Long, JianGe As Long)
'strFileName Ҫ���ŵ��ļ���(��·��)
'playCount ���ŵĴ���
'JianGe  ��β���ʱ,ÿ�ε�ʱ����

If Len(Dir(strFileName)) = 0 Then Exit Sub

If PlayCount = 0 Then Exit Sub

If JianGe < 1000 Then JianGe = 1000

DoEvents
sndPlaySound strFileName, 16 + 1

Sleep JianGe

Call PlayWavFile(strFileName, PlayCount - 1, JianGe)

End Sub

