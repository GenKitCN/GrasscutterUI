VERSION 5.00
Object = "{A2A736C2-8DAC-4CDB-B1CB-3B077FBB14F9}#6.2#0"; "VB6Resizer2.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Grasscutter UI"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8655
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdGive 
      Caption         =   "������Ʒ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdSpawn 
      Caption         =   "���ɵй�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtMITMProxy 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "HW"
      ToolTipText     =   "�Ҽ�����л�Ϊ Grasscutter ��־"
      Top             =   120
      Width           =   6495
   End
   Begin VB.CheckBox chkProxy 
      BackColor       =   &H80000005&
      Caption         =   "���ô���"
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Tag             =   "T"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Tag             =   "HW"
      Text            =   "frmMain.frx":514A
      ToolTipText     =   "�Ҽ�����л�Ϊ MITMProxy ��־"
      Top             =   120
      Width           =   6495
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "����������"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Tag             =   "T"
      Top             =   5040
      Width           =   1815
   End
   Begin VB6ResizerLib2.VB6Resizer VB6Resizer1 
      Left            =   8040
      Top             =   5160
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin GrasscutterUI.ShellPipe ShellIConv 
      Left            =   600
      Top             =   1800
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Grasscutter"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin GrasscutterUI.ShellPipe MITMDump 
      Left            =   1560
      Top             =   4320
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin GrasscutterUI.ShellPipe Server 
      Left            =   7560
      Top             =   5160
      _ExtentX        =   635
      _ExtentY        =   635
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkProxy_Click()
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    If chkProxy.Value = 1 Then
        Call MITMDump.Run(YH & MitmDumpFile & YH & " -s " & YH & MitmProxyFile & YH & " -k")
        OrigProxyEnable = wsh.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")
        OrigProxyServer = wsh.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer")
        '����ע�����ϵͳ�����������
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", "127.0.0.1:8080"
        Shell "certutil -addstore root " & Environ("UserProfile") & "\.mitmproxy\mitmproxy-ca-cert.cer", vbHide    '��װ֤��
        ProxyEnabled = True
        MsgBox "ϵͳ�������ɹ���"
    Else
        MITMDump.ClosePipe
        Shell "cmd /c taskkill /f /im mitmdump.exe"
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", OrigProxyEnable
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", OrigProxyServer
        '�ظ�ע���
        Shell YH & App.Path & "\DisableProxy.bat" & YH, vbHide
        ProxyEnabled = False
        MsgBox "ϵͳ����رճɹ���"
    End If
End Sub

Private Sub cmdSpawn_Click()
    frmOperate.TriggerEnemy
End Sub

Private Sub cmdStart_Click()
    If ServerStarted = False Then
        cmdStart.Caption = "ֹͣ������"
        txtLog.Text = ""
        txtMITMProxy.Text = ""
        '��mongodb
        If Dir(ServerFolder & "\Database", vbDirectory) = "" Then MkDir ServerFolder & "\Database"
        Shell YH & MongoFile & YH & " --dbpath " & YH & ServerFolder & "\Database" & YH
        'ServerΪ�����Ľ���pipe
        Dim SPResult As SP_RESULTS
        SPResult = Server.Run(YH & JREFile & YH & " -jar " & YH & ServerFile & YH, ServerFolder)
        Select Case SPResult
            'Case SP_SUCCESS
        Case SP_CREATEPIPEFAILED
            Shell "cmd /c taskkill /f /im java.exe"
            Shell "cmd /c taskkill /f /im mongod.exe"
            MsgBox "����������ʧ�ܣ��޷������ܵ���", vbOKOnly Or vbExclamation

        Case SP_CREATEPROCFAILED
            Shell "cmd /c taskkill /f /im java.exe"
            Shell "cmd /c taskkill /f /im mongod.exe"
            MsgBox "����������ʧ�ܣ��޷��������̡�", vbOKOnly Or vbExclamation
        End Select
        ServerStarted = True
    Else
        cmdStart.Caption = "����������"
        Server.ClosePipe
        Shell "cmd /c taskkill /f /im java.exe"
        Shell "cmd /c taskkill /f /im mongod.exe"
        ServerStarted = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Server.ClosePipe
    Shell "cmd /c taskkill /f /im java.exe"
    Shell "cmd /c taskkill /f /im mongod.exe"
    If ProxyEnabled Then
        Shell "cmd /c taskkill /f /im mitmdump.exe"
        MITMDump.ClosePipe
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", OrigProxyEnable
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", OrigProxyServer
        '�ظ�ע���
        Shell YH & App.Path & "\DisableProxy.bat" & YH, vbHide
        ProxyEnabled = False
    End If
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Server.ClosePipe
    Shell "cmd /c taskkill /f /im java.exe"
    Shell "cmd /c taskkill /f /im mongod.exe"
    If ProxyEnabled Then
        Shell "cmd /c taskkill /f /im mitmdump.exe"
        MITMDump.ClosePipe
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", OrigProxyEnable
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", OrigProxyServer
        '�ظ�ע���
        Shell YH & App.Path & "\DisableProxy.bat" & YH, vbHide
        ProxyEnabled = False
    End If
    End
End Sub

Private Sub Server_DataArrival(ByVal CharsTotal As Long)
    With Server
        Do While .HasLine
            txtLog.Text = txtLog.Text & .GetLine() & vbCrLf
            If txtLog.Visible Then txtLog.SelStart = &HFFFF&
        Loop
    End With
End Sub

Private Sub MITMDump_DataArrival(ByVal CharsTotal As Long)
    With MITMDump
        Do While .HasLine
            txtMITMProxy.Text = txtMITMProxy.Text & .GetLine() & vbCrLf
            If txtMITMProxy.Visible Then txtMITMProxy.SelStart = &HFFFF&
        Loop
    End With
End Sub

Private Sub Form_Initialize()
    txtLog.Visible = True
    txtMITMProxy.Visible = False
    ServerStarted = False
    txtLog.Text = "��ӭ���� Grasscutter������������������"
    FirstInit
    LoadHandbook
End Sub

Private Sub txtLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        txtMITMProxy.Visible = True
        txtLog.Visible = False
    End If
End Sub

Private Sub txtMITMProxy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        txtMITMProxy.Visible = False
        txtLog.Visible = True
    End If
End Sub

'�ֲ�
Public Sub LoadHandbook()
        ShellIConv.Run YH & App.Path & "\iconv.exe" & YH & " -c -f utf-8 -t gbk " & YH & HandbookFile & YH
End Sub

Private Sub ShellIConv_ChildFinished()
On Error Resume Next
    Dim HandbookStr As String, sTemp As String, sTemp2 As String
    HandbookStr = ShellIConv.GetData()
    Dim EachItem As Variant, EachItem2 As Variant, InStrTemp As Variant
    For Each EachItem In Split(HandbookStr, "=== ")
    sTemp = Split(EachItem, vbLf)(0)
    If Right(sTemp, 3) = "===" Then
        'MsgBox EachItem
        sTemp2 = Replace(sTemp, " ===", "")
        Handbook.Add sTemp2, New Dictionary
        For Each EachItem2 In Split(EachItem, vbLf)
            InStrTemp = InStr(EachItem2, ": ")
            If InStrTemp <> False Then
                 Handbook.Item(sTemp2).Add Left(EachItem2, InStrTemp - 1), Right(EachItem2, Len(EachItem2) - InStrTemp - 1)
            End If
        Next
    End If
    Next
    Debug.Print "�ֵ�������"
    HandbookLoaded = True
End Sub
