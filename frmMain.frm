VERSION 5.00
Object = "{A2A736C2-8DAC-4CDB-B1CB-3B077FBB14F9}#6.2#0"; "VB6Resizer2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Éú²Ý»ú"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   9
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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton cmdDrop 
      Caption         =   "µôÂäÎïÆ·"
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "µ±Ç°ÔÚÏß"
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtConsole 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   16
      Text            =   "Ö´ÐÐÃüÁî..."
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdChar 
      Caption         =   "¸øÓè½ÇÉ«"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdStats 
      Caption         =   "ÐÞ¸ÄÊýÖµ"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdKillAll 
      Caption         =   "Çå³ýµÐ¹Ö"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdScene 
      Caption         =   "³¡¾°´«ËÍ"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdAccount 
      Caption         =   "ÕËºÅ²Ù×÷"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "ÉèÖÃ"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   3360
      Width           =   855
   End
   Begin RichTextLib.RichTextBox txtLog 
      Height          =   5415
      Left            =   2040
      TabIndex        =   9
      ToolTipText     =   "ÓÒ¼üµã»÷ÇÐ»»Îª MITMProxy ÈÕÖ¾"
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9551
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":514A
   End
   Begin VB.TextBox txtUID 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   4290
      Width           =   1335
   End
   Begin VB.CommandButton cmdGive 
      Caption         =   "¸øÓèÎïÆ·"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdSpawn 
      Caption         =   "Éú³ÉµÐ¹Ö"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
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
      TabIndex        =   2
      Tag             =   "HW"
      ToolTipText     =   "ÓÒ¼üµã»÷ÇÐ»»Îª Grasscutter ÈÕÖ¾"
      Top             =   120
      Width           =   6495
   End
   Begin VB.CheckBox chkProxy 
      BackColor       =   &H80000005&
      Caption         =   "ÆôÓÃ´úÀí"
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Tag             =   "T"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Æô¶¯·þÎñÆ÷"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "UID:"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Grasscutter UI"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   500
      Width           =   1815
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
      Caption         =   "Éú²Ý»ú"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
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
        '±¸·Ý×¢²á±íÖÐÏµÍ³´úÀíÏà¹ØÉèÖÃ
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", "127.0.0.1:8080"
        Shell "certutil -addstore root " & Environ("UserProfile") & "\.mitmproxy\mitmproxy-ca-cert.cer", vbHide    '°²×°Ö¤Êé
        ProxyEnabled = True
        MsgBox "ÏµÍ³´úÀí¿ªÆô³É¹¦£¡"
    Else
        MITMDump.ClosePipe
        Shell "cmd /c taskkill /f /im mitmdump.exe"
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", OrigProxyEnable
        wsh.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", OrigProxyServer
        '»Ø¸´×¢²á±í
        Shell YH & App.Path & "\DisableProxy.bat" & YH, vbHide
        ProxyEnabled = False
        MsgBox "ÏµÍ³´úÀí¹Ø±Õ³É¹¦£¡"
    End If
End Sub

Private Sub cmdChar_Click()
    If ServerStarted Then frmOperate.TriggerChar
End Sub

Private Sub cmdDrop_Click()
    frmOperate.TriggerDrop
End Sub

Private Sub cmdList_Click()
    If ServerStarted Then Server.SendLine "!list"
    MsgBox "Çë¿´ÈÕÖ¾¡£"
End Sub

Private Sub cmdStats_Click()
    frmOperate.TriggerStats
End Sub

Private Sub cmdGive_Click()
    frmOperate.TriggerGive
End Sub

Private Sub cmdAccount_Click()
    If ServerStarted Then frmOperate.TriggerAccount
End Sub

Private Sub cmdKillAll_Click()
    If ServerStarted Then Server.SendLine "!killall " & txtUID.Text
    MsgBox "ÒÑ¾­Çå³ý " & txtUID.Text & " ËùÔÚ³¡¾°µÄÈ«²¿µÐ¹Ö¡£"
End Sub

Private Sub cmdScene_Click()
    frmOperate.TriggerScene
End Sub

Private Sub cmdSettings_Click()
frmSettings.Show
End Sub

Private Sub cmdSpawn_Click()
    frmOperate.TriggerEnemy
End Sub

Private Sub cmdStart_Click()
    If ServerStarted = False Then
        cmdStart.Caption = "Í£Ö¹·þÎñÆ÷"
        txtLog.TextRTF = ""
        txtMITMProxy.Text = ""
        '¿ªmongodb
        If Dir(ServerFolder & "\Database", vbDirectory) = "" Then MkDir ServerFolder & "\Database"
        Shell YH & MongoFile & YH & " --dbpath " & YH & ServerFolder & "\Database" & YH
        'ServerÎª´´½¨µÄ½ø³Ìpipe
        Dim SPResult As SP_RESULTS
        SPResult = Server.Run(YH & JREFile & YH & " -jar " & YH & ServerFile & YH, ServerFolder)
        Select Case SPResult
            'Case SP_SUCCESS
        Case SP_CREATEPIPEFAILED
            Shell "cmd /c taskkill /f /im java.exe"
            Shell "cmd /c taskkill /f /im mongod.exe"
            MsgBox "·þÎñÆ÷Æô¶¯Ê§°Ü£¬ÎÞ·¨´´½¨¹ÜµÀ¡£", vbOKOnly Or vbExclamation

        Case SP_CREATEPROCFAILED
            Shell "cmd /c taskkill /f /im java.exe"
            Shell "cmd /c taskkill /f /im mongod.exe"
            MsgBox "·þÎñÆ÷Æô¶¯Ê§°Ü£¬ÎÞ·¨´´½¨½ø³Ì¡£", vbOKOnly Or vbExclamation
        End Select
        ServerStarted = True
    Else
        cmdStart.Caption = "Æô¶¯·þÎñÆ÷"
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
        '»Ø¸´×¢²á±í
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
        '»Ø¸´×¢²á±í
        Shell YH & App.Path & "\DisableProxy.bat" & YH, vbHide
        ProxyEnabled = False
    End If
    End
End Sub

'´òÓ¡ÈÕÖ¾
Private Sub Server_DataArrival(ByVal CharsTotal As Long)
Dim txt As String, txtOrig As String
    With Server
        Do While .HasLine
            txt = Replace(Replace(.GetLine(), "[[34m", "\cf1 "), "[0;39m]", "\cf0 ")
            txtOrig = txtLog.TextRTF
            If Len(txtOrig) > (LogCount + 2) Then
            txtOrig = Right(txtOrig, LogCount)
            End If
            txtLog.TextRTF = RTF_HEADER & txtOrig & txt & "\par" & vbCrLf & "}"
            If txtLog.Visible Then txtLog.SelStart = &HFFFF&
        Loop
    End With
End Sub

Private Sub MITMDump_DataArrival(ByVal CharsTotal As Long)
    With MITMDump
        Do While .HasLine
            txtMITMProxy.Text = Right(txtMITMProxy.Text, LogCount) & .GetLine() & vbCrLf
            If txtMITMProxy.Visible Then txtMITMProxy.SelStart = &HFFFF&
        Loop
    End With
End Sub

Private Sub Form_Initialize()
    InitCommonControls
    Me.Caption = "Éú²Ý»ú " & App.Major & "." & App.Minor & "." & App.Revision
    txtLog.Visible = True
    txtMITMProxy.Visible = False
    ServerStarted = False
    txtLog.Text = TXT_HEADER_1
    txtUID.Text = GetIni("GCUI", "PlayerUID", App.Path & "\Config.ini")
    If txtUID.Text = "" Then txtUID.Text = "10001"
    FirstInit
    LoadHandbook
End Sub


Private Sub txtConsole_Click()
    If txtConsole.Text = "Ö´ÐÐÃüÁî..." Then txtConsole.Text = ""
End Sub

Private Sub txtConsole_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Server.SendLine txtConsole.Text
        txtConsole.Text = ""
    End If
End Sub

Private Sub txtLog_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        txtMITMProxy.Visible = True
        txtLog.Visible = False
    End If
End Sub

Private Sub txtMITMProxy_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        txtMITMProxy.Visible = False
        txtLog.Visible = True
    End If
End Sub

'ÊÖ²á
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
    Debug.Print "×Öµä¼ÓÔØÍê±Ï"
    HandbookLoaded = True
End Sub

Private Sub txtUID_Change()
        WriteIni "GCUI", "PlayerUID", txtUID.Text, App.Path & "\Config.ini"
End Sub
