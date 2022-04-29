Attribute VB_Name = "Global"
Public HandbookFile As String    '手册文件名
Public ServerFile As String    'gc服务器文件名
Public MongoFile As String    'mongodb服务器文件名
Public MitmDumpFile As String    'mitmdump.exe
Public MitmProxyFile As String    'mitmproxy.py
Public JREFile As String    'jre.exe | java.exe | jdk.exe
Public ServerFolder As String    '服务器所在文件夹

Public LogCount As String    '日志保留字符数

Public OrigProxyEnable As Integer '原始系统代理
Public OrigProxyServer As String '原始系统代理
Public ProxyEnabled As Boolean '系统代理是否开启
Public ServerStarted As Boolean '割草机是否开启

Public Handbook As New Dictionary '手册
Public HandbookLoaded As Boolean '手册是否加载

Public Const YH As String = """"
Public Const TXT_HEADER_1 As String = vbCrLf & "  所谓薙刀，乃是斩除芜杂之利器。" & vbCrLf & "  秉薙刀之人，意在守护恒常之道。" & vbCrLf & vbCrLf & "  当高踞雷云之上者俯视她所倾心的凡世，" & vbCrLf & "  所见无不浅薄的争端，闪灭的执欲泡影…" & vbCrLf & "  争夺源于无谓爱执与狂欲，乃恒世之敌。" & vbCrLf & "  搅扰不变恒世的杂草，将交由雷光殛灭。" & vbCrLf & vbCrLf & "  既然谁人都无法逆转现世之无常、绪绝的独乐，" & vbCrLf & "  那就将心中的常世净土带到她所珍爱的国度吧。" & vbCrLf & vbCrLf & "  -- 生草机 (Grasscutter UI) By YidaozhanYa"
Public Const RTF_HEADER As String = "{\rtf1\ansi\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Consolas;}}{\colortbl ;\red0\green176\blue80;}\pard\sl240\slmult1\f0\fs20\lang2052 "

Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Sub FirstInit()
    HandbookFile = GetIni("GCUI", "HandbooKFile", App.Path & "\Config.ini")
    If HandbookFile = "" Then
        Do While HandbookFile = ""
            HandbookFile = ChooseFile("选择 Handbook 文件（本程序自带）", "Handbook", "*.txt", frmMain)
        Loop
        WriteIni "GCUI", "HandbookFile", HandbookFile, App.Path & "\Config.ini"
    End If
    ServerFile = GetIni("GCUI", "ServerFile", App.Path & "\Config.ini")
    If ServerFile = "" Then
        Do While ServerFile = ""
            ServerFile = ChooseFile("选择 Grasscutter 服务端文件", "Grasscutter", "*.jar", frmMain)
        Loop
        WriteIni "GCUI", "ServerFile", ServerFile, App.Path & "\Config.ini"
    End If

    ServerFolder = GetIni("GCUI", "ServerFolder", App.Path & "\Config.ini")
    If ServerFolder = "" Then
        Do While ServerFolder = ""
            ServerFolder = ChooseDir("选择 Grasscutter 服务端所在文件夹", frmMain)
        Loop
        WriteIni "GCUI", "ServerFolder", ServerFolder, App.Path & "\Config.ini"
    End If

    MongoFile = GetIni("GCUI", "MongoFile", App.Path & "\Config.ini")
    If MongoFile = "" Then
        Do While MongoFile = ""
            MongoFile = ChooseFile("选择 MongoDB 可执行文件", "MongoDB", "mongod.exe", frmMain)
        Loop
        WriteIni "GCUI", "MongoFile", MongoFile, App.Path & "\Config.ini"
    End If

    MitmDumpFile = GetIni("GCUI", "MitmDumpFile", App.Path & "\Config.ini")
    If MitmDumpFile = "" Then
        Do While MitmDumpFile = ""
            MitmDumpFile = ChooseFile("选择 mitmdump.exe 文件", "MITMProxy", "mitmdump.exe", frmMain)
        Loop
        WriteIni "GCUI", "MitmDumpFile", MitmDumpFile, App.Path & "\Config.ini"
    End If

    MitmProxyFile = GetIni("GCUI", "MitmProxyFile", App.Path & "\Config.ini")
    If MitmProxyFile = "" Then
        Do While MitmProxyFile = ""
            MitmProxyFile = ChooseFile("选择 MITMProxy 脚本文件", "Python 脚本", "*.py", frmMain)
        Loop
        WriteIni "GCUI", "MitmProxyFile", MitmProxyFile, App.Path & "\Config.ini"
    End If

    JREFile = GetIni("GCUI", "JREFile", App.Path & "\Config.ini")
    If JREFile = "" Then
        Do While JREFile = ""
            JREFile = ChooseFile("选择 JRE", "可执行文件", "*.exe", frmMain)
        Loop
        WriteIni "GCUI", "JREFile", JREFile, App.Path & "\Config.ini"
    End If
    
    If GetIni("GCUI", "LogCount", App.Path & "\Config.ini") = "" Then
        LogCount = 10000
        WriteIni "GCUI", "LogCount", "10000", App.Path & "\Config.ini"
    Else
        LogCount = CInt(GetIni("GCUI", "LogCount", App.Path & "\Config.ini"))
    End If
    
End Sub


''---------------------------
Public Function GetIni(strSection As String, strKey As String, INIFileName As String)
    With New ClassINI
        .INIFileName = INIFileName
        GetIni = .GetIniKey(strSection, strKey)
    End With
End Function

Public Sub WriteIni(strSection As String, strKey As String, strNewValue As String, INIFileName As String)
    With New ClassINI
        .INIFileName = INIFileName
        .WriteIniKey strSection, strKey, strNewValue
    End With
End Sub

Public Function ChooseFile(ByVal frmTitle As String, ByVal fileDescription As String, ByVal fileFilter As String, ByVal onForm As Object) As String
'oleexp 选择文件
    On Error Resume Next
    Dim pChoose As New FileOpenDialog
    Dim psiResult As IShellItem
    Dim lpPath As Long, sPath As String
    Dim tFilt() As COMDLG_FILTERSPEC
    ReDim tFilt(0)
    tFilt(0).pszName = fileDescription
    tFilt(0).pszSpec = fileFilter
    With pChoose
        .SetFileTypes UBound(tFilt) + 1, VarPtr(tFilt(0))
        .SetTitle frmTitle
        .SetOptions FOS_FILEMUSTEXIST + FOS_DONTADDTORECENT
        .Show onForm.hWnd
        .GetResult psiResult
        If (psiResult Is Nothing) = False Then
            psiResult.GetDisplayName SIGDN_FILESYSPATH, lpPath
            If lpPath Then
                SysReAllocString VarPtr(sPath), lpPath
                CoTaskMemFree lpPath
            End If
        End If
    End With
    If BStrFromLPWStr(lpPath) <> "" Then ChooseFile = BStrFromLPWStr(lpPath)
End Function

Public Function BStrFromLPWStr(lpWStr As Long) As String
    SysReAllocString VarPtr(BStrFromLPWStr), lpWStr
End Function

Public Function ChooseDir(ByVal frmTitle As String, onForm As Object) As String
'oleexp 选择目录
    On Error Resume Next
    Dim pChooseDir As New FileOpenDialog
    Dim psiResult As IShellItem
    Dim lpPath As Long, sPath As String
    With pChooseDir
        .SetOptions FOS_PICKFOLDERS
        .SetTitle frmTitle
        .Show onForm.hWnd
        .GetResult psiResult
        If (psiResult Is Nothing) = False Then
            psiResult.GetDisplayName SIGDN_FILESYSPATH, lpPath
            If lpPath Then
                SysReAllocString VarPtr(sPath), lpPath
                CoTaskMemFree lpPath
            End If
        End If
    End With
    If BStrFromLPWStr(lpPath) <> "" Then ChooseDir = BStrFromLPWStr(lpPath)
End Function



Public Function ReadTextFile(sFilePath As String) As String
   On Error Resume Next
   Dim handle As Integer
   If LenB(Dir$(sFilePath)) > 0 Then
      handle = FreeFile
      Open sFilePath For Binary As #handle
      ReadTextFile = Space$(LOF(handle))
      Get #handle, , ReadTextFile
      Close #handle
   End If
End Function

