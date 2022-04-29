Attribute VB_Name = "Global"
Public HandbookFile As String    '�ֲ��ļ���
Public ServerFile As String    'gc�������ļ���
Public MongoFile As String    'mongodb�������ļ���
Public MitmDumpFile As String    'mitmdump.exe
Public MitmProxyFile As String    'mitmproxy.py
Public JREFile As String    'jre.exe | java.exe | jdk.exe
Public ServerFolder As String    '�����������ļ���

Public LogCount As String    '��־�����ַ���

Public OrigProxyEnable As Integer 'ԭʼϵͳ����
Public OrigProxyServer As String 'ԭʼϵͳ����
Public ProxyEnabled As Boolean 'ϵͳ�����Ƿ���
Public ServerStarted As Boolean '��ݻ��Ƿ���

Public Handbook As New Dictionary '�ֲ�
Public HandbookLoaded As Boolean '�ֲ��Ƿ����

Public Const YH As String = """"
Public Const TXT_HEADER_1 As String = vbCrLf & "  ��ν�S��������ն������֮������" & vbCrLf & "  ���S��֮�ˣ������ػ��㳣֮����" & vbCrLf & vbCrLf & "  ���߾�����֮���߸����������ĵķ�����" & vbCrLf & "  �����޲�ǳ�������ˣ������ִ����Ӱ��" & vbCrLf & "  ����Դ����ν��ִ��������˺���֮�С�" & vbCrLf & "  ���Ų���������Ӳݣ��������׹�����" & vbCrLf & vbCrLf & "  ��Ȼ˭�˶��޷���ת����֮�޳��������Ķ��֣�" & vbCrLf & "  �Ǿͽ����еĳ����������������䰮�Ĺ��Ȱɡ�" & vbCrLf & vbCrLf & "  -- ���ݻ� (Grasscutter UI) By YidaozhanYa"
Public Const RTF_HEADER As String = "{\rtf1\ansi\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Consolas;}}{\colortbl ;\red0\green176\blue80;}\pard\sl240\slmult1\f0\fs20\lang2052 "

Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Sub FirstInit()
    HandbookFile = GetIni("GCUI", "HandbooKFile", App.Path & "\Config.ini")
    If HandbookFile = "" Then
        Do While HandbookFile = ""
            HandbookFile = ChooseFile("ѡ�� Handbook �ļ����������Դ���", "Handbook", "*.txt", frmMain)
        Loop
        WriteIni "GCUI", "HandbookFile", HandbookFile, App.Path & "\Config.ini"
    End If
    ServerFile = GetIni("GCUI", "ServerFile", App.Path & "\Config.ini")
    If ServerFile = "" Then
        Do While ServerFile = ""
            ServerFile = ChooseFile("ѡ�� Grasscutter ������ļ�", "Grasscutter", "*.jar", frmMain)
        Loop
        WriteIni "GCUI", "ServerFile", ServerFile, App.Path & "\Config.ini"
    End If

    ServerFolder = GetIni("GCUI", "ServerFolder", App.Path & "\Config.ini")
    If ServerFolder = "" Then
        Do While ServerFolder = ""
            ServerFolder = ChooseDir("ѡ�� Grasscutter ����������ļ���", frmMain)
        Loop
        WriteIni "GCUI", "ServerFolder", ServerFolder, App.Path & "\Config.ini"
    End If

    MongoFile = GetIni("GCUI", "MongoFile", App.Path & "\Config.ini")
    If MongoFile = "" Then
        Do While MongoFile = ""
            MongoFile = ChooseFile("ѡ�� MongoDB ��ִ���ļ�", "MongoDB", "mongod.exe", frmMain)
        Loop
        WriteIni "GCUI", "MongoFile", MongoFile, App.Path & "\Config.ini"
    End If

    MitmDumpFile = GetIni("GCUI", "MitmDumpFile", App.Path & "\Config.ini")
    If MitmDumpFile = "" Then
        Do While MitmDumpFile = ""
            MitmDumpFile = ChooseFile("ѡ�� mitmdump.exe �ļ�", "MITMProxy", "mitmdump.exe", frmMain)
        Loop
        WriteIni "GCUI", "MitmDumpFile", MitmDumpFile, App.Path & "\Config.ini"
    End If

    MitmProxyFile = GetIni("GCUI", "MitmProxyFile", App.Path & "\Config.ini")
    If MitmProxyFile = "" Then
        Do While MitmProxyFile = ""
            MitmProxyFile = ChooseFile("ѡ�� MITMProxy �ű��ļ�", "Python �ű�", "*.py", frmMain)
        Loop
        WriteIni "GCUI", "MitmProxyFile", MitmProxyFile, App.Path & "\Config.ini"
    End If

    JREFile = GetIni("GCUI", "JREFile", App.Path & "\Config.ini")
    If JREFile = "" Then
        Do While JREFile = ""
            JREFile = ChooseFile("ѡ�� JRE", "��ִ���ļ�", "*.exe", frmMain)
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
'oleexp ѡ���ļ�
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
'oleexp ѡ��Ŀ¼
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

