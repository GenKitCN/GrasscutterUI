VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------classIniFile.cls 代码----------------

'这里定义了一个classIniFile类

'一个绝对经典的在VB中操作.ini文件的通用类源代码

'程序编写：中国青岛·许家国

' 2002.6.16

'E－Mail: goj2000@163.com

'HomePage: http://www.gojclub.com

'

'Private member that holds a reference to

'the path of our ini file

Private strINI As String    '初始化文件

'Windows API Declares

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                                                   (ByVal lpApplicationName As String, _
                                                    ByVal lpKeyName As Any, _
                                                    ByVal lpString As Any, _
                                                    ByVal lpFileName As String) As Long


Private Declare Function GetPrivateProfileString _
                          Lib "kernel32" Alias "GetPrivateProfileStringA" _
                              (ByVal lpApplicationName As String, _
                               ByVal lpKeyName As Any, _
                               ByVal lpDefault As String, _
                               ByVal lpReturnedString As String, _
                               ByVal nSize As Long, _
                               ByVal lpFileName As String) As Long


Private Function MakePath(ByVal strDrv As String, ByVal strDir As String) As String

' Makes an INI file: Guarantees a sub dir

    Do While Right$(strDrv, 1) = "\"
        strDrv = Left$(strDrv, Len(strDrv) - 1)
    Loop

    Do While Left$(strDir, 1) = "\"
        strDir = Mid$(strDir, 2)
    Loop

    ' Return the path

    MakePath = strDrv & "\" & strDir

End Function


Private Sub CreateIni(strDrv As String, strDir As String)

' Make a new ini file

    strINI = MakePath(strDrv, strDir)

End Sub


'--------------------------
'功能：写入ini文件
'--------------------------
Public Sub WriteIniKey(strSection As String, strKey As String, strValue As String)

' Write to strINI

    WritePrivateProfileString strSection, strKey, strValue, strINI

End Sub


'---------------------------
'功能：从ini文件中获取信息
'---------------------------
Public Function GetIniKey(strSection As String, strKey As String) As String

    On Error GoTo errhandle
    Dim strTmp As String
    Dim lngRet As String
    Dim i As Integer
    Dim strTmp2 As String

    '先将strtmp定义成1024个字符的长度，保证一定能装下返回的字串
    strTmp = String$(1024, Chr(32))

    'lpApplicationName String，欲在其中查找条目的小节名称。这个字串不区分大小写。如设为vbNullString，就在lpReturnedString缓冲区内装载这个ini文件所有小节的列表
    'lpKeyName String，欲获取的项名或条目名。这个字串不区分大小写。如设为vbNullString，就在lpReturnedString缓冲区内装载指定小节所有项的列表
    'lpDefault String，指定的条目没有找到时返回的默认值。可设为空（""）
    'lpReturnedString String，指定一个字串缓冲区，长度至少为nSize,返回的字符串
    'nSize Long，指定装载到lpReturnedString缓冲区的最大字符数量,返回字符串的长度
    'lpFileName String，初始化文件的名字。如没有指定一个完整路径名，windows就在Windows目录中查找文件


    lngRet = GetPrivateProfileString(strSection, strKey, "", strTmp, Len(strTmp), strINI)

    'strtmp现在已经是返回的字串了，所以要进行截尾处理
    strTmp = Trim(strTmp)
    strTmp2 = ""

    'ascii码为0对应的为空字符
    For i = 1 To Len(strTmp)
        If Asc(Mid(strTmp, i, 1)) <> 0 Then
            strTmp2 = strTmp2 + Mid(strTmp, i, 1)
        End If
    Next i

    GetIniKey = strTmp2
    Exit Function
errhandle:
    GetIniKey = ""

End Function


Public Property Let INIFileName(ByVal New_IniPath As String)

' Sets the new ini path

    strINI = New_IniPath

End Property

Public Property Get INIFileName() As String

' Returns the current ini path

    INIFileName = strINI

End Property

'***************************************清除KeyWord"键"(Sub)***********************************************

Public Function DelIniKey(ByVal SectionName As String, ByVal KeyWord As String)

    Dim RetVal As Integer

    RetVal = WritePrivateProfileString(SectionName, KeyWord, 0&, strINI)

End Function

'如果是清除section就少写一个Key多一个""。

'**************************************清除 Section"段"(Sub)***********************************************

Public Function DelIniSec(ByVal SectionName As String)    '清除section

    Dim RetVal As Integer

    RetVal = WritePrivateProfileString(SectionName, 0&, "", strINI)

End Function

