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

'--------classIniFile.cls ����----------------

'���ﶨ����һ��classIniFile��

'һ�����Ծ������VB�в���.ini�ļ���ͨ����Դ����

'�����д���й��ൺ�����ҹ�

' 2002.6.16

'E��Mail: goj2000@163.com

'HomePage: http://www.gojclub.com

'

'Private member that holds a reference to

'the path of our ini file

Private strINI As String    '��ʼ���ļ�

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
'���ܣ�д��ini�ļ�
'--------------------------
Public Sub WriteIniKey(strSection As String, strKey As String, strValue As String)

' Write to strINI

    WritePrivateProfileString strSection, strKey, strValue, strINI

End Sub


'---------------------------
'���ܣ���ini�ļ��л�ȡ��Ϣ
'---------------------------
Public Function GetIniKey(strSection As String, strKey As String) As String

    On Error GoTo errhandle
    Dim strTmp As String
    Dim lngRet As String
    Dim i As Integer
    Dim strTmp2 As String

    '�Ƚ�strtmp�����1024���ַ��ĳ��ȣ���֤һ����װ�·��ص��ִ�
    strTmp = String$(1024, Chr(32))

    'lpApplicationName String���������в�����Ŀ��С�����ơ�����ִ������ִ�Сд������ΪvbNullString������lpReturnedString��������װ�����ini�ļ�����С�ڵ��б�
    'lpKeyName String������ȡ����������Ŀ��������ִ������ִ�Сд������ΪvbNullString������lpReturnedString��������װ��ָ��С����������б�
    'lpDefault String��ָ������Ŀû���ҵ�ʱ���ص�Ĭ��ֵ������Ϊ�գ�""��
    'lpReturnedString String��ָ��һ���ִ�����������������ΪnSize,���ص��ַ���
    'nSize Long��ָ��װ�ص�lpReturnedString������������ַ�����,�����ַ����ĳ���
    'lpFileName String����ʼ���ļ������֡���û��ָ��һ������·������windows����WindowsĿ¼�в����ļ�


    lngRet = GetPrivateProfileString(strSection, strKey, "", strTmp, Len(strTmp), strINI)

    'strtmp�����Ѿ��Ƿ��ص��ִ��ˣ�����Ҫ���н�β����
    strTmp = Trim(strTmp)
    strTmp2 = ""

    'ascii��Ϊ0��Ӧ��Ϊ���ַ�
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

'***************************************���KeyWord"��"(Sub)***********************************************

Public Function DelIniKey(ByVal SectionName As String, ByVal KeyWord As String)

    Dim RetVal As Integer

    RetVal = WritePrivateProfileString(SectionName, KeyWord, 0&, strINI)

End Function

'��������section����дһ��Key��һ��""��

'**************************************��� Section"��"(Sub)***********************************************

Public Function DelIniSec(ByVal SectionName As String)    '���section

    Dim RetVal As Integer

    RetVal = WritePrivateProfileString(SectionName, 0&, "", strINI)

End Function
