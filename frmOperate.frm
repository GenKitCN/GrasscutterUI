VERSION 5.00
Object = "{A2A736C2-8DAC-4CDB-B1CB-3B077FBB14F9}#6.2#0"; "VB6Resizer2.ocx"
Begin VB.Form frmOperate 
   BackColor       =   &H80000005&
   Caption         =   "Grasscutter UI"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOperate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   7920
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   6600
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7440
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtSearch 
      Height          =   420
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "����"
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Tag             =   "TL"
      Top             =   3480
      Width           =   1575
   End
   Begin VB6ResizerLib2.VB6Resizer VB6Resizer1 
      Left            =   6240
      Top             =   3480
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   120
      TabIndex        =   0
      Tag             =   "HW"
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   2680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   5760
      TabIndex        =   11
      Top             =   2190
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   5760
      TabIndex        =   9
      Top             =   1710
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   1240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   780
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmOperate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OperateType As String
Public HandbookKey As String

Public Sub TriggerEnemy()
    If HandbookLoaded <> True Then MsgBox "�ֲ����ڼ����У����Եȼ����ӡ�", vbCritical: Exit Sub
    Me.Caption = "���ݻ� " & App.Major & "." & App.Minor & "." & App.Revision & " - ���ɵй�"
    cmdAction.Caption = "����"
    OperateType = "spawn"
    If Handbook.Exists("Monsters") Then
        HandbookKey = "Monsters"
    ElseIf Handbook.Exists("�ַ��б�") Then
        HandbookKey = "�ַ��б�"
    Else
        HandbookKey = InputBox("�ֲ᲻֧�֣����ֶ�����й�����")
    End If
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
    Label1.Visible = True
    Label1.Caption = "[�ȼ�]:"
    Text1.Visible = True
    Label2.Visible = True
    Label2.Caption = "(�ȼ�Ϊѡ��)"
    Me.Show
End Sub
Public Sub TriggerChar()
    If HandbookLoaded <> True Then MsgBox "�ֲ����ڼ����У����Եȼ����ӡ�", vbCritical: Exit Sub
    Me.Caption = "���ݻ� " & App.Major & "." & App.Minor & "." & App.Revision & " - �����ɫ"
    cmdAction.Caption = "����"
    OperateType = "givechar"
    If Handbook.Exists("Characters") Then
        HandbookKey = "Characters"
    ElseIf Handbook.Exists("��ɫ") Then
        HandbookKey = "��ɫ"
    Else
        HandbookKey = InputBox("�ֲ᲻֧�֣����ֶ������ɫ����")
    End If
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
    Label1.Visible = True
    Label1.Caption = "[Amount]:"
    Text1.Visible = True
    Me.Show
End Sub
Public Sub TriggerScene()
    If HandbookLoaded <> True Then MsgBox "�ֲ����ڼ����У����Եȼ����ӡ�", vbCritical: Exit Sub
    Me.Caption = "���ݻ� " & App.Major & "." & App.Minor & "." & App.Revision & " - ��������"
    cmdAction.Caption = "����"
    OperateType = "changescene"
    If Handbook.Exists("Scenes") Then
        HandbookKey = "Scenes"
    ElseIf Handbook.Exists("Scene") Then
        HandbookKey = "Scene"
    Else
        HandbookKey = InputBox("�ֲ᲻֧�֣����ֶ����볡������")
    End If
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
    Me.Show
End Sub

Public Sub TriggerAccount()
    Me.Caption = "���ݻ� " & App.Major & "." & App.Minor & "." & App.Revision & "  - ע���ɾ���˺�"
    cmdAction.Caption = "ִ��"
    OperateType = "account"
    lst.AddItem "ע���˺�"
    lst.AddItem "ɾ���˺�"
    Label1.Visible = True
    Label1.Caption = "����:"
    Text1.Visible = True
    Label2.Visible = True
    Label2.Caption = "(������д����)"
    Text2.Visible = False
    Label3.Visible = True
    Label3.Caption = "UID:"
    Text3.Visible = True
    Text2.Text = "10001"
    Me.Show
End Sub

Public Sub TriggerStats()
    Me.Caption = "���ݻ� " & App.Major & "." & App.Minor & "." & App.Revision & "  - ��ɫ��ֵ�޸�"
    cmdAction.Caption = "ִ��"
    OperateType = "setstats"
    lst.AddItem "����ֵ (ID:hp)"
    lst.AddItem "������ (ID:def)"
    lst.AddItem "������ (ID:atk)"
    lst.AddItem "Ԫ�ؾ�ͨ (ID:em)"
    lst.AddItem "Ԫ�س���Ч�� (ID:er)"
    lst.AddItem "������ (ID:crate)"
    lst.AddItem "�����˺� (ID:cdmg)"
    lst.AddItem "��Ԫ���˺��ӳ� (ID:epyro)"
    lst.AddItem "��Ԫ���˺��ӳ� (ID:ecyro)"
    lst.AddItem "ˮԪ���˺��ӳ� (ID:ehydro)"
    lst.AddItem "��Ԫ���˺��ӳ� (ID:egeo)"
    lst.AddItem "��Ԫ���˺��ӳ� (ID:edend)"
    lst.AddItem "��Ԫ���˺��ӳ� (ID:eelec)"
    lst.AddItem "�����˺��ӳ� (ID:ephys)"
    Label1.Visible = True
    Label1.Caption = "��ֵ:"
    Text1.Visible = True
    Me.Show
End Sub

Public Sub TriggerDrop()
    If HandbookLoaded <> True Then MsgBox "�ֲ����ڼ����У����Եȼ����ӡ�", vbCritical: Exit Sub
    Me.Caption = "���ݻ� " & App.Major & "." & App.Minor & "." & App.Revision & "  - ���ɵ�����"
    cmdAction.Caption = "����"
    OperateType = "drop"
    Combo1.Clear
    Combo1.AddItem "����"
    Combo1.AddItem "����"
    Combo1.AddItem "ʥ����"
    Combo1.Text = "����"
    HandbookKey = "����"
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
    Combo1.Visible = True
    Label1.Visible = True
    Label1.Caption = "����:"
    Text1.Visible = False
    Label2.Visible = True
    Label2.Caption = "����:"
    Text2.Visible = True
    Me.Show
End Sub
Public Sub TriggerGive()
    If HandbookLoaded <> True Then MsgBox "�ֲ����ڼ����У����Եȼ����ӡ�", vbCritical: Exit Sub
    Me.Caption = "���ݻ� " & App.Major & "." & App.Minor & "." & App.Revision & "  - ������Ʒ"
    cmdAction.Caption = "����"
    OperateType = "give"
    Combo1.Clear
    Combo1.AddItem "����"
    Combo1.AddItem "����"
    Combo1.AddItem "ʥ����"
    Combo1.Text = "����"
    HandbookKey = "����"
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
    Combo1.Visible = True
    Label1.Visible = True
    Label1.Caption = "����:"
    Text1.Visible = False
    Label2.Visible = True
    Label2.Caption = "����:"
    Text2.Visible = True
    Label3.Visible = True
    Label3.Caption = "[�ȼ�]:"
    Text3.Visible = True
    Label4.Visible = True
    Label4.Caption = "(�ȼ�Ϊѡ��)"
    Me.Show
End Sub


Public Sub Search()
    lst.Clear
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        If InStr(Handbook(HandbookKey)(EachItem), txtSearch.Text) Then lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
End Sub

Private Sub cmdAction_Click()
If lst.Text = "" Then Exit Sub
Dim ItemID As String
Select Case OperateType
Case "spawn"
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    MsgBox "!spawn " & ItemID & " " & Text1.Text & vbCrLf & "�Ѿ����Ƶ������塣"
    Clipboard.SetText "!spawn " & ItemID & " " & Text1.Text
Case "give"
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    If Text2.Text = "" Then Text2.Text = "1"
    frmMain.Server.SendLine "!give " & frmMain.txtUID.Text & " " & ItemID & " " & Text2.Text & " " & Text3.Text
    MsgBox "�� " & Text2.Text & " �� " & lst.Text & " ���� UID Ϊ " & frmMain.txtUID.Text & " ����ҡ�"
Case "givechar"
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    If Text1.Text = "" Then Text1.Text = "1"
    frmMain.Server.SendLine "!givechar " & frmMain.txtUID.Text & " " & ItemID & " " & Text1.Text
    MsgBox "�� " & lst.Text & " ���� UID Ϊ " & frmMain.txtUID.Text & " ����ҡ�"
Case "account"
    If Text1.Text = "" Then Exit Sub
    If lst.Text = "ע���˺�" Then
        If Text2.Text = "" Then Exit Sub
        If CInt(Text2.Text) <> Text2.Text Then Exit Sub
        frmMain.Server.SendLine "!account create " & Text1.Text & " " & Text2.Text
        frmMain.txtUID.Text = Text2.Text
        WriteIni "GCUI", "PlayerUID", Text2.Text, App.Path & "\Config.ini"
        MsgBox "ע�����˺ţ������ַΪ " & Text1.Text & "��UID Ϊ " & Text2.Text & "��"
    ElseIf lst.Text = "ɾ���˺�" Then
        frmMain.Server.SendLine "!account delete " & Text1.Text & " " & Text2.Text
        frmMain.txtUID.Text = Text2.Text
        WriteIni "GCUI", "PlayerUID", Text2.Text, App.Path & "\Config.ini"
        MsgBox "ɾ�����˺ţ������ַΪ " & Text1.Text & "��"
    Else
        Exit Sub
    End If
Case "changescene"
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    MsgBox "!changescene " & ItemID & vbCrLf & "�Ѿ����Ƶ������塣"
    Clipboard.SetText "!changescene " & ItemID
Case "setstats"
    If Text1.Text = "" Then Exit Sub
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    MsgBox "!setstats " & ItemID & " " & Text1.Text & vbCrLf & "�Ѿ����Ƶ������塣"
    Clipboard.SetText "!setstats " & ItemID & " " & Text1.Text
Case "drop"
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    MsgBox "!drop " & ItemID & " " & Text1.Text & vbCrLf & "�Ѿ����Ƶ������塣"
    Clipboard.SetText "!drop " & ItemID & " " & Text1.Text
Case Else
    MsgBox "������δ֪�������������ڽ�ͼ�������ߡ�" & vbCrLf & vbCrLf & "OperateType: " & OperateType & vbCrLf & Me.Caption
    End
End Select
Unload Me
End Sub

Private Sub cmdSearch_Click()
    Search
End Sub


Private Sub Combo1_Click()
    If Handbook.Exists(Combo1.Text) Then
        HandbookKey = Combo1.Text
    Else
        HandbookKey = InputBox("�ֲ᲻֧�֣����ֶ�����й�����")
        Combo1.Text = HandbookKey
    End If
    txtSearch.Text = ""
    lst.Clear
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtSearch.Text = "" Then
            lst.Clear
            Dim EachItem As Variant
            For Each EachItem In Handbook(HandbookKey).Keys
                lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
            Next
        Else
            Search
        End If
        End If
    End Sub
