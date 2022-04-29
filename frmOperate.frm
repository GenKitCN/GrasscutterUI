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
      Height          =   345
      Left            =   6720
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   345
      Left            =   6720
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   6720
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   6720
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   6720
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   6720
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
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
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   720
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
    Label1.Caption = "�ȼ�:"
    Text1.Visible = True
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
ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
Select Case OperateType
Case "spawn"
    MsgBox "!spawn " & ItemID & " " & Text1.Text & vbCrLf & "�Ѿ����Ƶ������塣"
    Clipboard.SetText "!spawn " & ItemID & " " & Text1.Text
Case "give"
    If Text2.Text = "" Then Text2.Text = "1"
    frmMain.Server.SendLine "!give " & frmMain.txtUID.Text & " " & ItemID & " " & Text2.Text & " " & Text3.Text
    MsgBox "�� " & Text2.Text & " �� " & lst.Text & " ���� UID Ϊ " & frmMain.txtUID.Text & " ����ҡ�"
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
