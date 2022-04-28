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
      Name            =   "微软雅黑"
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
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   6720
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "搜"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "生成"
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
         Name            =   "微软雅黑"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   720
      Width           =   975
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
    If HandbookLoaded <> True Then MsgBox "手册正在加载中，请稍等几秒钟。", vbCritical: Exit Sub
    Me.Caption = "Grasscutter UI - 生成敌怪"
    cmdAction.Caption = "生成"
    OperateType = "spawn"
    If Handbook.Exists("Monster") Then
        HandbookKey = "Monster"
    ElseIf Handbook.Exists("讨伐列表") Then
        HandbookKey = "讨伐列表"
    Else
        HandbookKey = InputBox("手册不支持，请手动输入敌怪类名")
    End If
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
    Label1.Visible = True
    Label1.Caption = "等级:"
    Text1.Visible = True
    Me.Show
End Sub


Public Sub TriggerSearch()
    lst.Clear
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        If InStr(Handbook(HandbookKey)(EachItem), txtSearch.Text) Then lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
End Sub

Private Sub cmdAction_Click()
Dim ItemID As String
ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
Select Case OperateType
Case "spawn"
MsgBox "/spawn" & ItemID & " " & Text1.Text
End Select
End Sub

Private Sub cmdSearch_Click()
    TriggerSearch
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
            TriggerSearch
        End If
        End If
    End Sub
